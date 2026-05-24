' ******************************************************************************
' SolidWorks 2025 深度递归属性工具 (全自动化统计版)
' 优化：1. 增加零件/装配体分类计数 2. 增强链接丢失容错 3. 格式错误分类统计
' ******************************************************************************

Option Explicit

Dim swApp As SldWorks.SldWorks
Dim errorLog As String

' 定义全局计数器
Dim partCount As Long ' 正常零件数
Dim asmCount As Long ' 正常装配体数
Dim errPartCount As Long ' 格式错误零件数
Dim errAsmCount As Long ' 格式错误装配体数

Sub main()
 Dim ShellApp As Object
 Dim Folder As Object
 Dim folderPath As String

 Set swApp = Application.SldWorks

 ' 初始化数据
 errorLog = ""
 partCount = 0: asmCount = 0
 errPartCount = 0: errAsmCount = 0

 ' 1. 选择主目录
 Set ShellApp = CreateObject("Shell.Application")
 Set Folder = ShellApp.BrowseForFolder(0, "请选择需要处理的工程文件夹", 0)

 If Folder Is Nothing Then Exit Sub
 folderPath = Folder.Items.Item.path

 ' 2. 深度静默与链接容错设置 (使用数字常量避免环境引用错误)
 On Error Resume Next
 swApp.SetUserPreferenceToggle 110, True ' swInputMessages_DontShowAgain
 swApp.SetUserPreferenceToggle 147, False ' swShowErrorsEverySave
 swApp.SetUserPreferenceToggle 152, False ' 不搜索外部参考
 swApp.SetUserPreferenceToggle 153, False ' 不搜索参考列表
 On Error GoTo 0

 ' 3. 执行递归
 RecursiveFolder folderPath

 ' 4. 恢复标准设置
 On Error Resume Next
 swApp.SetUserPreferenceToggle 110, False
 swApp.SetUserPreferenceToggle 147, True
 swApp.SetUserPreferenceToggle 152, True
 swApp.SetUserPreferenceToggle 153, True
 On Error GoTo 0

 ' 5. 汇总报告
 Dim summaryMsg As String
 summaryMsg = "属性更新任务处理完成！" & vbCrLf & vbCrLf & _
 "成功处理数量：" & vbCrLf & _
 " - 零件 (.SLDPRT): " & partCount & " 件" & vbCrLf & _
 " - 装配体 (.SLDASM): " & asmCount & " 件" & vbCrLf & vbCrLf

 If (errPartCount + errAsmCount) > 0 Then
 summaryMsg = summaryMsg & "格式不符(已填空模板)：" & vbCrLf & _
 " - 零件: " & errPartCount & " 件" & vbCrLf & _
 " - 装配体: " & errAsmCount & " 件" & vbCrLf & vbCrLf & _
 "详情清单：" & vbCrLf & errorLog
 MsgBox summaryMsg, vbExclamation, "处理报告"
 Else
 MsgBox summaryMsg & "所有文件格式均符合规范！", vbInformation, "处理报告"
 End If
End Sub

' --- 递归核心函数 ---
Sub RecursiveFolder(ByVal path As String)
 Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
 Dim parentFolder As Object: Set parentFolder = fso.GetFolder(path)
 Dim subFolder As Object, file As Object

 For Each file In parentFolder.Files
 Dim ext As String: ext = LCase(fso.GetExtensionName(file.path))
 If ext = "sldprt" Or ext = "sldasm" Then
 ProcessSingleFile file.path, ext
 End If
 Next file

 For Each subFolder In parentFolder.SubFolders
 RecursiveFolder subFolder.path
 Next subFolder
End Sub

' --- 单个文件处理逻辑 ---
Sub ProcessSingleFile(filePath As String, ext As String)
 Dim swModel As SldWorks.ModelDoc2
 Dim errors As Long, warnings As Long
 Dim docType As Long

 docType = IIf(ext = "sldprt", 1, 2)

 ' 静默打开
 Set swModel = swApp.OpenDoc6(filePath, docType, 1, "", errors, warnings)

 If Not swModel Is Nothing Then
 ' 更新属性
 UpdateProperties swModel, filePath

 ' 保存并关闭
 swModel.ForceRebuild3 True
 swModel.Save3 1, errors, warnings
 swApp.CloseDoc swModel.GetTitle
 End If
End Sub

' --- 核心属性更新逻辑 (带计数器) ---
Sub UpdateProperties(model As SldWorks.ModelDoc2, filePath As String)
 Dim swCustPropMgr As SldWorks.CustomPropertyManager
 Dim rawFileName As String, extStr As String
 Dim fileNameParts As Variant
 Dim vPropNames As Variant
 Dim i As Long
 Dim valOut As String, resValOut As String
 Dim oldDate As String, isFormatError As Boolean
 Dim docType As Long: docType = model.GetType

 ' 提取文件名
 rawFileName = Mid(filePath, InStrRev(filePath, "\") + 1)
 If InStrRev(rawFileName, ".") > 0 Then
 extStr = Mid(rawFileName, InStrRev(rawFileName, "."))
 rawFileName = Left(rawFileName, InStrRev(rawFileName, ".") - 1)
 End If

 ' 解析格式
 fileNameParts = Split(rawFileName, "+")
 isFormatError = (UBound(fileNameParts) < 2)

 ' --- 计数统计开始 ---
 If isFormatError Then
 errorLog = errorLog & "- " & rawFileName & vbCrLf
 If docType = 1 Then errPartCount = errPartCount + 1 Else errAsmCount = errAsmCount + 1
 Else
 If docType = 1 Then partCount = partCount + 1 Else asmCount = asmCount + 1
 End If
 ' --- 计数统计结束 ---

 Set swCustPropMgr = model.Extension.CustomPropertyManager("")

 ' 保留原始设计日期
 swCustPropMgr.Get6 "设计日期", False, valOut, resValOut, False, False
 oldDate = IIf(Trim(resValOut) <> "", Trim(resValOut), CStr(Date))

 ' 清空现有属性
 vPropNames = swCustPropMgr.GetNames
 If Not IsEmpty(vPropNames) Then
 For i = 0 To UBound(vPropNames)
 swCustPropMgr.Delete vPropNames(i)
 Next i
 End If

 ' 属性写入
 swCustPropMgr.Add3 "名称", 30, IIf(isFormatError, "", Trim(fileNameParts(0))), 1
 swCustPropMgr.Add3 "零部件号", 30, IIf(isFormatError, "", Trim(fileNameParts(1))), 1
 swCustPropMgr.Add3 "图样代号", 30, IIf(isFormatError, "", Trim(fileNameParts(1))), 1

 ' 材料逻辑
 If docType = 1 And Not isFormatError Then
 swCustPropMgr.Add3 "材料", 30, Chr(34) & "SW-Material@" & rawFileName & extStr & Chr(34), 1
 Else
 swCustPropMgr.Add3 "材料", 30, "", 1
 End If

 swCustPropMgr.Add3 "质量", 30, Chr(34) & "SW-Mass@" & rawFileName & extStr & Chr(34), 1
 swCustPropMgr.Add3 "单位", 30, "Pcs", 1
 swCustPropMgr.Add3 "设计", 30, "Abel", 1
 swCustPropMgr.Add3 "设计日期", 30, oldDate, 1

 ' 占位与版本
 swCustPropMgr.Add3 "审核", 30, "", 1
 swCustPropMgr.Add3 "审核日期", 30, "", 1
 swCustPropMgr.Add3 "工艺", 30, "", 1
 swCustPropMgr.Add3 "工艺日期", 30, "", 1
 swCustPropMgr.Add3 "批准", 30, "", 1
 swCustPropMgr.Add3 "批准日期", 30, "", 1
 swCustPropMgr.Add3 "版本号", 30, IIf(isFormatError, "", Trim(fileNameParts(2))), 1
 swCustPropMgr.Add3 "备注", 30, "", 1
End Sub
