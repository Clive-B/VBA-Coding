Attribute VB_Name = "CSTSingleFileProcessing"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Sub ProcessSingleCSTFromForm()
    Dim frm As CSTUserForm

    Set frm = New CSTUserForm
    frm.Show

    If frm.FormSubmitted = False Then
        MsgBox "Process cancelled.", vbInformation, "Single CST Processing"
        Unload frm
        Exit Sub
    End If

    ProcessSingleCSTFiles frm.CSTExcelName, frm.SelectedCSTFilePath, frm.StorageFolderPath
    Unload frm
End Sub

Public Sub ProcessSingleCSTFromSubmittedForm(Optional ByVal frm As Object = Nothing)
    If frm Is Nothing Then
        MsgBox "Call this helper from the UserForm as: ProcessSingleCSTFromSubmittedForm Me", _
               vbExclamation, "Single CST Processing"
        Exit Sub
    End If

    ProcessSingleCSTFiles frm.cstExcelName, frm.selectedCSTFilePath, frm.storageFolderPath
End Sub

Public Sub ProcessSingleCSTFiles(ByVal cstExcelName As String, ByVal selectedCSTFilePath As String, ByVal storageFolderPath As String)
    Dim fso As Object
    Dim relatedFiles As Collection
    Dim mocFiles As Collection
    Dim mtcFiles As Collection
    Dim filePath As Variant
    Dim workbookPath As String
    Dim openedWorkbook As Workbook
    Dim trends As String
    Dim operatorName As String
    Dim selectedCallType As String
    Dim finalMtcPath As String
    Dim isFinalFile As Boolean
    Dim engineConfigured As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo FatalError

    Set fso = CreateObject("Scripting.FileSystemObject")
    cstExcelName = Trim$(cstExcelName)
    selectedCSTFilePath = Trim$(selectedCSTFilePath)
    storageFolderPath = Trim$(storageFolderPath)

    If cstExcelName = "" Then
        MsgBox "Please enter the CST Excel Name.", vbExclamation, "Single CST Processing"
        Exit Sub
    End If

    If Not IsValidNmfPath(selectedCSTFilePath, fso) Then
        MsgBox "Please select a valid .nmf or .nmfs CST log file.", vbExclamation, "Single CST Processing"
        Exit Sub
    End If

    If storageFolderPath = "" Or Not fso.FolderExists(storageFolderPath) Then
        MsgBox "Please select a valid output folder.", vbExclamation, "Single CST Processing"
        Exit Sub
    End If

    selectedCallType = GetSingleCstCallType(selectedCSTFilePath, fso)
    If selectedCallType <> "MOC" And selectedCallType <> "MTC" Then
        MsgBox "The selected log was not identified as MOC or MTC.", vbExclamation, "Single CST Processing"
        Exit Sub
    End If

    Set relatedFiles = GetRelatedSingleCstLogs(selectedCSTFilePath, fso)
    Set mocFiles = New Collection
    Set mtcFiles = New Collection

    For Each filePath In relatedFiles
        Select Case GetSingleCstCallType(CStr(filePath), fso)
            Case "MOC"
                AddPathSorted mocFiles, CStr(filePath), fso
            Case "MTC"
                AddPathSorted mtcFiles, CStr(filePath), fso
        End Select
    Next filePath

    If mocFiles.Count = 0 Or mtcFiles.Count = 0 Then
        MsgBox "The selected CST set must include both MOC and MTC logs." & vbCrLf & _
               "MOC files found: " & mocFiles.Count & vbCrLf & _
               "MTC files found: " & mtcFiles.Count, vbExclamation, "Single CST Processing"
        Exit Sub
    End If

    trends = Sheets("Option").Range("AO1").Value
    operatorName = DetectSingleCstOperator(selectedCSTFilePath, fso)
    If operatorName = "" Then operatorName = DetectOperatorFromPath(selectedCSTFilePath)
    If operatorName = "" Then
        MsgBox "Unable to determine the network/operator from the selected log.", vbExclamation, "Single CST Processing"
        Exit Sub
    End If

    workbookPath = GetDesktopPath() & "QoS Automation\Templates\Voice Template.xlsm"
    If Dir(workbookPath) = "" Then
        MsgBox "Voice Template workbook was not found:" & vbCrLf & workbookPath, vbExclamation, "Single CST Processing"
        Exit Sub
    End If

    CloseOtherCstWorkbooks
    Sleep 1000
    Set openedWorkbook = Workbooks.Open(workbookPath)

    Application.Run "'" & openedWorkbook.FullName & "'!ConfigureSingleCSTOutput", storageFolderPath, cstExcelName, True
    engineConfigured = True

    For Each filePath In mocFiles
        Application.Run "'" & openedWorkbook.FullName & "'!CSTCOVMOS", CStr(filePath), operatorName, trends, False, False, True
    Next filePath

    finalMtcPath = CStr(mtcFiles(mtcFiles.Count))
    For Each filePath In mtcFiles
        isFinalFile = (StrComp(CStr(filePath), finalMtcPath, vbTextCompare) = 0)
        Application.Run "'" & openedWorkbook.FullName & "'!CSTCOVMOS", CStr(filePath), operatorName, trends, isFinalFile, isFinalFile, True
    Next filePath

    engineConfigured = False
    MsgBox "Single CST processing completed.", vbInformation, "Single CST Processing"
    Exit Sub

FatalError:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If engineConfigured And Not openedWorkbook Is Nothing Then
        Application.Run "'" & openedWorkbook.FullName & "'!ClearSingleCSTOutput"
    End If
    MsgBox "Single CST processing failed:" & vbCrLf & _
           "Error " & errNumber & ": " & errDescription, vbExclamation, "Single CST Processing"
End Sub

Private Function GetRelatedSingleCstLogs(ByVal selectedFilePath As String, ByVal fso As Object) As Collection
    Dim selectedFolder As Object
    Dim fileItem As Object
    Dim selectedKey As String
    Dim candidateKey As String
    Dim related As Collection
    Dim fallbackRelated As Collection
    Dim selectedTownDateKey As String

    Set related = New Collection
    Set fallbackRelated = New Collection
    Set selectedFolder = fso.GetFolder(fso.GetParentFolderName(selectedFilePath))

    selectedKey = BuildSingleCstSessionKey(fso.GetBaseName(selectedFilePath))
    selectedTownDateKey = BuildTownDateKey(fso.GetFileName(selectedFilePath))

    For Each fileItem In selectedFolder.Files
        If IsValidNmfPath(fileItem.Path, fso) Then
            candidateKey = BuildSingleCstSessionKey(fso.GetBaseName(fileItem.Path))
            If StrComp(candidateKey, selectedKey, vbTextCompare) = 0 Then
                AddPathSorted related, fileItem.Path, fso
            ElseIf selectedTownDateKey <> "" And StrComp(BuildTownDateKey(fileItem.Name), selectedTownDateKey, vbTextCompare) = 0 Then
                AddPathSorted fallbackRelated, fileItem.Path, fso
            End If
        End If
    Next fileItem

    If related.Count > 1 Then
        Set GetRelatedSingleCstLogs = related
    ElseIf fallbackRelated.Count > 0 Then
        Set GetRelatedSingleCstLogs = fallbackRelated
    Else
        related.Add selectedFilePath
        Set GetRelatedSingleCstLogs = related
    End If
End Function

Private Sub AddPathSorted(ByRef target As Collection, ByVal filePath As String, ByVal fso As Object)
    Dim indexNo As Long
    Dim fileNameText As String
    Dim currentName As String

    fileNameText = UCase$(fso.GetFileName(filePath))
    For indexNo = 1 To target.Count
        currentName = UCase$(fso.GetFileName(CStr(target(indexNo))))
        If StrComp(fileNameText, currentName, vbTextCompare) < 0 Then
            target.Add filePath, Before:=indexNo
            Exit Sub
        End If
    Next indexNo

    target.Add filePath
End Sub

Private Function BuildSingleCstSessionKey(ByVal baseName As String) As String
    Dim keyText As String
    Dim lastDot As Long

    keyText = NormalizeSpaces(baseName)
    lastDot = InStrRev(keyText, ".")
    If lastDot > 0 Then
        If IsNumeric(Mid$(keyText, lastDot + 1)) Then
            keyText = Left$(keyText, lastDot - 1)
        End If
    End If

    BuildSingleCstSessionKey = UCase$(keyText)
End Function

Private Function BuildTownDateKey(ByVal fileName As String) As String
    Dim townText As String
    Dim dateText As String

    dateText = ExtractFirstToken(fileName)
    townText = ExtractThirdToken(fileName)
    If dateText <> "" And townText <> "" Then
        BuildTownDateKey = UCase$(dateText & "_" & townText)
    End If
End Function

Private Function GetSingleCstCallType(ByVal filePath As String, ByVal fso As Object) As String
    Dim textStream As Object
    Dim lineText As String
    Dim fields() As String
    Dim parts() As String

    On Error GoTo CleanUp
    Set textStream = fso.OpenTextFile(filePath)
    Do Until textStream.AtEndOfStream
        lineText = textStream.ReadLine
        fields = Split(lineText, ",")
        If UBound(fields) >= 3 Then
            If fields(0) = "#DL" Then
                parts = Split(fields(3), " ")
                If UBound(parts) >= 1 Then
                    GetSingleCstCallType = UCase$(Replace(parts(1), """", ""))
                    Exit Do
                End If
            End If
        End If
    Loop

CleanUp:
    On Error Resume Next
    If Not textStream Is Nothing Then textStream.Close
End Function

Private Function DetectSingleCstOperator(ByVal filePath As String, ByVal fso As Object) As String
    Dim textStream As Object
    Dim lineText As String
    Dim fields() As String
    Dim parts() As String

    On Error GoTo CleanUp
    Set textStream = fso.OpenTextFile(filePath)
    Do Until textStream.AtEndOfStream
        lineText = textStream.ReadLine
        fields = Split(lineText, ",")
        If UBound(fields) >= 3 Then
            If fields(0) = "#DL" Then
                parts = Split(fields(3), " ")
                If UBound(parts) >= 0 Then
                    DetectSingleCstOperator = UCase$(Replace(parts(0), """", ""))
                    Exit Do
                End If
            End If
        End If
    Loop

CleanUp:
    On Error Resume Next
    If Not textStream Is Nothing Then textStream.Close
End Function

Private Function DetectOperatorFromPath(ByVal filePath As String) As String
    Dim pathText As String

    pathText = UCase$(filePath)
    If InStr(1, pathText, "\MTN\", vbTextCompare) > 0 Then
        DetectOperatorFromPath = "MTN"
    ElseIf InStr(1, pathText, "\TELECEL\", vbTextCompare) > 0 Then
        DetectOperatorFromPath = "TELECEL"
    ElseIf InStr(1, pathText, "\AT\", vbTextCompare) > 0 Then
        DetectOperatorFromPath = "AT"
    ElseIf InStr(1, pathText, "\AIRTELTIGO\", vbTextCompare) > 0 Then
        DetectOperatorFromPath = "AIRTELTIGO"
    End If
End Function

Private Function IsValidNmfPath(ByVal filePath As String, ByVal fso As Object) As Boolean
    Dim extText As String

    If filePath = "" Then Exit Function
    If Not fso.FileExists(filePath) Then Exit Function

    extText = UCase$(fso.GetExtensionName(filePath))
    IsValidNmfPath = (extText = "NMF" Or extText = "NMFS")
End Function

Private Function ExtractFirstToken(ByVal textValue As String) As String
    Dim parts() As String

    parts = Split(NormalizeSpaces(textValue), " ")
    If UBound(parts) >= 0 Then ExtractFirstToken = parts(0)
End Function

Private Function ExtractThirdToken(ByVal textValue As String) As String
    Dim parts() As String

    parts = Split(NormalizeSpaces(textValue), " ")
    If UBound(parts) >= 2 Then ExtractThirdToken = parts(2)
End Function

Private Function NormalizeSpaces(ByVal textValue As String) As String
    Dim result As String

    result = Trim$(textValue)
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
    NormalizeSpaces = result
End Function

Private Function GetDesktopPath() As String
    Dim shellObj As Object

    Set shellObj = CreateObject("WScript.Shell")
    GetDesktopPath = shellObj.SpecialFolders("Desktop") & "\"
End Function

Private Sub CloseOtherCstWorkbooks()
    Dim wb As Workbook

    For Each wb In Workbooks
        If wb.Name <> "QoS.xlsm" And wb.Name <> "Voice Template.xlsm" Then
            wb.Close SaveChanges:=False
            Exit For
        End If
    Next wb
End Sub
