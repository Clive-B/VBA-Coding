Attribute VB_Name = "FourGDataSingleFileProcessing"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const SINGLE_4G_DATA_FORM_NAME As String = "ThreeGDataUserForm"

Public Sub OpenSingle4GDataForm()
    ProcessSingle4GDataFromForm
End Sub

Public Sub ProcessSingle4GDataFromForm()
    Dim frm As Object

    On Error GoTo FormLoadError
    Set frm = UserForms.Add(SINGLE_4G_DATA_FORM_NAME)
    On Error GoTo 0

    frm.Show

    If ReadFormBoolean(frm, "FormSubmitted") = False Then
        MsgBox "Process cancelled.", vbInformation, "Single 4G Data Processing"
        Unload frm
        Exit Sub
    End If

    ProcessSingle4GDataFiles ReadFirstFormText(frm, "DataExcelName", "dataExcelName", "fourGDataExcelName", "threeGDataExcelName", "cstExcelName"), _
                             ReadFirstFormText(frm, "SelectedDataFilePath", "selectedDataFilePath", "selected4GDataFilePath", "selected3GDataFilePath", "selectedCSTFilePath"), _
                             ReadFirstFormText(frm, "StorageFolderPath", "storageFolderPath", "outputFolderPath", "selectedOutputFolderPath")
    Unload frm
    Exit Sub

FormLoadError:
    MsgBox "Unable to open " & SINGLE_4G_DATA_FORM_NAME & "." & vbCrLf & _
           "Create that UserForm or call ProcessSingle4GDataFiles directly.", _
           vbExclamation, "Single 4G Data Processing"
End Sub

Public Sub ProcessSingle4GDataFromSubmittedForm(Optional ByVal frm As Object = Nothing)
    If frm Is Nothing Then
        MsgBox "Call this helper from the UserForm as: ProcessSingle4GDataFromSubmittedForm Me", _
               vbExclamation, "Single 4G Data Processing"
        Exit Sub
    End If

    ProcessSingle4GDataFiles ReadFirstFormText(frm, "DataExcelName", "dataExcelName", "fourGDataExcelName", "threeGDataExcelName", "cstExcelName"), _
                             ReadFirstFormText(frm, "SelectedDataFilePath", "selectedDataFilePath", "selected4GDataFilePath", "selected3GDataFilePath", "selectedCSTFilePath"), _
                             ReadFirstFormText(frm, "StorageFolderPath", "storageFolderPath", "outputFolderPath", "selectedOutputFolderPath")
End Sub

Public Sub ProcessSingle4GDataFiles(ByVal dataExcelName As String, ByVal selectedDataFilePath As String, ByVal storageFolderPath As String)
    Dim fso As Object
    Dim relatedFiles As Collection
    Dim dataFiles As Collection
    Dim filePath As Variant
    Dim workbookPath As String
    Dim openedWorkbook As Workbook
    Dim trends As String
    Dim selectedTech As String
    Dim finalFilePath As String
    Dim isFinalFile As Boolean
    Dim fileNameText As String
    Dim townNameText As String
    Dim locationName As String
    Dim fileDate As String
    Dim dupBreak As String
    Dim engineConfigured As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo FatalError

    Set fso = CreateObject("Scripting.FileSystemObject")
    dataExcelName = Trim$(dataExcelName)
    selectedDataFilePath = Trim$(selectedDataFilePath)
    storageFolderPath = Trim$(storageFolderPath)

    If dataExcelName = "" Then
        MsgBox "Please enter the 4G Data Excel Name.", vbExclamation, "Single 4G Data Processing"
        Exit Sub
    End If

    If Not IsValidNmfPath(selectedDataFilePath, fso) Then
        MsgBox "Please select a valid .nmf or .nmfs 4G Data log file.", vbExclamation, "Single 4G Data Processing"
        Exit Sub
    End If

    If storageFolderPath = "" Or Not fso.FolderExists(storageFolderPath) Then
        MsgBox "Please select a valid output folder.", vbExclamation, "Single 4G Data Processing"
        Exit Sub
    End If

    selectedTech = GetSingleDataTech(selectedDataFilePath, fso)
    If selectedTech <> "4G" Then
        MsgBox "The selected log was identified as '" & selectedTech & "', not 4G.", _
               vbExclamation, "Single 4G Data Processing"
        Exit Sub
    End If

    Set relatedFiles = GetRelatedSingle4GDataLogs(selectedDataFilePath, fso)
    Set dataFiles = New Collection

    For Each filePath In relatedFiles
        If GetSingleDataTech(CStr(filePath), fso) = "4G" Then
            AddPathSorted dataFiles, CStr(filePath), fso
        End If
    Next filePath

    If dataFiles.Count = 0 Then
        MsgBox "No related 4G Data logs were found for the selected file.", _
               vbExclamation, "Single 4G Data Processing"
        Exit Sub
    End If

    trends = Sheets("Option").Range("AO1").Value
    workbookPath = GetDesktopPath() & "QoS Automation\Templates\4G Data Template.xlsm"
    If Dir(workbookPath) = "" Then
        MsgBox "4G Data Template workbook was not found:" & vbCrLf & workbookPath, _
               vbExclamation, "Single 4G Data Processing"
        Exit Sub
    End If

    CloseOther4GDataWorkbooks
    Sleep 1000
    Set openedWorkbook = Workbooks.Open(workbookPath)
    openedWorkbook.Activate

    Application.Run "'" & openedWorkbook.FullName & "'!FourGDataMacro.ConfigureSingle4GDataOutput", storageFolderPath, dataExcelName, True
    engineConfigured = True

    finalFilePath = CStr(dataFiles(dataFiles.Count))
    For Each filePath In dataFiles
        isFinalFile = (StrComp(CStr(filePath), finalFilePath, vbTextCompare) = 0)
        fileNameText = fso.GetFileName(CStr(filePath))
        townNameText = ExtractTownFromDataFile(fileNameText)
        locationName = ExtractLocationFromDataFile(fileNameText)
        fileDate = ExtractDateFromDataFile(fileNameText)
        dupBreak = fileDate & " " & townNameText & " " & locationName

        Application.Run "'" & openedWorkbook.FullName & "'!FourGDataMacro.FourGDataSingle", _
                        CStr(filePath), dupBreak, fileNameText, locationName, fileDate, trends, False, _
                        True, True, isFinalFile, isFinalFile, 1, locationName, locationName, townNameText
    Next filePath

    engineConfigured = False
    MsgBox "Single 4G Data processing completed.", vbInformation, "Single 4G Data Processing"
    Exit Sub

FatalError:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If engineConfigured And Not openedWorkbook Is Nothing Then
        Application.Run "'" & openedWorkbook.FullName & "'!FourGDataMacro.ClearSingle4GDataOutput"
    End If
    MsgBox "Single 4G Data processing failed:" & vbCrLf & _
           "Error " & errNumber & ": " & errDescription, vbExclamation, "Single 4G Data Processing"
End Sub

Private Function GetRelatedSingle4GDataLogs(ByVal selectedFilePath As String, ByVal fso As Object) As Collection
    Dim selectedFolder As Object
    Dim fileItem As Object
    Dim selectedKey As String
    Dim candidateKey As String
    Dim related As Collection

    Set related = New Collection
    Set selectedFolder = fso.GetFolder(fso.GetParentFolderName(selectedFilePath))
    selectedKey = BuildSingleDataKey(fso.GetFileName(selectedFilePath))

    For Each fileItem In selectedFolder.Files
        If IsValidNmfPath(fileItem.Path, fso) Then
            candidateKey = BuildSingleDataKey(fileItem.Name)
            If selectedKey <> "" And StrComp(candidateKey, selectedKey, vbTextCompare) = 0 Then
                AddPathSorted related, fileItem.Path, fso
            End If
        End If
    Next fileItem

    If related.Count = 0 Then related.Add selectedFilePath
    Set GetRelatedSingle4GDataLogs = related
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

Private Function BuildSingleDataKey(ByVal fileNameText As String) As String
    Dim fileDate As String
    Dim fileDay As String

    fileDate = ExtractDateFromDataFile(fileNameText)
    fileDay = ExtractDayFromDataFile(fileNameText)

    If fileDate = "" Or fileDay = "" Then Exit Function

    BuildSingleDataKey = UCase$(fileDate & "|" & fileDay)
End Function

Private Function GetSingleDataTech(ByVal filePath As String, ByVal fso As Object) As String
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
                    GetSingleDataTech = UCase$(Replace(parts(1), """", ""))
                    Exit Do
                End If
            End If
        End If
    Loop

CleanUp:
    On Error Resume Next
    If Not textStream Is Nothing Then textStream.Close
End Function

Private Function IsValidNmfPath(ByVal filePath As String, ByVal fso As Object) As Boolean
    Dim extText As String

    If filePath = "" Then Exit Function
    If Not fso.FileExists(filePath) Then Exit Function

    extText = UCase$(fso.GetExtensionName(filePath))
    IsValidNmfPath = (extText = "NMF" Or extText = "NMFS")
End Function

Private Function ExtractTownFromDataFile(ByVal fileNameText As String) As String
    Dim parts() As String

    parts = Split(NormalizeSpaces(fileNameText), " ")
    If UBound(parts) >= 2 Then ExtractTownFromDataFile = parts(2)
End Function

Private Function ExtractLocationFromDataFile(ByVal fileNameText As String) As String
    Dim parts() As String
    Dim indexNo As Long

    parts = Split(NormalizeSpaces(fileNameText), " ")
    For indexNo = 3 To UBound(parts)
        If UCase$(parts(indexNo)) = "DATA" Then Exit For
        ExtractLocationFromDataFile = ExtractLocationFromDataFile & parts(indexNo) & " "
    Next indexNo
    ExtractLocationFromDataFile = Trim$(ExtractLocationFromDataFile)
End Function

Private Function ExtractDateFromDataFile(ByVal fileNameText As String) As String
    Dim parts() As String

    parts = Split(NormalizeSpaces(fileNameText), " ")
    If UBound(parts) >= 0 Then ExtractDateFromDataFile = parts(0)
End Function

Private Function ExtractDayFromDataFile(ByVal fileNameText As String) As String
    Dim parts() As String
    Dim indexNo As Long
    Dim dayNumber As String
    Dim dotPosition As Long
    Dim foundDataMarker As Boolean

    parts = Split(NormalizeSpaces(fileNameText), " ")
    For indexNo = 0 To UBound(parts) - 1
        If UCase$(parts(indexNo)) = "DATA" Then foundDataMarker = True
        If foundDataMarker And UCase$(parts(indexNo)) = "DAY" Then
            dayNumber = parts(indexNo + 1)
            dotPosition = InStr(1, dayNumber, ".", vbTextCompare)
            If dotPosition > 0 Then dayNumber = Left$(dayNumber, dotPosition - 1)
            ExtractDayFromDataFile = "DAY " & dayNumber
            Exit Function
        End If
    Next indexNo
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

Private Sub CloseOther4GDataWorkbooks()
    Dim wb As Workbook

    For Each wb In Workbooks
        If wb.Name <> "QoS.xlsm" And wb.Name <> "4G Data Template.xlsm" Then
            wb.Close SaveChanges:=False
            Exit For
        End If
    Next wb
End Sub

Private Function ReadFirstFormText(ByVal frm As Object, ParamArray propertyNames() As Variant) As String
    Dim indexNo As Long
    Dim result As String

    For indexNo = LBound(propertyNames) To UBound(propertyNames)
        result = ReadFormText(frm, CStr(propertyNames(indexNo)))
        If result <> "" Then
            ReadFirstFormText = result
            Exit Function
        End If
    Next indexNo
End Function

Private Function ReadFormText(ByVal frm As Object, ByVal propertyName As String) As String
    On Error Resume Next
    ReadFormText = Trim$(CStr(CallByName(frm, propertyName, VbGet)))
    On Error GoTo 0
End Function

Private Function ReadFormBoolean(ByVal frm As Object, ByVal propertyName As String) As Boolean
    On Error Resume Next
    ReadFormBoolean = CBool(CallByName(frm, propertyName, VbGet))
    On Error GoTo 0
End Function
