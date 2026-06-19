Attribute VB_Name = "SaveWorkbookFourGMacro"
Option Compare Text

Sub SaveTelcoWorkbook(Optional ByVal outputFolderPath As String = "", Optional ByVal dirPathValue As String = "")

    Dim abName As String
    Dim destFold As String
    Dim dirPathText As String
    Dim NetworkWorkbook As Workbook
    Dim sourceWorkbook As Workbook
    Dim ws As Worksheet
    Dim oWSHShell As Object
    Dim GetDesktop As String

    Set sourceWorkbook = ActiveWorkbook

    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop")
    GetDesktop = GetDesktop & "\"
    Set oWSHShell = Nothing

    dirPathText = Trim$(dirPathValue)
    If dirPathText = "" Then dirPathText = DirPathGlobal

    If Trim$(outputFolderPath) <> "" Then
        destFold = EnsureTrailingSlash(outputFolderPath)
        If Trim$(Single4GDataOutputName) <> "" Then
            abName = CleanWorkbookFileName(Single4GDataOutputName)
        Else
            abName = CleanWorkbookFileName(GetKeepZero & " " & GetTownNameTwo & " " & GetKeepOne & " DATA" & GetDayyZero)
        End If
    Else
        destFold = GetDesktop & "QoS Automation\" & dirPathText & "\Telcos\" & GetKeepZero & "\DATA\EXCELS\"
        abName = CleanWorkbookFileName(GetKeepZero & " " & GetTownNameTwo & " " & GetKeepOne & " DATA" & GetDayyZero)
    End If
    EnsureFolderExists destFold

    Set NetworkWorkbook = Workbooks.Add
    NetworkWorkbook.SaveAs destFold & abName & ".xlsx"

    CopyDataSheet sourceWorkbook, NetworkWorkbook, "Ping", "K", "O"
    CopyDataSheet sourceWorkbook, NetworkWorkbook, "HTTPIPServiceSetupTime", "K", "O"
    CopyDataSheet sourceWorkbook, NetworkWorkbook, "FTPThroughput", "K", "Q"

    DeleteSheetIfPresent NetworkWorkbook, "Sheet1"

    For Each ws In NetworkWorkbook.Worksheets
        ws.Cells.EntireColumn.AutoFit
    Next ws

    NetworkWorkbook.Close SaveChanges:=True

    sourceWorkbook.Activate
    sourceWorkbook.Save
    sourceWorkbook.Sheets(GetKeepZero).Select

End Sub

Private Sub CopyDataSheet(ByVal sourceWorkbook As Workbook, ByVal targetWorkbook As Workbook, ByVal sheetName As String, ByVal firstCol As String, ByVal lastCol As String)
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim sourceRange As Range

    Set sourceSheet = sourceWorkbook.Worksheets(sheetName)
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, firstCol).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1

    Set targetSheet = targetWorkbook.Worksheets.Add(After:=targetWorkbook.Worksheets(targetWorkbook.Worksheets.Count))
    targetSheet.Name = sheetName

    Set sourceRange = sourceSheet.Range(firstCol & "1:" & lastCol & lastRow)
    sourceRange.Copy
    With targetSheet.Range("A1")
        .PasteSpecial xlPasteValues
        .PasteSpecial xlPasteFormats
    End With
    Application.CutCopyMode = False

    targetSheet.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    targetSheet.Columns("B:B").NumberFormat = "hh:mm:ss.000"
End Sub

Private Sub DeleteSheetIfPresent(ByVal targetWorkbook As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = targetWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

Private Function EnsureTrailingSlash(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        EnsureTrailingSlash = folderPath
    Else
        EnsureTrailingSlash = folderPath & "\"
    End If
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Dim parentPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If fso.FolderExists(folderPath) Then Exit Sub

    parentPath = fso.GetParentFolderName(folderPath)
    If parentPath <> "" And Not fso.FolderExists(parentPath) Then
        EnsureFolderExists parentPath
    End If

    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

Private Function CleanWorkbookFileName(ByVal fileNameText As String) As String
    Dim badChars As Variant
    Dim badChar As Variant

    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    fileNameText = Trim$(fileNameText)
    For Each badChar In badChars
        fileNameText = Replace(fileNameText, CStr(badChar), " ")
    Next badChar

    Do While InStr(fileNameText, "  ") > 0
        fileNameText = Replace(fileNameText, "  ", " ")
    Loop

    CleanWorkbookFileName = fileNameText
End Function
