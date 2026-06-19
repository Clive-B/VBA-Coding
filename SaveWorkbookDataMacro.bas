Attribute VB_Name = "SaveWorkbookDataMacro"
Option Compare Text
Sub SaveTelcoWorkbook(Optional ByVal outputFolderPath As String = "", Optional ByVal dirPathValue As String = "")

    Dim WSVame, WsName, abName As String
    Dim NetworkWorkbook As Workbook
    Dim sourceWorkbook As Workbook
    Dim ws As Worksheet
    Dim oWSHShell As Object
    Dim destFold As String
    Dim lastRow As Long

    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop")
    GetDesktop = GetDesktop & "\"
    Set oWSHShell = Nothing
    Set sourceWorkbook = ActiveWorkbook

'    Sheets("Measurement_Info").Select
'    adName = Sheets("Measurement_Info").Range("A2").Value
'    Pen() = Split(Range("I2").Value, " ")
'    Yop() = Split(adName, " ")
'    Cop() = Split(adName, " DATA ")
'    Dop() = Split(Cop(0), Yop(2))
'    Zop() = Split(Yop(7), ".")
    abName = GetKeepZero & " " & GetTownNameTwo & " " & GetKeepOne & " " & "DATA" & GetDayyZero

    If Trim$(dirPathValue) = "" Then dirPathValue = DirPathGlobal
    If Trim$(outputFolderPath) <> "" Then
        destFold = EnsureTrailingSlash(outputFolderPath)
    Else
        destFold = GetDesktop & "QoS Automation" & "\" & dirPathValue & "\Telcos\" & GetKeepZero & "\DATA\EXCELS\"
    End If
    EnsureFolderExists destFold
    
    Set NetworkWorkbook = Workbooks.Add
    NetworkWorkbook.SaveAs destFold & abName & ".xlsx"
                
'copy second sheet'
    
    WSVame = "GSM_DataReport_multimetric_5"
    Set ws = NetworkWorkbook.Sheets.Add(After:=NetworkWorkbook.Sheets(NetworkWorkbook.Worksheets.Count))
    ws.Name = WSVame

                With sourceWorkbook.Sheets("GSM_DataReport_multimetric_5")
                    lastRow = .Cells(.Rows.Count, "K").End(xlUp).row
                    If lastRow < 1 Then lastRow = 1
                    .Range("K1:R" & lastRow).Copy
                    With NetworkWorkbook.Sheets(WSVame).Range("A1")
                        .PasteSpecial xlPasteValues
                        .PasteSpecial xlPasteFormats
                    End With
                End With
                Application.CutCopyMode = False
                
                
                Application.DisplayAlerts = False
                NetworkWorkbook.Sheets("Sheet1").Delete
                Application.DisplayAlerts = True
            
                'AutoFit Every Worksheet Column in a Workbook
                        For Each ws In NetworkWorkbook.Sheets
                        ws.Cells.EntireColumn.AutoFit
                        Next ws
                'MsgBox "Autofit Successfull!!"
                'Range("E:E").NumberFormat = "ss.00"
                NetworkWorkbook.Sheets(WSVame).Range("B:B").NumberFormat = "hh:mm:ss.000"
                NetworkWorkbook.Sheets(WSVame).Range("A:A").NumberFormat = "mm/dd/yyyy"
                NetworkWorkbook.Close SaveChanges:=True
                


sourceWorkbook.Activate
sourceWorkbook.Save
sourceWorkbook.Sheets(GetKeepZero).Select

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

    If Trim$(folderPath) = "" Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    parentPath = fso.GetParentFolderName(folderPath)
    If parentPath <> "" And Not fso.FolderExists(parentPath) Then EnsureFolderExists parentPath
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub
