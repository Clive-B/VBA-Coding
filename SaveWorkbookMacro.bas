Attribute VB_Name = "SaveWorkbookMacro"
Option Compare Text
Sub SaveTelcoWorkbook(Optional ByVal outputFolderPath As String = "", Optional ByVal dirPathValue As String = "")

    Dim WSVame, WsName, abName As String
    Dim NetworkWorkbook As Workbook
    Dim ws As Worksheet
    Dim oWSHShell As Object
    Dim Pen() As String
    Dim Zop() As String
    Dim Yop() As String
    Dim destFold As String
    Dim cstDestFold As String
    Dim mosDestFold As String

    
    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop")
    GetDesktop = GetDesktop & "\"
    Set oWSHShell = Nothing

    Sheets("Measurement_Info").Select
    adName = Sheets("Measurement_Info").Range("A2").value
    Pen() = Split(Range("I2").value, " ")
    Yop() = Split(adName, " ")
    Zop() = Split(Yop(5), ".")
    abName = Pen(0) & " " & Yop(2) & " " & Yop(3) & " " & Yop(4) & " " & Zop(0)
    
    If Trim$(outputFolderPath) <> "" Then
        cstDestFold = EnsureTrailingSlash(outputFolderPath)
        mosDestFold = cstDestFold
    ElseIf Trim$(dirPathValue) <> "" Then
        cstDestFold = GetDesktop & "QoS Automation" & "\" & dirPathValue & "\Telcos\" & Pen(0) & "\CST\EXCELS\"
        mosDestFold = GetDesktop & "QoS Automation" & "\" & dirPathValue & "\Telcos\" & Pen(0) & "\MOS\EXCELS\"
    Else
        MsgBox "SaveTelcoWorkbook needs an output folder or campaign path.", vbExclamation, "Save Workbook"
        Exit Sub
    End If

    destFold = cstDestFold
    
    Set NetworkWorkbook = Workbooks.Add
    NetworkWorkbook.SaveAs destFold & abName & ".xlsx"
                
'copy second sheet'
    
    WSVame = "GSM_VoiceCST_multimetric_6"
    Set ws = NetworkWorkbook.Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = WSVame

                ThisWorkbook.Sheets("GSM_VoiceCST_multimetric_6").Range("K1:O1").Copy
                NetworkWorkbook.Sheets(WSVame).Range("A1:E1").PasteSpecial
                
                ThisWorkbook.Sheets("GSM_VoiceCST_multimetric_6").Range("K2:K1000000").Copy
                NetworkWorkbook.Sheets(WSVame).Range("A2:A1000000").PasteSpecial xlPasteValues
                
                ThisWorkbook.Sheets("GSM_VoiceCST_multimetric_6").Range("L2:L1000000").Copy
                NetworkWorkbook.Sheets(WSVame).Range("B2:B1000000").PasteSpecial xlPasteValues
                
                ThisWorkbook.Sheets("GSM_VoiceCST_multimetric_6").Range("M2:M1000000").Copy
                NetworkWorkbook.Sheets(WSVame).Range("C2:C1000000").PasteSpecial xlPasteValues
                
                ThisWorkbook.Sheets("GSM_VoiceCST_multimetric_6").Range("N2:N1000000").Copy
                NetworkWorkbook.Sheets(WSVame).Range("D2:D1000000").PasteSpecial xlPasteValues
                
                ThisWorkbook.Sheets("GSM_VoiceCST_multimetric_6").Range("O2:O1000000").Copy
                NetworkWorkbook.Sheets(WSVame).Range("E2:E1000000").PasteSpecial xlPasteValues
                
                
                Application.DisplayAlerts = False
                NetworkWorkbook.Sheets("Sheet1").Delete
                Application.DisplayAlerts = True
            
                'AutoFit Every Worksheet Column in a Workbook
                        For Each ws In NetworkWorkbook.Sheets
                        ws.Cells.EntireColumn.AutoFit
                        Next ws
                'MsgBox "Autofit Successfull!!"
                Range("E:E").NumberFormat = "ss.00"
                Range("A:A").NumberFormat = "hh:mm:ss.000"
                Range("B:B").NumberFormat = "mm/dd/yyyY"
                ActiveWorkbook.Close SaveChanges:=True
                


Workbooks(abName & ".xlsm").Activate
ActiveWorkbook.Save
ThisWorkbook.Sheets(Pen(0)).Select

'Save MOS

    abName = Pen(0) & " " & Yop(2) & " MOS " & Yop(4) & " " & Zop(0)
    destFold = mosDestFold
    
    Set NetworkWorkbook = Workbooks.Add
    NetworkWorkbook.SaveAs destFold & abName & ".xlsx"
                
'copy second sheet'
    
    WSVame = "LTE_POLQA_MOS_MultiMetric_5"
    Set ws = NetworkWorkbook.Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = WSVame

                ThisWorkbook.Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("K1:O1").Copy
                NetworkWorkbook.Sheets(WSVame).Range("A1:E1").PasteSpecial
                
                ThisWorkbook.Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("K2:K1000000").Copy
                NetworkWorkbook.Sheets(WSVame).Range("A2:A1000000").PasteSpecial xlPasteValues
                
                ThisWorkbook.Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("L2:L1000000").Copy
                NetworkWorkbook.Sheets(WSVame).Range("B2:B1000000").PasteSpecial xlPasteValues
                
                ThisWorkbook.Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("M2:M1000000").Copy
                NetworkWorkbook.Sheets(WSVame).Range("C2:C1000000").PasteSpecial xlPasteValues
                
                ThisWorkbook.Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("N2:N1000000").Copy
                NetworkWorkbook.Sheets(WSVame).Range("D2:D1000000").PasteSpecial xlPasteValues
                
                ThisWorkbook.Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("O2:O1000000").Copy
                NetworkWorkbook.Sheets(WSVame).Range("E2:E1000000").PasteSpecial xlPasteValues
                
                
                Application.DisplayAlerts = False
                NetworkWorkbook.Sheets("Sheet1").Delete
                Application.DisplayAlerts = True
            
                'AutoFit Every Worksheet Column in a Workbook
                        For Each ws In NetworkWorkbook.Sheets
                        ws.Cells.EntireColumn.AutoFit
                        Next ws
                'MsgBox "Autofit Successfull!!"
                Range("E:E").NumberFormat = "0.00"
                Range("A:A").NumberFormat = "hh:mm:ss.000"
                Range("B:B").NumberFormat = "mm/dd/yyyy"
                ActiveWorkbook.Close SaveChanges:=True
                

abName = Pen(0) & " " & Yop(2) & " " & Yop(3) & " " & Yop(4) & " " & Zop(0)
Workbooks(abName & ".xlsm").Activate
ActiveWorkbook.Save
ThisWorkbook.Sheets(Pen(0)).Select


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
