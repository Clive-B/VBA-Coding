Attribute VB_Name = "FourGDataMacro"
' Add this at the top of your module (before any procedures)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public DirPathGlobal As String, workbookPathGlobal As String, GetKeepZero As String, GetKeepOne As String
Public GetSaveZero As String, GetDayyZero As String, endRow As String, Beginning As String, BeginningHttp As String, BeginningFtp As String, Ending As String, EndingHttp As String, EndingFtp As String, GetTownNameTwo As String
Public prevlocKey As Variant ' Declare global variable
Public prevtown, PrevFileDate, PrevlocationName, PrevDupBreak As String ' Declare global variable
Sub FourGData(Fil As String, DupBreak As String, LogfileName As String, locationName As String, fileDate As String, DirPath As String, Finish As Boolean, _
               isLastInTownDate As Boolean, isLastArrayElement As Boolean, _
               isLastForLocation As Boolean, FinalTime As Boolean, passNumber As Integer, _
               locKey As Variant, _
               Optional location As String = "")

Dim KPI As Scripting.TextStream
Dim ArryLine As String, latitude As String, longitude As String, Get_date As String, test_date As String, sh As String, Fgname As String
Dim CurrentLine() As String
Dim fileName() As String
Dim Tdate() As String
Dim CurrentRowFtp As Long, CurrentRowHttp As Long, CurrentRow As Long, CurrentRowCov As Long, WordCount As Long, WS_Count As Long
Dim fso As Scripting.FileSystemObject
Dim lRow As Long
Dim oWSHShell As Object
Dim TownName() As String
Dim SaveName() As String
Dim Dayy() As String
Dim GetSaveName() As String
Dim hasChanged As Boolean
Dim NameTown As String
Dim cellValue As String, cellValueHttp As String, cellValueFTP As String
Dim numberValue As Long, numberValueHttp As Long, numberValueFTP As Long
Dim Count As Long
Dim cellVal As Variant
Dim firstUnusedRow As Long
Dim LastUsedRow As Long
Dim firstFreeRow As Long
Dim LastOccupiedRow As Long
Dim excelSettingsOff As Boolean
Dim errNumber As Long, errDescription As String
Dim i As Long
Dim isSame As Boolean
Dim ws As Worksheet
Dim AA As String
Dim lastRow As Long
Dim SpotNum As String
Dim Band As String
Dim Device As String
Dim ST As String
Dim Sp As String
Dim MCC As String
Dim MNC As String
Dim spotName As String

On Error GoTo FatalError

NameTown = Application.Run("'QoS.xlsm'!GetPublicVar")

Set oWSHShell = CreateObject("WScript.Shell")
GetDesktop = oWSHShell.SpecialFolders("Desktop")
GetDesktop = GetDesktop & "\"
Set oWSHShell = Nothing

Set fso = New Scripting.FileSystemObject

Fgname = Fil

 ' Set your worksheet - change "Sheet1" to your sheet name or index
Set ws = ThisWorkbook.Worksheets("Logs") ' Or Worksheets(1)

' Method 1: Find last used row in the entire sheet (most reliable)
LastOccupiedRow = LastUsedRowInSheet(ws)

' Method 2: Alternative for last used row in specific column (e.g., Column A)
' lastUsedRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Find first unused row (next empty row after last used row)
firstFreeRow = LastOccupiedRow + 1

For i = 2 To firstFreeRow
        cellVal = ws.Cells(i, "B").value
        isSame = False ' Initialize to False
        
        ' Simple comparison - set isSame to True if values are the same
        If Not IsEmpty(cellVal) Then
            If cellVal = DupBreak Then
                isSame = True
                Exit For
            End If
        End If
        
        
Next i

Sheets("Logs").Range("A" & firstFreeRow) = LogfileName
Sheets("Logs").Range("B" & firstFreeRow) = DupBreak
Sheets("Logs").Range("C" & firstFreeRow) = NameTown & " " & locationName


If Not IsEmpty(prevlocKey) Then
        ' Compare prevValue and currentValue
        If prevlocKey = locKey Then
            hasChanged = False
            Ending = Sheets("Ping").Range("AO1")
            EndingHttp = Sheets("HTTPIPServiceSetupTime").Range("AO1")
            EndingFtp = Sheets("FTPThroughput").Range("AO1")
        Else
            hasChanged = True
             ' Set your worksheet - change "Sheet1" to your sheet name or index
            Set ws = ThisWorkbook.Worksheets("Locations") ' Or Worksheets(1)

            ' Method 1: Find last used row in the entire sheet (most reliable)
            LastUsedRow = LastUsedRowInSheet(ws)

            ' Method 2: Alternative for last used row in specific column (e.g., Column A)
            ' lastUsedRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            ' Find first unused row (next empty row after last used row)
            firstUnusedRow = LastUsedRow + 1
            
            'Get numbers into words
            If LastUsedRow = 1 Then
                SpotNum = "One"
            ElseIf LastUsedRow = 2 Then
                SpotNum = "Two"
            ElseIf LastUsedRow = 3 Then
                SpotNum = "Three"
            ElseIf LastUsedRow = 4 Then
                SpotNum = "Four"
            ElseIf LastUsedRow = 5 Then
                SpotNum = "Five"
            ElseIf LastUsedRow = 6 Then
                SpotNum = "Six"
            ElseIf LastUsedRow = 7 Then
                SpotNum = "Seven"
            ElseIf LastUsedRow = 8 Then
                SpotNum = "Eight"
            ElseIf LastUsedRow = 9 Then
                SpotNum = "Nine"
            ElseIf LastUsedRow = 10 Then
                SpotNum = "Ten"
            End If
            
            If PrevFileDate = fileDate And PrevlocationName <> locationName Then
            
            If isSame = True Then
            Sheets("Locations").Range("D" & LastUsedRow) = numberValue - 1
            Else
            If Dir((GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & GetKeepZero & "\DATA\" & GetKeepZero & " " & prevtown & " " & GetKeepOne & " " & "DATA" & GetDayyZero & ".xlsm")) = "" Then
            Sheets("Locations").Range("A" & firstUnusedRow) = "Spot" & SpotNum
            Sheets("Locations").Range("B" & firstUnusedRow) = GetSaveZero
            cellValue = Sheets("Ping").Range("AO1").value
            numberValue = Val(cellValue)  ' Converts string to number
            cellValueHttp = Sheets("HTTPIPServiceSetupTime").Range("AO1").value
            numberValueHttp = Val(cellValueHttp)
            cellValueFTP = Sheets("FTPThroughput").Range("AO1").value
            numberValueFTP = Val(cellValueFTP)
            Sheets("Locations").Range("C" & firstUnusedRow) = Beginning
            Sheets("Locations").Range("D" & firstUnusedRow) = numberValue - 1
            Sheets("Locations").Range("E" & firstUnusedRow) = BeginningHttp
            Sheets("Locations").Range("F" & firstUnusedRow) = numberValueHttp - 1
            Sheets("Locations").Range("G" & firstUnusedRow) = BeginningFtp
            Sheets("Locations").Range("H" & firstUnusedRow) = numberValueFTP - 1
            Beginning = Sheets("Ping").Range("AO1")
            BeginningHttp = Sheets("HTTPIPServiceSetupTime").Range("AO1")
            BeginningFtp = Sheets("FTPThroughput").Range("AO1")
'            If Dir((GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & GetKeepZero & "\DATA\" & GetKeepZero & " " & GetSaveZero & " " & GetKeepOne & " " & "DATA" & GetDayyZero & ".xlsm")) = "" Then
'            ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & GetKeepZero & "\DATA\" & GetKeepZero & " " & GetSaveZero & " " & GetKeepOne & " " & "DATA" & GetDayyZero & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
'            End If
            End If
            End If
            End If
            PrevFileDate = fileDate
            prevlocKey = locKey
            prevtown = NameTown
            PrevlocationName = locationName
            PrevDupBreak = DupBreak
            
'            If CStr(prevtown) = CStr(NameTown) Then
'            Beginning = Sheets("Ping").Range("AO1")
'            BeginningHttp = Sheets("HTTPIPServiceSetupTime").Range("AO1")
'            BeginningFtp = Sheets("FTPThroughput").Range("AO1")
'            If Dir((GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & GetKeepZero & "\DATA\" & GetKeepZero & " " & GetSaveZero & " " & GetKeepOne & " " & "DATA" & GetDayyZero & ".xlsm")) = "" Then
'            ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & GetKeepZero & "\DATA\" & GetKeepZero & " " & GetSaveZero & " " & GetKeepOne & " " & "DATA" & GetDayyZero & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
'            End If
'            End If
            'Optional: Handle the change (e.g., log, exit, etc.)
        End If
Else
Beginning = Sheets("Ping").Range("AO1")
BeginningHttp = Sheets("HTTPIPServiceSetupTime").Range("AO1")
BeginningFtp = Sheets("FTPThroughput").Range("AO1")
PrevFileDate = fileDate
prevlocKey = locKey
prevtown = NameTown
PrevlocationName = locationName
PrevDupBreak = DupBreak

End If

Set KPI = fso.OpenTextFile(Fgname)
'Sheets("Ping").Select

CurrentRow = Sheets("Ping").Range("AO1")
CurrentRowHttp = Sheets("HTTPIPServiceSetupTime").Range("AO1")
CurrentRowFtp = Sheets("FTPThroughput").Range("AO1")
CurrentRowCov = Sheets("LTE_4GIDLE_MultiMetric_5").Range("AO1")

Sheets("Measurement_Info").Select

Call TurnOffStuff
excelSettingsOff = True
Do Until KPI.AtEndOfStream
    ArryLine = KPI.ReadLine
     CurrentLine = Split(ArryLine, ",")
    
     
     Select Case CurrentLine(0)
     
        Case Is = "#DL"
            If UBound(CurrentLine) < 3 Then GoTo NextLogLine
            Device = CurrentLine(3)
            Keep = Split(CurrentLine(3), " ")
            If UBound(Keep) < 1 Then GoTo NextLogLine
            Keep(0) = Replace(Keep(0), """", "")
            Keep(1) = Replace(Keep(1), """", "")
            GetKeepOne = Keep(1)
            GetKeepZero = Keep(0)
        Case Is = "#START"
            If UBound(CurrentLine) < 3 Then GoTo NextLogLine
            ST = CurrentLine(1)
            Get_date = Mid(CurrentLine(3), 2, Len(CurrentLine(3)) - 2)
            Tdate = Split(Get_date, ".")
            If UBound(Tdate) >= 2 Then test_date = Tdate(1) & "/" & Tdate(0) & "/" & Tdate(2)
        Case Is = "DRATE"
            
            If UBound(CurrentLine) < 6 Then GoTo NextLogLine
            If CurrentLine(4) = "3" Then
                Sheets("FTPThroughput").Select
                Sheets("FTPThroughput").Range("K" & CurrentRowFtp) = test_date
                Sheets("FTPThroughput").Range("L" & CurrentRowFtp) = CurrentLine(1)
                Sheets("FTPThroughput").Range("M" & CurrentRowFtp) = latitude
                Sheets("FTPThroughput").Range("N" & CurrentRowFtp) = longitude
                Sheets("FTPThroughput").Range("Q" & CurrentRowFtp) = CurrentLine(6) / 1000000
                CurrentRowFtp = CurrentRowFtp + 1
            End If
            
        Case Is = "DREQ"
        
            If UBound(CurrentLine) < 5 Then GoTo NextLogLine
            If CurrentLine(5) = "3" Then
                Sheets("FTPThroughput").Select
                Sheets("FTPThroughput").Range("K" & CurrentRowFtp) = test_date
                Sheets("FTPThroughput").Range("L" & CurrentRowFtp) = CurrentLine(1)
                Sheets("FTPThroughput").Range("M" & CurrentRowFtp) = latitude
                Sheets("FTPThroughput").Range("N" & CurrentRowFtp) = longitude
                Sheets("FTPThroughput").Range("O" & CurrentRowFtp) = "Data Transfer Request"
                CurrentRowFtp = CurrentRowFtp + 1
            End If
            
        Case Is = "DCOMP"
            
            If UBound(CurrentLine) < 5 Then GoTo NextLogLine
            If CurrentLine(4) = "3" Then
                Sheets("FTPThroughput").Select
                Sheets("FTPThroughput").Range("K" & CurrentRowFtp) = test_date
                Sheets("FTPThroughput").Range("L" & CurrentRowFtp) = CurrentLine(1)
                Sheets("FTPThroughput").Range("M" & CurrentRowFtp) = latitude
                Sheets("FTPThroughput").Range("N" & CurrentRowFtp) = longitude
                Sheets("FTPThroughput").Range("P" & CurrentRowFtp) = "Data Transfer Success"
                CurrentRowFtp = CurrentRowFtp + 1
                
            ElseIf (CurrentLine(4) = "4" Or CurrentLine(4) = "11") And CurrentLine(5) = "1" Then
                If UBound(CurrentLine) < 15 Then GoTo NextLogLine
                If CurrentLine(14) <> "" And CurrentLine(15) <> "" Then
                    Sheets("HTTPIPServiceSetupTime").Select
                    Sheets("HTTPIPServiceSetupTime").Range("K" & CurrentRowHttp) = test_date
                    Sheets("HTTPIPServiceSetupTime").Range("L" & CurrentRowHttp) = CurrentLine(1)
                    Sheets("HTTPIPServiceSetupTime").Range("M" & CurrentRowHttp) = latitude
                    Sheets("HTTPIPServiceSetupTime").Range("N" & CurrentRowHttp) = longitude
                    Sheets("HTTPIPServiceSetupTime").Range("O" & CurrentRowHttp) = CInt(CurrentLine(14)) - CInt(CurrentLine(15))
                    CurrentRowHttp = CurrentRowHttp + 1
                End If
            End If
            
        
        Case Is = "RTT"
            
            If UBound(CurrentLine) < 6 Then GoTo NextLogLine
            If CurrentLine(4) = "12" Then
                Sheets("Ping").Select
                Sheets("Ping").Range("K" & CurrentRow) = test_date
                Sheets("Ping").Range("L" & CurrentRow) = CurrentLine(1)
                Sheets("Ping").Range("M" & CurrentRow) = latitude
                Sheets("Ping").Range("N" & CurrentRow) = longitude
                Sheets("Ping").Range("O" & CurrentRow) = CurrentLine(6)
                CurrentRow = CurrentRow + 1
            End If
            
            
        Case Is = "CELLMEAS"
                If UBound(CurrentLine) < 12 Then
                CurrentRowCov = CurrentRowCov - 1
                GoTo CurrentLineIncrement
                End If

                If (CurrentLine(3) = "7" Or CurrentLine(3) = "8") And CurrentLine(4) = "0" And IsPositiveNumber(CurrentLine(5)) Then  'checking for LTE technology
                    
                    If CurrentLine(7) = "0" Then
                    Sheets("LTE_4GIDLE_MultiMetric_5").Range("K" & CurrentRowCov) = CurrentLine(1)
                    Sheets("LTE_4GIDLE_MultiMetric_5").Range("L" & CurrentRowCov) = test_date
                    Sheets("LTE_4GIDLE_MultiMetric_5").Range("M" & CurrentRowCov) = latitude
                    Sheets("LTE_4GIDLE_MultiMetric_5").Range("N" & CurrentRowCov) = longitude
                    Sheets("LTE_4GIDLE_MultiMetric_5").Range("O" & CurrentRowCov) = CurrentLine(12)
                    Sheets("LTE_4GIDLE_MultiMetric_5").Range("P" & CurrentRowCov) = CurrentLine(9)
                    Sheets("LTE_4GIDLE_MultiMetric_5").Range("Q" & CurrentRowCov) = MNC
                    Sheets("LTE_4GIDLE_MultiMetric_5").Range("R" & CurrentRowCov) = MCC
                    Band = CurrentLine(8)
                    Sheets("LTE_4GIDLE_MultiMetric_5").Range("S" & CurrentRowCov) = getBand((Band))
                    Else
                    CurrentRowCov = CurrentRowCov - 1
                    GoTo CurrentLineIncrement
                    End If
                
                Else
                CurrentRowCov = CurrentRowCov - 1
                GoTo CurrentLineIncrement
                End If
CurrentLineIncrement:
                CurrentRowCov = CurrentRowCov + 1
            
        Case Is = "SEI"
            If UBound(CurrentLine) < 7 Then GoTo NextLogLine
            MCC = CurrentLine(6)
            MNC = CurrentLine(7)
        
        Case Is = "GPS"
            If UBound(CurrentLine) < 4 Then GoTo NextLogLine
            longitude = CurrentLine(3)
            latitude = CurrentLine(4)
            
        Case Is = "#STOP"
            If UBound(CurrentLine) < 1 Then GoTo NextLogLine
    
            Sp = CurrentLine(1)
            Exit Do
    
    End Select
NextLogLine:
    
    
Loop

KPI.Close

WS_Count = ActiveWorkbook.Worksheets.Count
For i = 1 To WS_Count

    If Sheets(i).Name = "Ping" Or Sheets(i).Name = "HTTPIPServiceSetupTime" Or Sheets(i).Name = "FTPThroughput" Or Sheets(i).Name = "LTE_4GIDLE_MultiMetric_5" Then
        Sheets(i).Select
        If Sheets(i).Name = "LTE_4GIDLE_MultiMetric_5" Then
'        Sheets("LTE_4GIDLE_MultiMetric_5").Select
        Sheets("LTE_4GIDLE_MultiMetric_5").Columns(11).NumberFormat = "hh:mm:ss.000"
        Sheets("LTE_4GIDLE_MultiMetric_5").Columns(12).NumberFormat = "mm/dd/yyyy"
        Else
        Sheets(i).Columns(12).NumberFormat = "hh:mm:ss.000"
        Sheets(i).Columns("K:S").AutoFit
        End If
    End If

Next i

Sheets("Ping").Range("AO1") = CurrentRow
Sheets("HTTPIPServiceSetupTime").Range("AO1") = CurrentRowHttp
Sheets("FTPThroughput").Range("AO1") = CurrentRowFtp
Sheets("LTE_4GIDLE_MultiMetric_5").Range("AO1") = CurrentRowCov

Sheets("Measurement_Info").Select

    If KeepValue(Keep, 1) = "4G" Then
    Sheets("Measurement_Info").Select
    lastRow = Cells(rows.Count, "A").End(xlUp).row
    lastRow = lastRow + 1
    
    fileName = Split(Fil, "\")
    WordCount = UBound(fileName())
    Worksheets("Measurement_Info").Range("D" & lastRow).value = test_date
    Worksheets("Measurement_Info").Range("A" & lastRow).value = fileName(WordCount)
    Worksheets("Measurement_Info").Range("I" & lastRow).value = Replace(Device, """", "")
    Worksheets("Measurement_Info").Range("B" & lastRow).value = ST
    Worksheets("Measurement_Info").Range("C" & lastRow).value = Sp
    Sheets("Measurement_Info").Columns(2).NumberFormat = "hh:mm:ss.000"
    Sheets("Measurement_Info").Columns(3).NumberFormat = "hh:mm:ss.000"
'
'AA = FileName(WordCount)
'    Name = Split(AA, " ")
'    Name = Split(AA, Name(1))
'    Name = Split(Name(1), " DATA")
'    Dayy = Split(AA, " DATA")
'    Dayy = Split(Dayy(1), ".")
    
    AA = fileName(WordCount)
    Name = Split(AA, " ")
    TownName = Split(AA, " ")
    GetTownNameTwo = TownName(2)
    GetSaveName = Split(AA, TownName(1) & " ")
    SaveName = Split(GetSaveName(1), " DATA")
    GetSaveZero = SaveName(0)
    Name = Split(AA, Name(1))
    Name = Split(Name(1), " DATA")
    Dayy = Split(AA, " DATA")
    Dayy = Split(Dayy(1), ".")
    GetDayyZero = Dayy(0)
    End If
    

Call TurnOnStuff
excelSettingsOff = False

    DirPathGlobal = DirPath

    If passNumber = 1 And isLastInTownDate = True Then
        If FinalTime = True Then
        Set ws = ThisWorkbook.Worksheets("Locations") ' Or Worksheets(1)
        
        ' Method 1: Find last used row in the entire sheet (most reliable)
        LastUsedRow = LastUsedRowInSheet(ws)
        
        ' Method 2: Alternative for last used row in specific column (e.g., Column A)
        ' lastUsedRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Find first unused row (next empty row after last used row)
        firstUnusedRow = LastUsedRow + 1
        
        'Get numbers into words
        If LastUsedRow = 1 Then
            SpotNum = "One"
        ElseIf LastUsedRow = 2 Then
            SpotNum = "Two"
        ElseIf LastUsedRow = 3 Then
            SpotNum = "Three"
        ElseIf LastUsedRow = 4 Then
            SpotNum = "Four"
        ElseIf LastUsedRow = 5 Then
            SpotNum = "Five"
        ElseIf LastUsedRow = 6 Then
            SpotNum = "Six"
        ElseIf LastUsedRow = 7 Then
            SpotNum = "Seven"
        ElseIf LastUsedRow = 8 Then
            SpotNum = "Eight"
        ElseIf LastUsedRow = 9 Then
            SpotNum = "Nine"
        ElseIf LastUsedRow = 10 Then
            SpotNum = "Ten"
        End If
        
        Sheets("Locations").Range("A" & firstUnusedRow) = "Spot" & SpotNum
        Sheets("Locations").Range("B" & firstUnusedRow) = GetSaveZero
        cellValue = Sheets("Ping").Range("AO1").value
        numberValue = Val(cellValue)  ' Converts string to number
        cellValueHttp = Sheets("HTTPIPServiceSetupTime").Range("AO1").value
        numberValueHttp = Val(cellValueHttp)
        cellValueFTP = Sheets("FTPThroughput").Range("AO1").value
        numberValueFTP = Val(cellValueFTP)
        Sheets("Locations").Range("C" & firstUnusedRow) = Beginning
        Sheets("Locations").Range("D" & firstUnusedRow) = numberValue - 1
        Sheets("Locations").Range("E" & firstUnusedRow) = BeginningHttp
        Sheets("Locations").Range("F" & firstUnusedRow) = numberValueHttp - 1
        Sheets("Locations").Range("G" & firstUnusedRow) = BeginningFtp
        Sheets("Locations").Range("H" & firstUnusedRow) = numberValueFTP - 1
        CreateSpotSheets
        Worksheets("Measurement_Info").Range("A" & (lastRow + 1)).value = TownName(2) & " " & KeepValue(Keep, 1) & " " & "DATA" & Dayy(0)
        
        ' Set the target worksheet
        Set ws = ThisWorkbook.Sheets("Measurement_Info")
        
        ' Write the formula in cell F,G,H
        ws.Range("F" & (lastRow + 1)).Formula = "=AVERAGE(Ping!O:O)"
        ws.Range("G" & (lastRow + 1)).Formula = "=AVERAGE(HTTPIPServiceSetupTime!O:O)"
        ws.Range("H" & (lastRow + 1)).Formula = "=AVERAGE(FTPThroughput!Q:Q)"
        ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & KeepValue(Keep, 0) & "\DATA\" & KeepValue(Keep, 0) & " " & TownName(2) & " " & KeepValue(Keep, 1) & " " & "DATA" & Dayy(0) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
        CopyARFCNChannels
        Sleep 3000  ' 3 second delay
        SaveTelcoWorkbook
        Sleep 3000  ' 3 second delay
        Copy_3GData_TO_Results
        Sleep 3000  ' 3 second delay
        BuildStationaryDataMapLayers
        ActiveWorkbook.Save
        Sleep 3000  ' 3 second delay
        Ending = CurrentRow
        EndingHttp = CurrentRowHttp
        EndingFtp = CurrentRowFtp
        
        Set ws = ThisWorkbook.Sheets("Locations")
        
        ' Find last row in Locations sheet
        lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
        
        Application.DisplayAlerts = False
        
        ' Loop through each row in Locations sheet, starting from row 2
        For i = 2 To lastRow
        
            spotName = ws.Cells(i, "A").value & "Ping"
            If SheetExists(spotName) Then ThisWorkbook.Sheets(spotName).Delete
        
            spotName = ws.Cells(i, "A").value & "Http"
            If SheetExists(spotName) Then ThisWorkbook.Sheets(spotName).Delete
        
            spotName = ws.Cells(i, "A").value & "FTP"
            If SheetExists(spotName) Then ThisWorkbook.Sheets(spotName).Delete
        
        Next i
        
        Application.DisplayAlerts = True
        
        Set ws = Nothing
'        Set ws = ThisWorkbook.Sheets("Locations")
'
'        ' Find last row in Locations sheet
'        lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
'
'        Dim sheetNames As Variant
'        ' Loop through each row in Locations sheet (starting from row 2)
'        For i = 2 To lastRow
'        spotName = ws.Cells(i, "A").value + "Ping"
'        ' Turn off screen updating for faster execution
'        Application.DisplayAlerts = False
'        ThisWorkbook.Sheets(spotName).Delete
'        Application.DisplayAlerts = True  ' Restore confirmation dialogs
'        spotName = ws.Cells(i, "A").value + "Http"
'        ' Turn off screen updating for faster execution
'        Application.DisplayAlerts = False
'        ThisWorkbook.Sheets(spotName).Delete
'        Application.DisplayAlerts = True  ' Restore confirmation dialogs
'        spotName = ws.Cells(i, "A").value + "FTP"
'        ' Turn off screen updating for faster execution
'        Application.DisplayAlerts = False
'        ThisWorkbook.Sheets(spotName).Delete
'        Application.DisplayAlerts = True  ' Restore confirmation dialogs
'ContinueLoop:
'        ' Reset worksheet object for next iteration
'        Next i
'
'        Set ws = Nothing
        
        ' Turn screen updating back on
        Application.ScreenUpdating = True
'        Call DeleteFirstLocRows
        dh = "Locations"
        hh = "Logs"
        fh = "Ping"
        rr = "HTTPIPServiceSetupTime"
        ss = "FTPThroughput"
        rs = "LTE_4GIDLE_MultiMetric_5"
        Sheets(dh).Range("A2:AN1000000").ClearContents
        Sheets(dh).Range("B2:AN1000000").ClearContents
        Sheets(dh).Range("C2:AN1000000").ClearContents
        Sheets(dh).Range("D2:AN1000000").ClearContents
'        Sheets(sh).Range("A2:AN1000000").ClearContents
        Sheets(hh).Range("A2:E1000000").ClearContents
'        Sheets(sh).Range("AO1") = 2
        Sheets(fh).Range("K2:O1000000").ClearContents
        Sheets(rr).Range("K2:O1000000").ClearContents
        Sheets(ss).Range("K2:Q1000000").ClearContents
        Sheets(rs).Range("K2:T1000000").ClearContents
        Sheets("Measurement_Info").Range("A2:I1000000").ClearContents
'        Call DeleteFirstLocRows
        Sheets("Measurement_Info").Range("A2:E1000000").ClearContents
        Sheets("Measurement_Info").Range("AO1") = 2
        Sheets("Measurement_Info").Range("A2:D2").ClearContents
        Sheets("Measurement_Info").Range("I2").ClearContents
        Sheets("Ping").Range("AO1") = 2
        Sheets("HTTPIPServiceSetupTime").Range("AO1") = 2
        Sheets("FTPThroughput").Range("AO1") = 2
        Sheets("LTE_4GIDLE_MultiMetric_5").Range("AO1") = 2
        Beginning = Sheets("Ping").Range("AO1")
        BeginningHttp = Sheets("HTTPIPServiceSetupTime").Range("AO1")
        BeginningFtp = Sheets("FTPThroughput").Range("AO1")

    
        End If
    End If

    'ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & Keep(0) & " " & Name(0) & " " & Keep(1) & " " & "DATA" & " " & Dayy(0) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
'    ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & Keep(0) & Name(0) & " " & Keep(1) & " " & "DATA" & Dayy(0) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Exit Sub

FatalError:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If Not KPI Is Nothing Then KPI.Close
    If excelSettingsOff Then TurnOnStuff
    MsgBox "FourGData failed while processing:" & vbCrLf & CStr(Fil) & vbCrLf & _
           "Error " & errNumber & ": " & errDescription, vbExclamation, "FourGData"
End Sub

Sub TurnOffStuff()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

Sub TurnOnStuff()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
Function getBand(Band As String) As String
Dim result As String
    
    Select Case Band
    
    Case Is = "70001"
    
        result = "LTE FDD 2100"
        
    Case Is = "70002"
    
        result = "LTE FDD 1900"
        
    Case Is = "70003"
    
        result = "LTE FDD 1800"
        
    Case Is = "70004"
    
        result = "LTE FDD 2100 AWS"
        
    Case Is = "70005"
    
        result = "LTE FDD 850"
        
    Case Is = "70006"
    
        result = "LTE FDD 850"
        
    Case Is = "70007"
    
        result = "LTE FDD 2600"
        
    Case Is = "70008"
    
        result = "LTE FDD 900"
        
    Case Is = "70009"
    
        result = "LTE FDD 1800"
        
    Case Is = "70010"
    
        result = "LTE FDD 2100"
        
    Case Is = "70011"
    
        result = "LTE FDD 1400"
    
    Case Is = "70012"
    
        result = "LTE FDD 700"
        
    Case Is = "70013"
    
        result = "LTE FDD 700"
        
    Case Is = "70014"
    
        result = "LTE FDD 700"
        
    Case Is = "70017"
    
        result = "LTE FDD 700"
        
    Case Is = "70018"
    
        result = "LTE FDD 850"
        
    Case Is = "70019"
    
        result = "LTE FDD 850"
        
    Case Is = "70020"
    
        result = "LTE FDD 800"
        
    Case Is = "70021"
    
        result = "LTE FDD 1500"
        
    Case Is = "70022"
    
        result = "LTE FDD 3500"
        
    Case Is = "70023"
    
        result = "LTE FDD 2200"
        
    Case Is = "70024"
    
        result = "LTE FDD 1500"
        
    Case Is = "70025"
    
        result = "LTE FDD 1900"
        
    Case Is = "70026"
    
        result = "LTE FDD 850"
        
    Case Is = "70027"
    
        result = "LTE FDD 800"
        
    Case Is = "70028"
    
        result = "LTE FDD 700"
        
    Case Is = "70029"
    
        result = "LTE FDD 700"
        
    Case Is = "70030"
    
        result = "LTE FDD 2350"
        
    Case Is = "70031"
    
        result = "LTE FDD 450"
        
    Case Is = "70032"
    
        result = "LTE FDD 1500 L"
        
    Case Is = "70064"
    
        result = "LTE FDD 390-470"
        
    Case Is = "70065"
    
        result = "LTE FDD 2100"
        
    Case Is = "70066"
    
        result = "LTE FDD AWS-3 2100"
        
    Case Is = "70067"
    
        result = "LTE FDD 700 EU"
        
    Case Is = "70068"
    
        result = "LTE FDD 700 ME"
        
    Case Is = "70069"
    
        result = "LTE FDD 2500"
        
    Case Is = "70070"
    
        result = "LTE FDD AWS-4"
        
    Case Is = "70071"
    
        result = "LTE FDD 600"
        
    Case Is = "70072"
    
        result = "LTE FDD 450 PMR/PAMR"
        
    Case Is = "70073"
    
        result = "LTE FDD 450 APAC"
        
    Case Is = "70074"
    
        result = "LTE FDD 1500 USA L"
        
    Case Is = "70075"
    
        result = "LTE FDD 1500 EU L"
        
    Case Is = "70076"
    
        result = "LTE FDD 1500 ext EU L"
        
    Case Is = "70085"
    
        result = "LTE FDD 700 a+"
        
    Case Is = "70087"
    
        result = "LTE FDD 420-425"
        
    Case Is = "70088"
    
        result = "LTE FDD 422-427"
        
    Case Is = "70240"
    
        result = "LTE FDD 5154-5925"
        
    Case Is = "70250"
    
        result = "LTE FDD 3550-3700"
        
    Case Is = "70252"
    
        result = "LTE FDD 5200 NII-1"
        
    Case Is = "70255"
    
        result = "LTE FDD 5700 NII-3"
        
    Case Is = "79999"
    
        result = "LTE FDD"
        
    Case Is = "80033"
    
        result = "LTE TDD 1900-1920"
        
    Case Is = "80034"
    
        result = "LTE TDD 2010-2025"
        
    Case Is = "80035"
    
        result = "LTE TDD 1850-1910"
        
    Case Is = "80036"
    
        result = "LTE TDD 1930-1990"
        
    Case Is = "80037"
    
        result = "LTE TDD 1910-1930"
        
    Case Is = "80038"
    
        result = "LTE TDD 2570-2620"
        
    Case Is = "80039"
    
        result = "LTE TDD 1880-1920"
        
    Case Is = "80040"
    
        result = "LTE TDD 2300-2400"
        
    Case Is = "80041"
    
        result = "LTE TDD 2496-2690"
        
    Case Is = "80042"
    
        result = "LTE TDD 3400-3600"
        
    Case Is = "80043"
    
        result = "LTE TDD 3600-3800"
        
    Case Is = "80044"
    
        result = "LTE TDD 703-803"
        
    Case Is = "80045"
    
        result = "LTE TDD 1447-1467"
        
    Case Is = "80046"
    
        result = "LTE TDD 5154-5925"
        
    Case Is = "80047"
    
        result = "LTE TDD 5855-5925"
        
    Case Is = "80048"
    
        result = "LTE TDD 3550-3700"
        
    Case Is = "80049"
    
        result = "LTE TDD 3550-3700"
        
    Case Is = "80050"
    
        result = "LTE TDD 1432-1517"
        
    Case Is = "80051"
    
        result = "LTE TDD 1427-1432"
        
    Case Is = "80052"
    
        result = "LTE TDD 3300-3400"
    
    Case Is = "80053"
    
        result = "LTE TDD 2483-2495"
        
    Case Is = "80061"
    
        result = "LTE TDD 1447-1467"
        
    Case Is = "80062"
    
        result = "LTE TDD 1785-1805"
        
    Case Is = "80087"
    
        result = "LTE TDD 1447-1467"
        
    Case Is = "80088"
    
        result = "LTE TDD 1785-1805"
        
    Case Is = "89999"
    
        result = "LTE TDD"
    
    End Select
    
    getBand = result
    
End Function
' Helper function to delete a sheet if it exists
' Returns True if sheet existed and was deleted, False if it didn't exist
Function DeleteSheetIfExists(sheetName As String) As Boolean
    On Error Resume Next ' Ignore errors if sheet doesn't exist
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' Prevent confirmation dialog
        ThisWorkbook.Sheets(sheetName).Delete
        Application.DisplayAlerts = True
        DeleteSheetIfExists = True
    Else
        DeleteSheetIfExists = False
    End If
    
    On Error GoTo 0 ' Reset error handling
End Function
Private Function LastUsedRowInSheet(ByVal ws As Worksheet) As Long
    Dim foundCell As Range

    Set foundCell = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedRowInSheet = 1
    Else
        LastUsedRowInSheet = foundCell.row
    End If
End Function

Private Function IsPositiveNumber(ByVal valueText As Variant) As Boolean
    If IsNumeric(valueText) Then
        IsPositiveNumber = (CDbl(valueText) > 0)
    End If
End Function

Private Function KeepValue(ByRef Keep() As String, ByVal index As Long) As String
    On Error GoTo Missing
    If index >= LBound(Keep) And index <= UBound(Keep) Then
        KeepValue = Keep(index)
    End If
    Exit Function
Missing:
    KeepValue = ""
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim sh As Worksheet

    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not sh Is Nothing
End Function
