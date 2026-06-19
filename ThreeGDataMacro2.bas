Attribute VB_Name = "ThreeGDataMacro2"
' Add this at the top of your module (before any procedures)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public DirPathGlobal As String, workbookPathGlobal As String, GetKeepZero As String, GetKeepOne As String
Public GetSaveZero As String, GetDayyZero As String, endRow As String, Beginning As String, Ending As String, GetTownNameTwo As String
Public SingleDataOutputEnabled As Boolean
Public SingleDataOutputFolder As String
Public SingleDataOutputName As String
Public prevlocKey As Variant ' Declare global variable
Public prevtown, PrevFileDate, PrevlocationName, PrevDupBreak As String ' Declare global variable
Public Sub ConfigureSingleDataOutput(ByVal outputFolder As String, ByVal outputName As String, Optional ByVal enabled As Boolean = True)
    SingleDataOutputFolder = Trim$(outputFolder)
    SingleDataOutputName = Trim$(outputName)
    SingleDataOutputEnabled = enabled
End Sub

Public Sub ClearSingleDataOutput()
    SingleDataOutputEnabled = False
    SingleDataOutputFolder = ""
    SingleDataOutputName = ""
End Sub

Public Sub ThreeGDataSingle(Fil As String, DupBreak As String, LogfileName As String, locationName As String, fileDate As String, dirPath As String, Finish As Boolean, _
               isLastInTownDate As Boolean, isLastArrayElement As Boolean, _
               isLastForLocation As Boolean, FinalTime As Boolean, passNumber As Integer, _
               locKey As Variant, _
               Optional location As String = "", Optional sourceTown As String = "")
    ThreeGData Fil, DupBreak, LogfileName, locationName, fileDate, dirPath, Finish, _
               isLastInTownDate, isLastArrayElement, isLastForLocation, FinalTime, _
               passNumber, locKey, location, sourceTown
End Sub

'Sub ThreeGData(Fil As String)
'Sub ThreeGData(Fil As String, DirPath As String, location As String, Finish As Boolean, isLastArrayElement As Boolean)
Sub ThreeGData(Fil As String, DupBreak As String, LogfileName As String, locationName As String, fileDate As String, dirPath As String, Finish As Boolean, _
               isLastInTownDate As Boolean, isLastArrayElement As Boolean, _
               isLastForLocation As Boolean, FinalTime As Boolean, passNumber As Integer, _
               locKey As Variant, _
               Optional location As String = "", Optional sourceTown As String = "")

Dim KPI As Scripting.textStream
Dim ArryLine As String, latitude As String, longitude As String, Get_date As String, test_date As String, sh As String, dh As String, hh As String, Fgname As String
Dim SpotNum As String
Dim Keep() As String
Dim CurrentLine() As String
Dim fileName() As String
Dim Name() As String
Dim TownName() As String
Dim SaveName() As String
Dim Dayy() As String
Dim GetSaveName() As String
Dim Tdate() As String
Dim wb As Workbook
Dim ws As Worksheet
Dim CurrentRow As Long, Tow As Long, WordCount As Long
Dim ThrputArray(50000) As Variant
Dim i As Long
Dim fso As Scripting.FileSystemObject
Dim lRow As Long
Dim oWSHShell As Object
Dim hasChanged As Boolean
Dim isSame As Boolean
Dim NameTown As String
Dim cellValue As String
Dim numberValue As Long
Dim Count As Long
Dim cellVal As Variant
Dim firstUnusedRow As Long
Dim LastUsedRow As Long
Dim firstFreeRow As Long
Dim LastOccupiedRow As Long
Dim excelSettingsOff As Boolean
Dim errNumber As Long, errDescription As String
Dim lastRow As Long
Dim AA As String
Dim Total As Double, Average As Double
Dim j As Long
Dim spotName As String

On Error GoTo FatalError


If Trim$(sourceTown) <> "" Then
    NameTown = sourceTown
Else
    NameTown = Application.Run("'QoS.xlsm'!GetPublicVar")
End If

Set oWSHShell = CreateObject("WScript.Shell")
GetDesktop = oWSHShell.SpecialFolders("Desktop")
GetDesktop = GetDesktop & "\"
Set oWSHShell = Nothing

i = 0
Set fso = New Scripting.FileSystemObject

'Fil = "C:\Users\LENOVO\Desktop\SERVICE IMPROVEMENT\DATA\25Feb18 085307 Central Feb Qos Data Try DATA Day 1.1.nmf"
'DirPath = "February 2025\"
Fgname = Fil

 ' Set your worksheet - change "Sheet1" to your sheet name or index
Set ws = ThisWorkbook.Worksheets("Logs") ' Or Worksheets(1)

' Method 1: Find last used row in the entire sheet (most reliable)
LastOccupiedRow = LastUsedRowInSheet(ws)

' Method 2: Alternative for last used row in specific column (e.g., Column A)
' lastUsedRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' Find first unused row (next empty row after last used row)
firstFreeRow = LastOccupiedRow + 1

For i = 1 To firstFreeRow
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
            Ending = Sheets("GSM_DataReport_multimetric_5").Range("AO1")
        Else
            hasChanged = True
            If CStr(prevtown) = CStr(NameTown) Then
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
            If SingleDataOutputEnabled Or Dir((GetDesktop & "QoS Automation\" & dirPath & "\NCA\" & GetKeepZero & "\DATA\" & GetKeepZero & " " & prevtown & " " & GetKeepOne & " " & "DATA" & GetDayyZero & ".xlsm")) = "" Then
            Sheets("Locations").Range("A" & firstUnusedRow) = "Spot" & SpotNum
            Sheets("Locations").Range("B" & firstUnusedRow) = GetSaveZero
            cellValue = Sheets("GSM_DataReport_multimetric_5").Range("AO1").value
            numberValue = Val(cellValue)  ' Converts string to number
            Sheets("Locations").Range("C" & firstUnusedRow) = Beginning
            Sheets("Locations").Range("D" & firstUnusedRow) = numberValue - 1
            Beginning = Sheets("GSM_DataReport_multimetric_5").Range("AO1")
'            If Dir((GetDesktop & "QoS Automation\" & dirPath & "\NCA\" & GetKeepZero & "\DATA\" & GetKeepZero & " " & GetSaveZero & " " & GetKeepOne & " " & "DATA" & GetDayyZero & ".xlsm")) = "" Then
'            ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & GetKeepZero & "\DATA\" & GetKeepZero & " " & GetSaveZero & " " & GetKeepOne & " " & "DATA" & GetDayyZero & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
            End If
            End If
            End If
            End If
            PrevFileDate = fileDate
            prevlocKey = locKey
            prevtown = NameTown
            PrevlocationName = locationName
            PrevDupBreak = DupBreak
            ' Optional: Handle the change (e.g., log, exit, etc.)
        End If
Else
Beginning = Sheets("GSM_DataReport_multimetric_5").Range("AO1")
PrevFileDate = fileDate
prevlocKey = locKey
prevtown = NameTown
PrevlocationName = locationName
PrevDupBreak = DupBreak
End If

Set KPI = fso.OpenTextFile(Fgname)

workbookPath = FindworkbookPath()
' Loop through open workbooks to find the one matching the path
    For Each wb In Workbooks
        If wb.fullName = workbookPath Then
            wb.Activate ' Activate the workbook
            Windows(wb.Name).Activate ' Bring its window to the front
            Exit For
        End If
    Next wb
sh = "GSM_DataReport_multimetric_5"

Sheets(sh).Select

CurrentRow = Range("AO1").value
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
            If KeepValue(Keep, 1) = "3G" Then
            If UBound(CurrentLine) < 3 Then GoTo NextLogLine
            ST = CurrentLine(1)
            Get_date = Mid(CurrentLine(3), 2, Len(CurrentLine(3)) - 2)
            Tdate = Split(Get_date, ".")
            If UBound(Tdate) >= 2 Then test_date = Tdate(1) & "/" & Tdate(0) & "/" & Tdate(2)
            End If
        Case Is = "DRATE"
            If KeepValue(Keep, 1) = "3G" Then
            If UBound(CurrentLine) < 6 Then GoTo NextLogLine
            If CurrentLine(4) = "3" Then
                Range("K" & CurrentRow) = test_date
                Range("L" & CurrentRow) = CurrentLine(1)
                Range("M" & CurrentRow) = latitude
                Range("N" & CurrentRow) = longitude
                Range("Q" & CurrentRow) = CurrentLine(6) / 1000
                ThrputArray(i) = Range("Q" & CurrentRow).value
                i = i + 1
                CurrentRow = CurrentRow + 1
            End If
            End If
        Case Is = "DREQ"
            If KeepValue(Keep, 1) = "3G" Then
            If UBound(CurrentLine) < 5 Then GoTo NextLogLine
            If CurrentLine(5) = "3" Then
                Range("K" & CurrentRow) = test_date
                Range("L" & CurrentRow) = CurrentLine(1)
                Range("M" & CurrentRow) = latitude
                Range("N" & CurrentRow) = longitude
                Range("O" & CurrentRow) = "Data Transfer Request"
                CurrentRow = CurrentRow + 1
            End If
            End If
        Case Is = "DCOMP"
            If KeepValue(Keep, 1) = "3G" Then
            If UBound(CurrentLine) < 4 Then GoTo NextLogLine
            If CurrentLine(4) = "3" Then
                Range("K" & CurrentRow) = test_date
                Range("L" & CurrentRow) = CurrentLine(1)
                Range("M" & CurrentRow) = latitude
                Range("N" & CurrentRow) = longitude
                Range("P" & CurrentRow) = "Data Transfer Success"
                For j = LBound(ThrputArray) To i
                Total = Total + ThrputArray(j)
                Next j
                If i > 0 Then
                Average = Total / (i - LBound(ThrputArray))
                Range("R" & CurrentRow) = Average
                End If
                Erase ThrputArray
                i = 0
                j = 0
                Total = 0
                CurrentRow = CurrentRow + 1
                End If
            End If
            
        Case Is = "GPS"
            If KeepValue(Keep, 1) = "3G" Then
            If UBound(CurrentLine) < 4 Then GoTo NextLogLine
            longitude = CurrentLine(3)
            latitude = CurrentLine(4)
            End If
        Case Is = "#STOP"
            If KeepValue(Keep, 1) = "3G" Then
            If UBound(CurrentLine) < 1 Then GoTo NextLogLine
            Sp = CurrentLine(1)
            Exit Do
            End If
    End Select
NextLogLine:
    
    
Loop

KPI.Close

    lRow = Cells(rows.Count, 18).End(xlUp).row
    Sheets(sh).Range("R2:R" & lRow).Font.Color = RGB(255, 0, 0)
    Sheets(sh).Columns(12).NumberFormat = "hh:mm:ss.000"
    Sheets(sh).Columns("K:R").AutoFit
    Range("AO1").value = CurrentRow
    
    If KeepValue(Keep, 1) = "3G" Then
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
DirPathGlobal = dirPath

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
        cellValue = Sheets("GSM_DataReport_multimetric_5").Range("AO1").value
        numberValue = Val(cellValue)  ' Converts string to number
        Sheets("Locations").Range("C" & firstUnusedRow) = Beginning
        Sheets("Locations").Range("D" & firstUnusedRow) = numberValue - 1
        CreateSpotSheets
        Worksheets("Measurement_Info").Range("A" & (lastRow + 1)).value = TownName(2) & " " & Keep(1) & " " & "DATA" & Dayy(0)
        
        ' Set the target worksheet
        Set ws = ThisWorkbook.Sheets("Measurement_Info")
        
        ' Write the formula in cell G1
        ws.Range("G" & (lastRow + 1)).Formula = "=PERCENTILE.INC(GSM_DataReport_multimetric_5!Q:Q,0.9)"
        ActiveWorkbook.SaveAs DataOutputWorkbookPath(GetDesktop, dirPath, Keep(0), TownName(2), Keep(1), Dayy(0)), FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Sleep 3000  ' 3 second delay
        If SingleDataOutputEnabled Then
            SaveTelcoWorkbook SingleDataOutputFolder, dirPath
        Else
            SaveTelcoWorkbook "", dirPath
        End If
        Sleep 3000  ' 3 second delay
        If Not SingleDataOutputEnabled Then
        Copy_3GData_TO_Results
        End If
        Sleep 3000  ' 3 second delay
        BuildStationaryDataMapLayers
        ActiveWorkbook.Save
        Sleep 3000  ' 3 second delay
        Ending = CurrentRow
        
        If Not SingleDataOutputEnabled Then
            Set ws = ThisWorkbook.Sheets("Locations")
            
            ' Find last row in Locations sheet
            lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).row
            
            ' Loop through each row in Locations sheet (starting from row 2)
            For i = 2 To lastRow
            spotName = ws.Cells(i, "A").value
            If SheetExists(spotName) Then
                Application.DisplayAlerts = False
                ThisWorkbook.Sheets(spotName).Delete
                Application.DisplayAlerts = True
            End If
            Next i
            
            Set ws = Nothing
        End If
        
        ' Turn screen updating back on
        Application.ScreenUpdating = True
'        Call DeleteFirstLocRows
        If Not SingleDataOutputEnabled Then
            dh = "Locations"
            hh = "Logs"
            Sheets(dh).Range("A2:AN1000000").ClearContents
            Sheets(dh).Range("B2:AN1000000").ClearContents
            Sheets(dh).Range("C2:AN1000000").ClearContents
            Sheets(dh).Range("D2:AN1000000").ClearContents
            Sheets(sh).Range("A2:AN1000000").ClearContents
            Sheets(hh).Range("A2:E1000000").ClearContents
            Sheets(sh).Range("AO1") = 2
            Beginning = 2
            Sheets("Measurement_Info").Range("A2:I1000000").ClearContents
        End If
        If SingleDataOutputEnabled Then ClearSingleDataOutput
        End If
'        ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & Keep(0) & "\DATA\" & Keep(0) & Name(0) & " " & Keep(1) & " " & "DATA" & Dayy(0) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
        
'    ElseIf passNumber = 1 And isLastInTownDate = False Then
'
'        If Dir((GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & GetKeepZero & "\DATA\" & GetKeepZero & " " & GetSaveZero & " " & GetKeepOne & " " & "DATA" & GetDayyZero & ".xlsm")) = "" Then
'            Ending = CurrentRow
'            Call DeleteFirstLocRows
'        End If
        
    ElseIf passNumber = 2 Then
        
    End If
    Exit Sub

FatalError:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If Not KPI Is Nothing Then KPI.Close
    If excelSettingsOff Then TurnOnStuff
    MsgBox "ThreeGData failed while processing:" & vbCrLf & CStr(Fil) & vbCrLf & _
           "Error " & errNumber & ": " & errDescription, vbExclamation, "ThreeGData"
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
Function FindworkbookPath() As String
    Dim wb As Workbook
    Dim workbookPath As String
    
    ' Initialize the result as an empty string
    workbookPath = ""
    
    ' Loop through all open workbooks
    For Each wb In Workbooks
        ' Exclude "QoS.xlsm" and find the other workbook
        If wb.Name <> "QoS.xlsm" Then
            workbookPath = wb.fullName ' Get the full path of the other workbook
            Exit For ' Exit the loop once the other workbook is found
        End If
    Next wb
    
    ' Return the path of the other workbook (empty string if not found)
    FindworkbookPath = workbookPath
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

Private Function KeepValue(ByVal Keep As Variant, ByVal index As Long) As String
    On Error GoTo Missing
    If index >= LBound(Keep) And index <= UBound(Keep) Then
        KeepValue = CStr(Keep(index))
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

Private Function DataOutputWorkbookPath(ByVal desktopPath As String, ByVal dirPath As String, ByVal operatorName As String, ByVal townNameText As String, ByVal techText As String, ByVal dayToken As String) As String
    Dim folderPath As String
    Dim outputName As String

    If SingleDataOutputEnabled And Trim$(SingleDataOutputFolder) <> "" And Trim$(SingleDataOutputName) <> "" Then
        folderPath = EnsureTrailingSlash(SingleDataOutputFolder)
        outputName = CleanWorkbookFileName(SingleDataOutputName)
        If LCase$(Right$(outputName, 5)) <> ".xlsm" Then outputName = outputName & ".xlsm"
        DataOutputWorkbookPath = folderPath & outputName
    Else
        DataOutputWorkbookPath = desktopPath & "QoS Automation\" & dirPath & "\NCA\" & operatorName & "\DATA\" & _
                                 operatorName & " " & townNameText & " " & techText & " DATA" & dayToken & ".xlsm"
    End If
End Function

Private Function EnsureTrailingSlash(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        EnsureTrailingSlash = folderPath
    Else
        EnsureTrailingSlash = folderPath & "\"
    End If
End Function

Private Function CleanWorkbookFileName(ByVal fileNameText As String) As String
    Dim invalidChars As Variant
    Dim invalidChar As Variant

    CleanWorkbookFileName = Trim$(fileNameText)
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each invalidChar In invalidChars
        CleanWorkbookFileName = Replace(CleanWorkbookFileName, CStr(invalidChar), "_")
    Next invalidChar

    If CleanWorkbookFileName = "" Then CleanWorkbookFileName = "3G Data Output"
End Function
