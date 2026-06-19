Attribute VB_Name = "FourGCovAUI"
Option Base 1
' Add this at the top of your module (before any procedures)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public DirPathGlobal As String, workbookPathGlobal As String, GetKeepZero As String, GetKeepOne As String
Public GetSaveZero As String, GetDayyZero As String, GetTownNameTwo As String
Public prevtownDayKey As Variant ' Declare global variable
Public prevtown As String ' Declare global variable
'Sub FourGCov(Fil As String)
Sub FourGCov(Fil As String, DirPath As String, Finish As Boolean, _
               isLastInTownDate As Boolean, isLastArrayElement As Boolean, _
               isLastForLocation As Boolean, FinalTime As Boolean, passNumber As Integer, _
               townDayKey As Variant, _
               Optional location As String = "")


Dim KPI As Scripting.TextStream
Dim ArryLine As String, cmrt As String, Alert As String, latitude As String, longitude As String, Get_date As String, test_date As String, CallAttempt As String, sh As String, Fgname As String, Band As String, Device As String, ST As String, Sp As String
Dim Keep() As String
Dim WordCount As Long
Dim CurrentLine() As String
Dim fileName() As String
Dim Tdate() As String
Dim wb As Workbook
Dim ws As Worksheet
Dim CurrentRow As Long, Tow As Long
Dim fso As Scripting.FileSystemObject
Dim lRow As Long
Dim LR As Long
Dim Dayy() As String
Dim hasChanged As Boolean
Dim NameTown As String
Dim workbookName As String
Dim oWSHShell As Object
Dim excelSettingsOff As Boolean
Dim errNumber As Long, errDescription As String

On Error GoTo FatalError


Set oWSHShell = CreateObject("WScript.Shell")
GetDesktop = oWSHShell.SpecialFolders("Desktop")
GetDesktop = GetDesktop & "\"
Set oWSHShell = Nothing

NameTown = Application.Run("'QoS.xlsm'!GetPublicVari")
NetworkCurr = Application.Run("'QoS.xlsm'!GetNetwork")
keepOne = Application.Run("'QoS.xlsm'!GetKeepOne")
GetDay = Application.Run("'QoS.xlsm'!GetDayy")

If Dir((GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & NetworkCurr & "\COVERAGE\" & NetworkCurr & " " & NameTown & " " & keepOne & " " & "COVERAGE" & GetDay & ".xlsm")) = "" Then


Set fso = New Scripting.FileSystemObject
Fgname = Fil

If Not IsEmpty(prevtownDayKey) Then
        ' Compare prevValue and currentValue
        If prevtownDayKey = townDayKey Then
            hasChanged = False
'            Ending = Sheets("UMTS_3GIDLE_MultiMetric_5").Range("AO1")
        Else
            hasChanged = True
            If CStr(prevtown) <> CStr(NameTown) Then
'            Beginning = Sheets("UMTS_3GIDLE_MultiMetric_5").Range("AO1")
            If Dir((GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & GetKeepZero & "\COVERAGE\" & GetKeepZero & " " & GetSaveZero & " " & GetKeepOne & " " & "COVERAGE" & GetDayyZero & ".xlsm")) = "" Then
            
            Else
            Application.DisplayAlerts = False
            ' Delete sheets if they exist
            DeleteSheetIfExists prevtown
            DeleteSheetIfExists "FILTER"
            DeleteSheetIfExists "AdvFILTER"
            DeleteSheetIfExists "AVGCAL"
            Sheets.Add(After:=Sheets(Sheets.count)).Name = "AVGCAL"
            ThisWorkbook.Sheets("AVGCAL").Visible = xlSheetHidden
            Application.DisplayAlerts = True
            End If
            ElseIf CStr(prevtown) = CStr(NameTown) And prevtownDayKey <> townDayKey Then
            Application.DisplayAlerts = False
            ' Delete sheets if they exist
            DeleteSheetIfExists prevtown
            DeleteSheetIfExists "FILTER"
            DeleteSheetIfExists "AdvFILTER"
            DeleteSheetIfExists "AVGCAL"
            Sheets.Add(After:=Sheets(Sheets.count)).Name = "AVGCAL"
            ThisWorkbook.Sheets("AVGCAL").Visible = xlSheetHidden
            Application.DisplayAlerts = True
            End If
            prevtownDayKey = townDayKey
            prevtown = NameTown
            ' Optional: Handle the change (e.g., log, exit, etc.)
        End If
Else
prevtownDayKey = townDayKey
prevtown = NameTown
End If

Set KPI = fso.OpenTextFile(Fgname)
sh = "LTE_4GIDLE_MultiMetric_5"

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
            If UBound(CurrentLine) < 3 Then GoTo NextLogLine
            ST = CurrentLine(1)
            Get_date = Mid(CurrentLine(3), 2, Len(CurrentLine(3)) - 2)
            Tdate = Split(Get_date, ".")
            If UBound(Tdate) >= 2 Then test_date = Tdate(1) & "/" & Tdate(0) & "/" & Tdate(2)
            
        Case Is = "CELLMEAS"
                If UBound(CurrentLine) < 12 Then
                    CurrentRow = CurrentRow - 1
                    GoTo CurrentLineIncrement
                End If

                If (CurrentLine(3) = "7" Or CurrentLine(3) = "8") And CurrentLine(4) = "0" And IsPositiveNumber(CurrentLine(5)) Then  'checking for LTE technology
                    
                    If CurrentLine(7) = "0" Then
                    Range("K" & CurrentRow) = CurrentLine(1)
                    Range("L" & CurrentRow) = test_date
                    Range("M" & CurrentRow) = latitude
                    Range("N" & CurrentRow) = longitude
                    Range("O" & CurrentRow) = CurrentLine(12)
                    Range("P" & CurrentRow) = CurrentLine(9)
                    Range("Q" & CurrentRow) = MNC
                    Range("R" & CurrentRow) = MCC
                    Band = CurrentLine(8)
                    Range("S" & CurrentRow) = getBand((Band))
                    Else
                    CurrentRow = CurrentRow - 1
                    GoTo CurrentLineIncrement
                    End If
                
                Else
                CurrentRow = CurrentRow - 1
                GoTo CurrentLineIncrement
                End If
CurrentLineIncrement:
                CurrentRow = CurrentRow + 1
        Case Is = "GPS"
            If UBound(CurrentLine) < 4 Then GoTo NextLogLine
        
            longitude = CurrentLine(3)
            latitude = CurrentLine(4)
            
        Case Is = "SEI"
            If UBound(CurrentLine) < 7 Then GoTo NextLogLine
            MCC = CurrentLine(6)
            MNC = CurrentLine(7)
            
        Case Is = "#STOP"
            If UBound(CurrentLine) < 1 Then GoTo NextLogLine
            Sp = CurrentLine(1)
            Exit Do
    
    End Select
NextLogLine:
    
    
Loop

KPI.Close

'Range("AO1").Value = CurrentRow
'FileName = Split(Fil, "\")
'WordCount = UBound(FileName())
'Worksheets("Measurement_Info").Range("D2").Value = test_date
'Worksheets("Measurement_Info").Range("A2").Value = FileName(WordCount)
'Worksheets("Measurement_Info").Range("I2").Value = Replace(Device, """", "")
'Worksheets("Measurement_Info").Range("B2").Value = ST
'Worksheets("Measurement_Info").Range("C2").Value = Sp

Range("AO1").value = CurrentRow
Sheets("Measurement_Info").Select
LR = Cells(Rows.count, 1).End(xlUp).row
LR = LR + 1

fileName = Split(Fil, "\")
WordCount = UBound(fileName())
Worksheets("Measurement_Info").Range("D" & LR).value = test_date
Worksheets("Measurement_Info").Range("A" & LR).value = fileName(WordCount)
Worksheets("Measurement_Info").Range("I" & LR).value = Replace(Device, """", "")
Worksheets("Measurement_Info").Range("B" & LR).value = ST
Worksheets("Measurement_Info").Range("C" & LR).value = Sp
'Sheets("Measurement_Info").Columns(2).NumberFormat = "hh:mm:ss.000"
'Sheets("Measurement_Info").Columns(3).NumberFormat = "hh:mm:ss.000"


Call TurnOnStuff
excelSettingsOff = False


AA = fileName(WordCount)
    Name = Split(AA, " ")
    TownName = Split(AA, " ")
    GetTownNameTwo = TownName(2)
    GetSaveName = Split(AA, TownName(1) & " ")
    SaveName = Split(GetSaveName(1), " CST")
    GetSaveZero = SaveName(0)
    Name = Split(AA, Name(1))
    Name = Split(Name(1), " CST")
    Dayy = Split(AA, " CST")
    Dayy = Split(Dayy(1), ".")
    GetDayyZero = Dayy(0)
    Sheets("LTE_4GIDLE_MultiMetric_5").Select
    Sheets("LTE_4GIDLE_MultiMetric_5").Columns(11).NumberFormat = "hh:mm:ss.000"
    Tweet

DirPathGlobal = DirPath

    If passNumber = 1 And isLastForLocation = True Then
        If FinalTime = True Then
        ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & Keep(0) & "\COVERAGE\" & Keep(0) & " " & TownName(2) & " " & Keep(1) & " " & "COVERAGE" & Dayy(0) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Sleep 6000  ' 3 second delay
        SaveTelcoWorkbook "", DirPath
        Sleep 6000  ' 3 second delay
        Code
        ActiveWorkbook.Save
        Sleep 6000  ' 3 second delay
        Copy_3GCov_TO_Results
        Sleep 6000  ' 3 second delay
        Ending = CurrentRow
'        Call DeleteFirstLocRows
        workbookName = Keep(0) & " " & TownName(2) & " " & Keep(1) & " " & "COVERAGE" & Dayy(0) & ".xlsm"
        Workbooks(workbookName).Activate
        Sheets(sh).Range("K2:S" & CurrentRow).ClearContents
        Sheets(sh).Range("AO1") = 2
        Sheets("Measurement_Info").Range("A2:D" & LR).ClearContents
        Sheets("Measurement_Info").Range("I" & LR).ClearContents

        End If
'       ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & Keep(0) & "\DATA\" & Keep(0) & Name(0) & " " & Keep(1) & " " & "DATA" & Dayy(0) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
        ElseIf passNumber = 2 Then
        
    
    End If
'AA = FileName(WordCount)
'     Dayy = Split(AA, " CST")
'     Dayy = Split(Dayy(1), ".")
'     FileName = Split(AA, " ")
'     FileName = Split(AA, FileName(1))
'     FileName = Split(FileName(1), " ")
'
'     Sheets("LTE_4GIDLE_MultiMetric_5").Select
'     Sheets("LTE_4GIDLE_MultiMetric_5").Columns(11).NumberFormat = "hh:mm:ss.000"
     
'     'Tweet
'     ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "NCA\" & Keep(0) & "\COVERAGE\" & Keep(0) & " " & FileName(1) & " " & Keep(1) & " COVERAGE" & Dayy(0) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
''     ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & Keep(0) & " " & FileName(1) & " " & Dayy(0) & " 4G COVERAGE.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled

'     Sheets("Measurement_Info").Select
End If
    Exit Sub

FatalError:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If Not KPI Is Nothing Then KPI.Close
    If excelSettingsOff Then TurnOnStuff
    MsgBox "FourGCov failed while processing:" & vbCrLf & CStr(Fil) & vbCrLf & _
           "Error " & errNumber & ": " & errDescription, vbExclamation, "FourGCov"
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
Dim Result As String
    
    Select Case Band
    
    Case Is = "70001"
    
        Result = "LTE FDD 2100"
        
    Case Is = "70002"
    
        Result = "LTE FDD 1900"
        
    Case Is = "70003"
    
        Result = "LTE FDD 1800"
        
    Case Is = "70004"
    
        Result = "LTE FDD 2100 AWS"
        
    Case Is = "70005"
    
        Result = "LTE FDD 850"
        
    Case Is = "70006"
    
        Result = "LTE FDD 850"
        
    Case Is = "70007"
    
        Result = "LTE FDD 2600"
        
    Case Is = "70008"
    
        Result = "LTE FDD 900"
        
    Case Is = "70009"
    
        Result = "LTE FDD 1800"
        
    Case Is = "70010"
    
        Result = "LTE FDD 2100"
        
    Case Is = "70011"
    
        Result = "LTE FDD 1400"
    
    Case Is = "70012"
    
        Result = "LTE FDD 700"
        
    Case Is = "70013"
    
        Result = "LTE FDD 700"
        
    Case Is = "70014"
    
        Result = "LTE FDD 700"
        
    Case Is = "70017"
    
        Result = "LTE FDD 700"
        
    Case Is = "70018"
    
        Result = "LTE FDD 850"
        
    Case Is = "70019"
    
        Result = "LTE FDD 850"
        
    Case Is = "70020"
    
        Result = "LTE FDD 800"
        
    Case Is = "70021"
    
        Result = "LTE FDD 1500"
        
    Case Is = "70022"
    
        Result = "LTE FDD 3500"
        
    Case Is = "70023"
    
        Result = "LTE FDD 2200"
        
    Case Is = "70024"
    
        Result = "LTE FDD 1500"
        
    Case Is = "70025"
    
        Result = "LTE FDD 1900"
        
    Case Is = "70026"
    
        Result = "LTE FDD 850"
        
    Case Is = "70027"
    
        Result = "LTE FDD 800"
        
    Case Is = "70028"
    
        Result = "LTE FDD 700"
        
    Case Is = "70029"
    
        Result = "LTE FDD 700"
        
    Case Is = "70030"
    
        Result = "LTE FDD 2350"
        
    Case Is = "70031"
    
        Result = "LTE FDD 450"
        
    Case Is = "70032"
    
        Result = "LTE FDD 1500 L"
        
    Case Is = "70064"
    
        Result = "LTE FDD 390-470"
        
    Case Is = "70065"
    
        Result = "LTE FDD 2100"
        
    Case Is = "70066"
    
        Result = "LTE FDD AWS-3 2100"
        
    Case Is = "70067"
    
        Result = "LTE FDD 700 EU"
        
    Case Is = "70068"
    
        Result = "LTE FDD 700 ME"
        
    Case Is = "70069"
    
        Result = "LTE FDD 2500"
        
    Case Is = "70070"
    
        Result = "LTE FDD AWS-4"
        
    Case Is = "70071"
    
        Result = "LTE FDD 600"
        
    Case Is = "70072"
    
        Result = "LTE FDD 450 PMR/PAMR"
        
    Case Is = "70073"
    
        Result = "LTE FDD 450 APAC"
        
    Case Is = "70074"
    
        Result = "LTE FDD 1500 USA L"
        
    Case Is = "70075"
    
        Result = "LTE FDD 1500 EU L"
        
    Case Is = "70076"
    
        Result = "LTE FDD 1500 ext EU L"
        
    Case Is = "70085"
    
        Result = "LTE FDD 700 a+"
        
    Case Is = "70087"
    
        Result = "LTE FDD 420-425"
        
    Case Is = "70088"
    
        Result = "LTE FDD 422-427"
        
    Case Is = "70240"
    
        Result = "LTE FDD 5154-5925"
        
    Case Is = "70250"
    
        Result = "LTE FDD 3550-3700"
        
    Case Is = "70252"
    
        Result = "LTE FDD 5200 NII-1"
        
    Case Is = "70255"
    
        Result = "LTE FDD 5700 NII-3"
        
    Case Is = "79999"
    
        Result = "LTE FDD"
        
    Case Is = "80033"
    
        Result = "LTE TDD 1900-1920"
        
    Case Is = "80034"
    
        Result = "LTE TDD 2010-2025"
        
    Case Is = "80035"
    
        Result = "LTE TDD 1850-1910"
        
    Case Is = "80036"
    
        Result = "LTE TDD 1930-1990"
        
    Case Is = "80037"
    
        Result = "LTE TDD 1910-1930"
        
    Case Is = "80038"
    
        Result = "LTE TDD 2570-2620"
        
    Case Is = "80039"
    
        Result = "LTE TDD 1880-1920"
        
    Case Is = "80040"
    
        Result = "LTE TDD 2300-2400"
        
    Case Is = "80041"
    
        Result = "LTE TDD 2496-2690"
        
    Case Is = "80042"
    
        Result = "LTE TDD 3400-3600"
        
    Case Is = "80043"
    
        Result = "LTE TDD 3600-3800"
        
    Case Is = "80044"
    
        Result = "LTE TDD 703-803"
        
    Case Is = "80045"
    
        Result = "LTE TDD 1447-1467"
        
    Case Is = "80046"
    
        Result = "LTE TDD 5154-5925"
        
    Case Is = "80047"
    
        Result = "LTE TDD 5855-5925"
        
    Case Is = "80048"
    
        Result = "LTE TDD 3550-3700"
        
    Case Is = "80049"
    
        Result = "LTE TDD 3550-3700"
        
    Case Is = "80050"
    
        Result = "LTE TDD 1432-1517"
        
    Case Is = "80051"
    
        Result = "LTE TDD 1427-1432"
        
    Case Is = "80052"
    
        Result = "LTE TDD 3300-3400"
    
    Case Is = "80053"
    
        Result = "LTE TDD 2483-2495"
        
    Case Is = "80061"
    
        Result = "LTE TDD 1447-1467"
        
    Case Is = "80062"
    
        Result = "LTE TDD 1785-1805"
        
    Case Is = "80087"
    
        Result = "LTE TDD 1447-1467"
        
    Case Is = "80088"
    
        Result = "LTE TDD 1785-1805"
        
    Case Is = "89999"
    
        Result = "LTE TDD"
    
    End Select
    
    getBand = Result
    
End Function
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
Private Function IsPositiveNumber(ByVal valueText As Variant) As Boolean
    If IsNumeric(valueText) Then
        IsPositiveNumber = (CDbl(valueText) > 0)
    End If
End Function
