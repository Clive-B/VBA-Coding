Attribute VB_Name = "ThreeGCovUI"
Option Base 1
' Add this at the top of your module (before any procedures)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public DirPathGlobal, workbookPathGlobal, GetKeepZero, GetKeepOne As String
Public GetSaveZero, GetDayyZero, EndRow, Beginning, Ending, GetTownNameTwo As String
Public prevtownDayKey As Variant ' Declare global variable
Public prevtown As String ' Declare global variable
'Sub ThreeGCov(Fil As String)
 Sub ThreeGCov(Fil As String, DirPath As String, Finish As Boolean, _
               isLastInTownDate As Boolean, isLastArrayElement As Boolean, _
               isLastForLocation As Boolean, FinalTime As Boolean, passNumber As Integer, _
               townDayKey As Variant, _
               Optional location As String = "")

Dim KPI As Scripting.TextStream
Dim ArryLine, cmrt, Alert, latitude, longitude, Get_date, test_date, CallAttempt, sh, Fgname, Band, Device, ST, Sp As String
Dim Keep() As String
Dim WordCount As Integer
Dim CurrentLine() As String
Dim fileName() As String
Dim Tdate() As String
Dim wb As Workbook
Dim ws As Worksheet
Dim CurrentRow, Tow, count As Integer
Dim fso As Scripting.FileSystemObject
Dim oWSHShell As Object
Dim lRow As Long
Dim Dayy() As String
Dim hasChanged As Boolean
Dim NameTown As String
Dim workbookName As String

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
            ThisWorkbook.Sheets(prevtown).Delete
            ThisWorkbook.Sheets("AVGCAL").Delete
            ThisWorkbook.Sheets("FILTERCov").Delete
            ThisWorkbook.Sheets("AdvFILTERCov").Delete
            Sheets.Add(After:=Sheets(Sheets.count)).Name = "AVGCAL"
            ThisWorkbook.Sheets("AVGCAL").Visible = xlSheetHidden
            Application.DisplayAlerts = True
            End If
            ElseIf CStr(prevtown) = CStr(NameTown) And prevtownDayKey <> townDayKey Then
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets(prevtown).Delete
            ThisWorkbook.Sheets("AVGCAL").Delete
            ThisWorkbook.Sheets("FILTERCov").Delete
            ThisWorkbook.Sheets("AdvFILTERCov").Delete
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
sh = "UMTS_3GIDLE_MultiMetric_5"

Sheets(sh).Select

CurrentRow = Range("AO1").value
Call TurnOffStuff
Do Until KPI.AtEndOfStream
    ArryLine = KPI.ReadLine
     CurrentLine = Split(ArryLine, ",")
     
     Select Case CurrentLine(0)
     
        Case Is = "#DL"
            Device = CurrentLine(3)
            Keep = Split(CurrentLine(3), " ")
            Keep(0) = Replace(Keep(0), """", "")
            Keep(1) = Replace(Keep(1), """", "")
            GetKeepOne = Keep(1)
            GetKeepZero = Keep(0)
        Case Is = "#START"
            ST = CurrentLine(1)
            Get_date = Mid(CurrentLine(3), 2, Len(CurrentLine(3)) - 2)
            Tdate = Split(Get_date, ".")
            test_date = Tdate(1) & "/" & Tdate(0) & "/" & Tdate(2)
        Case Is = "CELLMEAS"
                
                If UBound(CurrentLine) < 10 Or CurrentLine(5) = "0" Then
                CurrentRow = CurrentRow - 1
                GoTo CurrentLineIncrement
                ElseIf CurrentLine(12) = "" Then
                CurrentRow = CurrentRow - 1
                GoTo CurrentLineIncrement
                ElseIf CurrentLine(10) = "1" And CurrentLine(12) = "0" And CurrentLine(3) = "5" Or CurrentLine(3) = "6" Then
                Range("O" & CurrentRow) = CurrentLine(18)
                ElseIf CurrentLine(6) <> 3 Then
                CurrentRow = CurrentRow - 1
                GoTo CurrentLineIncrement
                ElseIf CurrentLine(11) <= 0 Then
                CurrentRow = CurrentRow - 1
                GoTo CurrentLineIncrement
                ElseIf CurrentLine(3) = "5" Or CurrentLine(3) = "6" Then
                Range("O" & CurrentRow) = getmax((ArryLine))
                End If
                If Range("O" & CurrentRow).value = "" Then
                CurrentRow = CurrentRow - 1
                GoTo CurrentLineIncrement
                Else
                Range("P" & CurrentRow) = CurrentLine(14)
                Range("K" & CurrentRow) = CurrentLine(1)
                Range("L" & CurrentRow) = test_date
                Range("M" & CurrentRow) = latitude
                Range("N" & CurrentRow) = longitude
                Range("Q" & CurrentRow) = MNC
                Range("R" & CurrentRow) = MCC
                Band = CurrentLine(9)
                Range("S" & CurrentRow) = getBand((Band))
                End If
CurrentLineIncrement:
                CurrentRow = CurrentRow + 1
        Case Is = "GPS"
            longitude = CurrentLine(3)
            latitude = CurrentLine(4)
            
        Case Is = "SEI"
            MCC = CurrentLine(6)
            MNC = CurrentLine(7)
        
        Case Is = "#STOP"
            Sp = CurrentLine(1)
            
            Exit Do
    
    End Select
    
Loop

KPI.Close

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
    Sheets("UMTS_3GIDLE_MultiMetric_5").Select
'    Sheets("UMTS_3GIDLE_MultiMetric_5").Columns(11).NumberFormat = "hh:mm:ss.000"
    Tweet


DirPathGlobal = DirPath

    If passNumber = 1 And isLastInTownDate = True Then
        If FinalTime = True Then
        ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & Keep(0) & "\COVERAGE\" & Keep(0) & " " & TownName(2) & " " & Keep(1) & " " & "COVERAGE" & Dayy(0) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Sleep 6000  ' 3 second delay
        SaveTelcoWorkbook
        Sleep 6000  ' 3 second delay
        Code
        ActiveWorkbook.Save
        Sleep 6000  ' 3 second delay
        Copy_3GCov_TO_Results
        Sleep 6000  ' 3 second delay
'        Ending = CurrentRow
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
Else
End If
'    If Finish = "True" Then
'    ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & Keep(0) & "\COVERAGE\" & Keep(0) & " " & FileName(1) & " " & Keep(1) & " COVERAGE" & Dayy(0) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
'    DirPathGlobal = DirPath
'    Sleep 6000  ' 3 second delay
'    SaveTelcoWorkbook
'    Sleep 6000  ' 3 second delay
'    Copy_3GCov_TO_Results
'    Sleep 6000  ' 3 second delay
'    Save
'    Sleep 6000  ' 3 second delay
'    If Last = False Then
''        Clear
'    Else
'        If Last = True And isLastArrayElement = True Then
'        For Each WB In Workbooks
'        If WB.FullName = GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & Keep(0) & "\COVERAGE\" & Keep(0) & " " & FileName(1) & " " & Keep(1) & " COVERAGE" & Dayy(0) & ".xlsm" Then
'            WB.Close SaveChanges:=False
'            Exit For
'        End If
'        Next WB
'        End If
'    End If
'
'End If
     'ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & "NCA\" & Keep(0) & "\COVERAGE\" & Keep(0) & " " & FileName(1) & " " & Keep(1) & " COVERAGE" & Dayy(0) & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled

End Sub
Function getmax(arrayB As String) As String
    Dim arrayspace As Long, count As Long
    count = 0
    Dim newArray() As String
    newArray = Split(arrayB, ",")
    
    ' Validate array has enough elements
    If UBound(newArray) < 10 Then
        getmax = "Error: Invalid input array"
        Exit Function
    End If
    
    ' Safely get arrayspace with error handling
    If IsNumeric(newArray(10)) Then
        arrayspace = CLng(newArray(10))
    Else
        getmax = "Error: arrayspace is not numeric"
        Exit Function
    End If
    
    ' Additional validation for arrayspace size
    If arrayspace <= 0 Or arrayspace > 10000 Then ' Adjust max limit as needed
        getmax = "Error: Invalid arrayspace size"
        Exit Function
    End If
    
    Dim mycollect As New Collection
    Dim i As Long, j As Long
    i = 12
    
    For j = 1 To arrayspace
        ' Calculate current index safely with Long data type
        Dim currentIndex As Long
        If i = 12 Then
            currentIndex = i
        Else
            ' Check for potential overflow before calculation
            If count > (2147483647 - 12) / 17 Then ' Max Long value safety check
                getmax = "Error: Index calculation overflow"
                Exit Function
            End If
            currentIndex = (count * 17) + 12
        End If
        
        ' Check if index is within bounds
        If currentIndex <= UBound(newArray) And currentIndex >= 0 Then
            If newArray(currentIndex) = "0" Then
                ' Calculate value index and check bounds
                Dim valueIndex As Long
                valueIndex = currentIndex + 6
                
                If valueIndex <= UBound(newArray) And valueIndex >= 0 Then
                    ' Safely convert to double with validation
                    If IsNumeric(newArray(valueIndex)) And newArray(valueIndex) <> "" Then
                        On Error Resume Next ' Handle any conversion errors
                        Dim dblValue As Double
                        dblValue = CDbl(newArray(valueIndex))
                        If Err.Number = 0 Then
                            mycollect.Add dblValue
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
        
        i = i + 1
        count = count + 1
    Next j
    
    ' Return result
    If mycollect.count > 0 Then
        getmax = CStr(GetMaxFromCollection(mycollect))
    Else
        getmax = "No valid data found"
    End If
End Function
Function GetMaxFromCollection(col As Collection) As Double
    Dim maxVal As Double
    Dim i As Long
    
    If col.count > 0 Then
        maxVal = col(1)
        For i = 2 To col.count
            If col(i) > maxVal Then
                maxVal = col(i)
            End If
        Next i
    End If
    
    GetMaxFromCollection = maxVal
End Function
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
Public Function CollectionToArray(myCol As Collection) As Variant

    Dim Result  As Variant
    Dim cnt    As Double

    ReDim Result(myCol.count) As Variant
    For cnt = 0 To myCol.count - 1
        Result(cnt + 1) = myCol(cnt + 1)
    Next cnt
    CollectionToArray = Result

End Function
Function getBand(Band As String) As String
Dim Result As String
    
    Select Case Band
    
    Case Is = "50001"
    
        Result = "UMTS FDD 2100"
        
    Case Is = "50002"
    
        Result = "UMTS FDD 1900"
        
    Case Is = "50003"
    
        Result = "UMTS FDD 1800"
        
    Case Is = "50004"
    
        Result = "UMTS FDD 2100 AWS"
        
    Case Is = "50005"
    
        Result = "UMTS FDD 850"
        
    Case Is = "50006"
    
        Result = "UMTS FDD 850"
        
    Case Is = "50007"
    
        Result = "UMTS FDD 2600"
        
    Case Is = "50008"
    
        Result = "UMTS FDD 900"
        
    Case Is = "50009"
    
        Result = "UMTS FDD 1800"
        
    Case Is = "500010"
    
        Result = "UMTS FDD 2100"
        
    Case Is = "500011"
    
        Result = "UMTS FDD 1400"
    
    Case Is = "500012"
    
        Result = "UMTS FDD 700"
        
    Case Is = "500013"
    
        Result = "UMTS FDD 700"
        
    Case Is = "500014"
    
        Result = "UMTS FDD 700"
        
    Case Is = "500019"
    
        Result = "UMTS FDD 850"
        
    Case Is = "500020"
    
        Result = "UMTS FDD 800"
        
    Case Is = "500021"
    
        Result = "UMTS FDD 1500"
        
    Case Is = "500022"
    
        Result = "UMTS FDD 3500"
        
    Case Is = "500025"
    
        Result = "UMTS FDD 1900"
        
    Case Is = "500026"
    
        Result = "UMTS FDD 850"
        
    Case Is = "59999"
    
        Result = "UMTS FDD"
        
    Case Is = "60001"
    
        Result = "UMTS TD-SCDMA 2000"
        
    Case Is = "60002"
    
        Result = "UMTS TD-SCDMA 1900"
        
    Case Is = "60003"
    
        Result = "UMTS TD-SCDMA 1900"
        
    Case Is = "60004"
    
        Result = "UMTS TD-SCDMA 2600"
        
    Case Is = "60005"
    
        Result = "UMTS TD-SCDMA 1900"
        
    Case Is = "60006"
    
        Result = "UMTS TD-SCDMA 2300"
        
    Case Is = "69999"
    
        Result = "UMTS TD-SCDMA"
    
    End Select
    
    getBand = Result
    
End Function
