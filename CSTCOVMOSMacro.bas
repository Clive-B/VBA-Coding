Attribute VB_Name = "CSTCOVMOSMacro"
' Add this at the top of your module (before any procedures)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
Public DirPathGlobal As String, workbookPathGlobal As String ' Declare global variable
Public SingleCSTOutputEnabled As Boolean
Public SingleCSTOutputFolder As String
Public SingleCSTOutputName As String
Private TopologyCacheBuiltFor As String
Private TopologySiteNames() As String
Private TopologyCellNames() As String
Private TopologyCellIds() As String
Private TopologyTechs() As String
Private TopologyAzimuths() As Double
Private TopologyLats() As Double
Private TopologyLons() As Double
Private TopologyCount As Long
Private TopologyGpsMatchCache As Object
Private ServingTimeCacheWb As String
Private ServingTimes() As Double
Private ServingRows() As Long
Private ServingTimeCount As Long
Private MosTimeCacheWb As String
Private MosTimes() As Double
Private MosRows() As Long
Private MosTimeCount As Long
Private CstTimeCacheWb As String
Private CstTimes() As Double
Private CstRows() As Long
Private CstTimeCount As Long

Public Sub ConfigureSingleCSTOutput(ByVal outputFolder As String, ByVal outputName As String, Optional ByVal enabled As Boolean = True)
    SingleCSTOutputFolder = Trim$(outputFolder)
    SingleCSTOutputName = Trim$(outputName)
    SingleCSTOutputEnabled = enabled
End Sub

Public Sub ClearSingleCSTOutput()
    SingleCSTOutputEnabled = False
    SingleCSTOutputFolder = ""
    SingleCSTOutputName = ""
End Sub

Private Sub WriteMosOrigRadioInfo(ByVal rowNo As Long, ByVal channelNo As String, ByVal bandInfo As String, ByVal callType As String, Optional ByVal networkName As String = "")
    Dim dlFreq As Variant
    Dim ulFreq As Variant

    GetChannelFrequency networkName, channelNo, bandInfo, dlFreq, ulFreq
    With Sheets("LTE_POLQA_MOS_MultiMetric_5")
        If Not IsEmpty(dlFreq) Then .Range("T" & rowNo).value = dlFreq
        If Not IsEmpty(ulFreq) Then .Range("U" & rowNo).value = ulFreq
        .Range("V" & rowNo).value = bandInfo
        .Range("X" & rowNo).value = callType
    End With
End Sub

Private Sub WriteMosTermRadioInfo(ByVal rowNo As Long, ByVal channelNo As String, ByVal bandInfo As String, ByVal callType As String, Optional ByVal networkName As String = "")
    Dim dlFreq As Variant
    Dim ulFreq As Variant

    GetChannelFrequency networkName, channelNo, bandInfo, dlFreq, ulFreq
    With Sheets("LTE_POLQA_MOS_MultiMetric_5")
        If Trim$(CStr(.Range("T" & rowNo).value)) = "" And Not IsEmpty(dlFreq) Then .Range("T" & rowNo).value = dlFreq
        If Trim$(CStr(.Range("U" & rowNo).value)) = "" And Not IsEmpty(ulFreq) Then .Range("U" & rowNo).value = ulFreq
        .Range("W" & rowNo).value = bandInfo
        .Range("X" & rowNo).value = callType
    End With
End Sub

Private Sub GetChannelFrequency(ByVal networkName As String, ByVal channelNo As String, ByVal bandInfo As String, ByRef dlFreq As Variant, ByRef ulFreq As Variant)
    Dim ch As Double
    Dim bandText As String
    Dim netText As String

    dlFreq = Empty
    ulFreq = Empty
    If Not IsNumeric(channelNo) Then Exit Sub

    ch = CDbl(channelNo)
    bandText = UCase$(Trim$(NormalizeBandInfo(bandInfo)))
    netText = UCase$(Trim$(networkName))

    If InStr(1, bandText, "UMTS", vbTextCompare) > 0 Or ch >= 10000 Then
        If InStr(1, bandText, "900", vbTextCompare) > 0 Or (ch >= 2937 And ch <= 3088) Then
            dlFreq = Round((ch / 5#) + 340#, 1)
            ulFreq = Round(CDbl(dlFreq) - 45#, 1)
        Else
            dlFreq = Round(ch / 5#, 1)
            ulFreq = Round(CDbl(dlFreq) - 190#, 1)
        End If
    ElseIf InStr(1, bandText, "LTE", vbTextCompare) > 0 Then
        If InStr(1, bandText, "B20", vbTextCompare) > 0 Or (ch >= 6150 And ch <= 6449) Then
            dlFreq = Round(791# + (0.1 * (ch - 6150#)), 1)
            ulFreq = Round(CDbl(dlFreq) + 41#, 1)
        ElseIf InStr(1, bandText, "B7", vbTextCompare) > 0 Or (ch >= 2750 And ch <= 3449) Then
            dlFreq = Round(2620# + (0.1 * (ch - 2750#)), 1)
            ulFreq = Round(CDbl(dlFreq) - 120#, 1)
        ElseIf InStr(1, bandText, "B3", vbTextCompare) > 0 Or (ch >= 1200 And ch <= 1949) Then
            dlFreq = Round(1805# + (0.1 * (ch - 1200#)), 1)
            ulFreq = Round(CDbl(dlFreq) - 95#, 1)
        ElseIf InStr(1, bandText, "B40", vbTextCompare) > 0 Or (ch >= 38650 And ch <= 39649) Then
            dlFreq = Round(2300# + (0.1 * (ch - 38650#)), 1)
            ulFreq = dlFreq
        ElseIf InStr(1, bandText, "B41", vbTextCompare) > 0 Or (ch >= 39650 And ch <= 41589) Then
            dlFreq = Round(2496# + (0.1 * (ch - 39650#)), 1)
            ulFreq = dlFreq
        End If
    Else
        GetGsmFrequency netText, ch, bandText, dlFreq, ulFreq
    End If
End Sub

Private Sub GetGsmFrequency(ByVal networkName As String, ByVal channelNo As Double, ByVal bandInfo As String, ByRef dlFreq As Variant, ByRef ulFreq As Variant)
    If InStr(1, bandInfo, "1800", vbTextCompare) > 0 Or channelNo >= 512 Then
        dlFreq = Round(1805.2 + (0.2 * (channelNo - 512#)), 1)
        ulFreq = Round(CDbl(dlFreq) - 95#, 1)
    ElseIf channelNo >= 0 And channelNo <= 124 Then
        If channelNo <= 25 And networkName = "AIRTELTIGO" Then
            dlFreq = Round(935# + (0.2 * channelNo), 1)
        Else
            dlFreq = Round(935# + (0.2 * channelNo), 1)
        End If
        ulFreq = Round(CDbl(dlFreq) - 45#, 1)
    End If
End Sub

Public Sub CSTCOVMOS(fil, element, DirPath As String, Finish As Boolean, Last As Boolean, Optional isLastArrayElement As Boolean)

'Sub CSTCOVMOS()

Dim KPI As Scripting.TextStream
Dim ArryLine As String, cmrt As String, Alert As String, latitude As String, longitude As String, Get_date As String, test_date As String, CallAttempt As String, sh As String, fgname As String
Dim ChNum As String
Dim Keep() As String
Dim fileName() As String
Dim Dayy() As String
Dim Checkk() As String
Dim Gett() As String
Dim LR As Long
Dim CurrentLine() As String
Dim Tdate() As String
Dim wb As Workbook
Dim ws As Worksheet
Dim CurrentRow As Long, CurrentRowMos As Long, CurrentRowRscp As Long, WordCount As Long, CurrentR As Long, CurrentRowGSM As Long, CallNo As Long, CurrentRowSeq As Long
Dim Check As Integer
Dim prt As String, areqt As String, arest As String, sut As String, cpt As String, cct As String, pt As String, ct As String, cat As String, dt As String, rt As String, rct As String, RSCPRxlev As String, workbookPath As String, file As String
Dim wbPath As String
Dim savePath As String
Dim Tow As Long
Dim fso As Scripting.FileSystemObject
Dim lRow As Long
Dim startTime As Date
Dim endTime As Date
Dim timeDifference As Double
Dim oWSHShell As Object
Dim AllCh(500000) As String
Dim outputArray() As String
Dim item As Variant
Dim BandInfo3G As String, BandInfo2G As String, RSCP As String, RXlev As String, Unitime As String
Dim latestChiTime As String, latestSystem As String, latestBand As String, latestRrcState As String, latestChannel As String, latestCellId As String
Dim servingRow As Long, handoverRow As Long
Dim excelSettingsOff As Boolean
Dim errNumber As Long, errDescription As String
'DirPath = "February 2025\"

On Error GoTo FatalError

Set oWSHShell = CreateObject("WScript.Shell")
GetDesktop = oWSHShell.SpecialFolders("Desktop")
GetDesktop = GetDesktop & "\"
Set oWSHShell = Nothing

Set fso = New Scripting.FileSystemObject

'Fil = "C:\Users\LENOVO\Desktop\QoS Automation\March 2025\Telcos\MTN\CST\LOGS\25Feb18 144809 POTSIN CST DAY 1.4.nmf"

'If objFSO.FolderExists(folderPath) Then
'            ' Get the folder object
'            Set objFolder = objFSO.GetFolder(folderPath)

fgname = fil

file = GetFileNameOnly(fgname)

Checkk = Split(file, " ")

Gett = Split(file, Checkk(1))

If Dir((GetDesktop & "QoS Automation\" & DirPath & "\NCA\" & element & "\CST\" & element & Gett(1) & ".xlsm")) = "" Then


    'Fgname = GetFileNameFromPath
    
    Set KPI = fso.OpenTextFile(fil)
    
    workbookPath = FindworkbookPath()
    ' Loop through open workbooks to find the one matching the path
        For Each wb In Workbooks
            If wb.fullName = workbookPath Then
                wb.Activate ' Activate the workbook
                Windows(wb.Name).Activate ' Bring its window to the front
                Exit For
            End If
        Next wb
    
    
    Sheets("LTE_POLQA_MOS_MultiMetric_5").Select
'    With Sheets("LTE_POLQA_MOS_MultiMetric_5")
'        .Range("R1").value = "ARFCN Originate"
'        .Range("S1").value = "ARFCN Terminate"
'        .Range("T1").value = "DL Frequency"
'        .Range("U1").value = "UL Frequency"
'        .Range("V1").value = "Band Info Originate"
'        .Range("W1").value = "Band Info Terminate"
'        .Range("X1").value = "Call Type"
'    End With
    
    CurrentRow = Range("AO1").value
    CurrentRowMos = Range("AO1").value
    CurrentRowRscp = Range("AO1").value
    CurrentR = Range("AO1").value
    CurrentRowGSM = Range("AO1").value
    CurrentRowSeq = 2
    servingRow = 2
    handoverRow = 2
    CallNo = 1
    Check = CallNo + 1
    InitializeCellHandoverInspectionSheets servingRow, handoverRow
    
    Call TurnOffStuff
    excelSettingsOff = True
    Do Until KPI.AtEndOfStream
        ArryLine = KPI.ReadLine
         CurrentLine = Split(ArryLine, ",")
         
         Select Case CurrentLine(0)
         
            Case Is = "#DL"
                Device = CurrentLine(3)
                Keep = Split(CurrentLine(3), " ")
                Keep(0) = Replace(Keep(0), """", "")
                Keep(1) = Replace(Keep(1), """", "")
            Case Is = "#START"
                ST = CurrentLine(1)
                Get_date = Mid(CurrentLine(3), 2, Len(CurrentLine(3)) - 2)
                Tdate = Split(Get_date, ".")
                test_date = Tdate(1) & "/" & Tdate(0) & "/" & Tdate(2)
            Case Is = "CAA"
                If CurrentLine(4) = "5" And CurrentLine(5) = "1" Then
                    CallAttempt = CurrentLine(1)
                End If
                
                If CurrentLine(4) = "1" And CurrentLine(5) = "1" Then
                    CallAttempt = CurrentLine(1)
                End If
            Case Is = "CHI"
                UpdateLatestCellIdentity CurrentLine, latestChiTime, latestSystem, latestBand, latestRrcState, latestChannel, latestCellId
                
            Case Is = "L3SM"
                
                '3G Technology
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """CM_SERVICE_REQUEST""" Then
                    cmrt = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("GSM_VoiceCST_multimetric_6").Range("P" & CurrentRow) = ChNum
                    Check = CallNo + 1
                    If Check = CurrentRowSeq Then
                    Else
                        CurrentRowSeq = Check
                    End If
                    Sheets("CALL_SEQUENCE_MOC").Range("K" & CurrentRowSeq) = CallNo
                    Sheets("CALL_SEQUENCE_MOC").Range("L" & CurrentRowSeq) = test_date
                    Sheets("CALL_SEQUENCE_MOC").Range("M" & CurrentRowSeq) = latitude
                    Sheets("CALL_SEQUENCE_MOC").Range("N" & CurrentRowSeq) = longitude
                    Sheets("CALL_SEQUENCE_MOC").Range("O" & CurrentRowSeq) = BandInfo3G
                    Sheets("CALL_SEQUENCE_MOC").Range("Q" & CurrentRowSeq) = cmrt
                    CallNo = CallNo + 1
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """PAGING_RESPONSE""" Then
                    prt = Format(CurrentLine(1), "hh:mm:ss.000")
                    Check = CallNo + 1
                    If Check = CurrentRowSeq Then
                    Else
                        CurrentRowSeq = Check
                    End If
                    Sheets("CALL_SEQUENCE_MTC").Range("K" & CurrentRowSeq) = CallNo
                    Sheets("CALL_SEQUENCE_MTC").Range("L" & CurrentRowSeq) = test_date
                    Sheets("CALL_SEQUENCE_MTC").Range("M" & CurrentRowSeq) = latitude
                    Sheets("CALL_SEQUENCE_MTC").Range("N" & CurrentRowSeq) = longitude
                    Sheets("CALL_SEQUENCE_MTC").Range("O" & CurrentRowSeq) = BandInfo3G
                    Sheets("CALL_SEQUENCE_MTC").Range("Q" & CurrentRowSeq) = prt
                    CallNo = CallNo + 1
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "2" And CurrentLine(5) = """AUTHENTICATION_REQUEST""" Then
                    areqt = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("CALL_SEQUENCE_MTC").Range("Q" & CurrentRowSeq) = areqt
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """AUTHENTICATION_RESPONSE""" Then
                    arest = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("CALL_SEQUENCE_MTC").Range("R" & CurrentRowSeq) = arest
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """SETUP""" Then
                    sut = Format(CurrentLine(1), "hh:mm:ss.000")
                    If Keep(1) = "MTC" Then
                    Else
                        If sut <> "" And cmrt <> "" Then
                            endTime = CDate(Left(sut, 8)) + CDbl(Mid(sut, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("R" & CurrentRowSeq) = timeDifference
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "2" And CurrentLine(5) = """SETUP""" Then
                    sut = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If sut <> "" And prt <> "" Then
                            endTime = CDate(Left(sut, 8)) + CDbl(Mid(sut, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("T" & CurrentRowSeq) = timeDifference
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """IDENTITY_REQUEST""" Then
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """IDENTITY_RESPONSE""" Then
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "2" And CurrentLine(5) = """CALL_PROCEEDING""" Then
                    cpt = Format(CurrentLine(1), "hh:mm:ss.000")
                    If Keep(1) = "MTC" Then
                    Else
                        If cpt <> "" And cmrt <> "" Then
                            endTime = CDate(Left(cpt, 8)) + CDbl(Mid(cpt, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("S" & CurrentRowSeq) = timeDifference
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """CALL_CONFIRMED""" Then
                    cct = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If cct <> "" And prt <> "" Then
                            endTime = CDate(Left(cct, 8)) + CDbl(Mid(cct, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("U" & CurrentRowSeq) = timeDifference
                        End If
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """ALERTING""" Then
                    Alert = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If Alert <> "" And prt <> "" Then
                            endTime = CDate(Left(Alert, 8)) + CDbl(Mid(Alert, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("V" & CurrentRowSeq) = timeDifference
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "2" And CurrentLine(5) = """ALERTING""" Then
                    Alert = Format(CurrentLine(1), "hh:mm:ss.000")
                    endTime = CDate(Left(Alert, 8)) + CDbl(Mid(Alert, 10)) / 86400000
                    startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                    Sheets("GSM_VoiceCST_multimetric_6").Range("K" & CurrentRow) = CallAttempt
                    Sheets("GSM_VoiceCST_multimetric_6").Range("L" & CurrentRow) = test_date
                    Sheets("GSM_VoiceCST_multimetric_6").Range("M" & CurrentRow) = latitude
                    Sheets("GSM_VoiceCST_multimetric_6").Range("N" & CurrentRow) = longitude
                    timeDifference = endTime - startTime
                    Sheets("GSM_VoiceCST_multimetric_6").Range("O" & CurrentRow) = timeDifference
                    Sheets("CALL_SEQUENCE_MOC").Range("T" & CurrentRowSeq) = timeDifference
                    'Sheets("GSM_VoiceCST_multimetric_6").Range("P" & CurrentRow) = ChNum
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                    CurrentR = CurrentR + 1
                    CurrentRow = CurrentRow + 1
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """PROGRESS""" Then
                    pt = Format(CurrentLine(1), "hh:mm:ss.000")
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """CONNECT""" Then
                    ct = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If ct <> "" And prt <> "" Then
                            endTime = CDate(Left(ct, 8)) + CDbl(Mid(ct, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("W" & CurrentRowSeq) = timeDifference
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "2" And CurrentLine(5) = """CONNECT""" Then
                    ct = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MTC" Then
                    Else
                        If ct <> "" And cmrt <> "" Then
                            endTime = CDate(Left(ct, 8)) + CDbl(Mid(ct, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("U" & CurrentRowSeq) = timeDifference
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """CONNECT_ACKNOWLEDGE""" Then
                    cat = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MTC" Then
                    Else
                        If cat <> "" And cmrt <> "" Then
                            endTime = CDate(Left(cat, 8)) + CDbl(Mid(cat, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("V" & CurrentRowSeq) = timeDifference
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "2" And CurrentLine(5) = """CONNECT_ACKNOWLEDGE""" Then
                    cat = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If cat <> "" And prt <> "" Then
                            endTime = CDate(Left(cat, 8)) + CDbl(Mid(cat, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("X" & CurrentRowSeq) = timeDifference
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """DISCONNECT""" Then
                    dt = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("GSM_VoiceCST_multimetric_6").Range("Q" & (CurrentRow - 1)) = ChNum
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
                    If Keep(1) = "MTC" Then
                    Else
                        If cmrt <> "" And dt <> "" Then
                            endTime = CDate(Left(dt, 8)) + CDbl(Mid(dt, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("W" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MOC").Range("P" & CurrentRowSeq).value = BandInfo3G
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                       End If
                    End If
                    
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "2" And CurrentLine(5) = """DISCONNECT""" Then
                    dt = Format(CurrentLine(1), "hh:mm:ss.000")
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
'                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If dt <> "" And prt <> "" Then
                            endTime = CDate(Left(dt, 8)) + CDbl(Mid(dt, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("Y" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MTC").Range("P" & CurrentRowSeq).value = BandInfo3G
                        End If
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """RELEASE""" Then
                    rt = Format(CurrentLine(1), "hh:mm:ss.000")
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
'                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If rt <> "" And prt <> "" Then
                            endTime = CDate(Left(rt, 8)) + CDbl(Mid(rt, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("Z" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MTC").Range("P" & CurrentRowSeq).value = BandInfo3G
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "2" And CurrentLine(5) = """RELEASE""" Then
                    rt = Format(CurrentLine(1), "hh:mm:ss.000")
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("GSM_VoiceCST_multimetric_6").Range("Q" & (CurrentRow - 1)) = ChNum
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
                    If Keep(1) = "MTC" Then
                    Else
                        If rt <> "" And cmrt <> "" Then
                            endTime = CDate(Left(rt, 8)) + CDbl(Mid(rt, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("X" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MOC").Range("P" & CurrentRowSeq).value = BandInfo3G
                            End If
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                End If
                
                'MOC
                If CurrentLine(3) = "5" And CurrentLine(4) = "1" And CurrentLine(5) = """RELEASE_COMPLETE""" Then
                    rct = Format(CurrentLine(1), "hh:mm:ss.000")
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("GSM_VoiceCST_multimetric_6").Range("Q" & (CurrentRow - 1)) = ChNum
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
                    If Keep(1) = "MTC" Then
                    Else
                    endTime = CDate(Left(rct, 8)) + CDbl(Mid(rct, 10)) / 86400000
                    startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                    timeDifference = endTime - startTime
                    Sheets("CALL_SEQUENCE_MOC").Range("Y" & CurrentRowSeq) = timeDifference
                    Sheets("CALL_SEQUENCE_MOC").Range("P" & CurrentRowSeq).value = BandInfo3G
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "5" And CurrentLine(4) = "2" And CurrentLine(5) = """RELEASE_COMPLETE""" Then
                    rct = Format(CurrentLine(1), "hh:mm:ss.000")
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
'                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If rct <> "" And prt <> "" Then
                            endTime = CDate(Left(rct, 8)) + CDbl(Mid(rct, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("AA" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MTC").Range("P" & CurrentRowSeq).value = BandInfo3G
                        End If
                    End If
                End If
                
                '2G Technology
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """CM_SERVICE_REQUEST""" Then
                    cmrt = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("GSM_VoiceCST_multimetric_6").Range("P" & CurrentRow) = ChNum
                    Check = CallNo + 1
                    If Check = CurrentRowSeq Then
                    Else
                        CurrentRowSeq = Check
                    End If
                    Sheets("CALL_SEQUENCE_MOC").Range("K" & CurrentRowSeq) = CallNo
                    Sheets("CALL_SEQUENCE_MOC").Range("L" & CurrentRowSeq) = test_date
                    Sheets("CALL_SEQUENCE_MOC").Range("M" & CurrentRowSeq) = latitude
                    Sheets("CALL_SEQUENCE_MOC").Range("N" & CurrentRowSeq) = longitude
                    Sheets("CALL_SEQUENCE_MOC").Range("O" & CurrentRowSeq) = BandInfo2G
                    Sheets("CALL_SEQUENCE_MOC").Range("Q" & CurrentRowSeq) = cmrt
                    CallNo = CallNo + 1
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """PAGING_RESPONSE""" Then
                    prt = Format(CurrentLine(1), "hh:mm:ss.000")
                    If Check = CurrentRowSeq Then
                    Else
                        CurrentRowSeq = Check
                    End If
                    Sheets("CALL_SEQUENCE_MTC").Range("K" & CurrentRowSeq) = CallNo
                    Sheets("CALL_SEQUENCE_MTC").Range("L" & CurrentRowSeq) = test_date
                    Sheets("CALL_SEQUENCE_MTC").Range("M" & CurrentRowSeq) = latitude
                    Sheets("CALL_SEQUENCE_MTC").Range("N" & CurrentRowSeq) = longitude
                    Sheets("CALL_SEQUENCE_MTC").Range("O" & CurrentRowSeq) = BandInfo2G
                    Sheets("CALL_SEQUENCE_MTC").Range("Q" & CurrentRowSeq) = prt
                    CallNo = CallNo + 1
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "2" And CurrentLine(5) = """AUTHENTICATION_REQUEST""" Then
                    areqt = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("CALL_SEQUENCE_MTC").Range("R" & CurrentRowSeq) = areqt
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """AUTHENTICATION_RESPONSE""" Then
                    arest = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("CALL_SEQUENCE_MTC").Range("S" & CurrentRowSeq) = arest
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """SETUP""" Then
                    sut = Format(CurrentLine(1), "hh:mm:ss.000")
                    sut = Format(CurrentLine(1), "hh:mm:ss.000")
                        If sut <> "" And cmrt <> "" Then
                            endTime = CDate(Left(sut, 8)) + CDbl(Mid(sut, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("R" & CurrentRowSeq) = timeDifference
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "2" And CurrentLine(5) = """SETUP""" Then
                    sut = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo2G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If sut <> "" And prt <> "" Then
                            endTime = CDate(Left(sut, 8)) + CDbl(Mid(sut, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("T" & CurrentRowSeq) = timeDifference
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """IDENTITY_REQUEST""" Then
                    'cmrt = Format(CurrentLine(1), "hh:mm:ss.000")
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """IDENTITY_RESPONSE""" Then
                    'cmrt = Format(CurrentLine(1), "hh:mm:ss.000")
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "2" And CurrentLine(5) = """CALL_PROCEEDING""" Then
                    cpt = Format(CurrentLine(1), "hh:mm:ss.000")
                        If cpt <> "" And cmrt <> "" Then
                            endTime = CDate(Left(cpt, 8)) + CDbl(Mid(cpt, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("S" & CurrentRowSeq) = timeDifference
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """CALL_CONFIRMED""" Then
                    cct = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo2G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If cct <> "" And prt <> "" Then
                            endTime = CDate(Left(cct, 8)) + CDbl(Mid(cct, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("U" & CurrentRowSeq) = timeDifference
                        End If
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """ALERTING""" Then
                    at = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo2G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If Alert <> "" And prt <> "" Then
                            endTime = CDate(Left(Alert, 8)) + CDbl(Mid(Alert, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("V" & CurrentRowSeq) = timeDifference
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "2" And CurrentLine(5) = """ALERTING""" Then
                    at = Format(CurrentLine(1), "hh:mm:ss.000")
                    Alert = Format(CurrentLine(1), "hh:mm:ss.000")
                    endTime = CDate(Left(Alert, 8)) + CDbl(Mid(Alert, 10)) / 86400000
                    startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                    Sheets("GSM_VoiceCST_multimetric_6").Range("K" & CurrentRow) = CallAttempt
                    Sheets("GSM_VoiceCST_multimetric_6").Range("L" & CurrentRow) = test_date
                    Sheets("GSM_VoiceCST_multimetric_6").Range("M" & CurrentRow) = latitude
                    Sheets("GSM_VoiceCST_multimetric_6").Range("N" & CurrentRow) = longitude
                    timeDifference = endTime - startTime
                    Sheets("GSM_VoiceCST_multimetric_6").Range("O" & CurrentRow) = timeDifference
                    Sheets("GSM_VoiceCST_multimetric_6").Range("O" & CurrentRow) = timeDifference
                    Sheets("CALL_SEQUENCE_MOC").Range("T" & CurrentRowSeq) = timeDifference
                    'Sheets("GSM_VoiceCST_multimetric_6").Range("P" & CurrentRow) = ChNum
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                    CurrentR = CurrentR + 1
                    'CurrentRow = CurrentRow + 1
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """PROGRESS""" Then
                    pt = Format(CurrentLine(1), "hh:mm:ss.000")
                    AllCh(Tow) = ChNum
                    Tow = Tow + 1
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """CONNECT""" Then
                    ct = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo2G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If ct <> "" And prt <> "" Then
                            endTime = CDate(Left(ct, 8)) + CDbl(Mid(ct, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("W" & CurrentRowSeq) = timeDifference
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "2" And CurrentLine(5) = """CONNECT""" Then
                    ct = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo2G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MTC" Then
                    Else
                        If ct <> "" And cmrt <> "" Then
                            endTime = CDate(Left(ct, 8)) + CDbl(Mid(ct, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("U" & CurrentRowSeq) = timeDifference
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """CONNECT_ACKNOWLEDGE""" Then
                    cat = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo2G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MTC" Then
                    Else
                        If cat <> "" And cmrt <> "" Then
                            endTime = CDate(Left(cat, 8)) + CDbl(Mid(cat, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("V" & CurrentRowSeq) = timeDifference
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "2" And CurrentLine(5) = """CONNECT_ACKNOWLEDGE""" Then
                    cat = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("R" & CurrentRowMos) = ChNum
                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo2G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If cat <> "" And prt <> "" Then
                            endTime = CDate(Left(cat, 8)) + CDbl(Mid(cat, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("X" & CurrentRowSeq) = timeDifference
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """DISCONNECT""" Then
                    dt = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("GSM_VoiceCST_multimetric_6").Range("Q" & (CurrentRow - 1)) = ChNum
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
'                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo2G, Keep(1), Keep(0)
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    If Keep(1) = "MTC" Then
                    Else
                        If dt <> "" And cmrt <> "" Then
                            endTime = CDate(Left(dt, 8)) + CDbl(Mid(dt, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("W" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MOC").Range("P" & CurrentRowSeq).value = BandInfo2G
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "2" And CurrentLine(5) = """DISCONNECT""" Then
                    dt = Format(CurrentLine(1), "hh:mm:ss.000")
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
'                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo2G, Keep(1), Keep(0)
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If dt <> "" And prt <> "" Then
                            endTime = CDate(Left(dt, 8)) + CDbl(Mid(dt, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("Y" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MTC").Range("P" & CurrentRowSeq).value = BandInfo2G
                        End If
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """RELEASE""" Then
                    rt = Format(CurrentLine(1), "hh:mm:ss.000")
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
'                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo2G, Keep(1), Keep(0)
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If rt <> "" And prt <> "" Then
                            endTime = CDate(Left(rt, 8)) + CDbl(Mid(rt, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("Z" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MTC").Range("P" & CurrentRowSeq).value = BandInfo2G
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "2" And CurrentLine(5) = """RELEASE""" Then
                    rt = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("GSM_VoiceCST_multimetric_6").Range("Q" & (CurrentRow - 1)) = ChNum
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
'                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo2G, Keep(1), Keep(0)
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    If Keep(1) = "MTC" Then
                    Else
                        If rt <> "" And cmrt <> "" Then
                            endTime = CDate(Left(rt, 8)) + CDbl(Mid(rt, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("X" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MOC").Range("P" & CurrentRowSeq).value = BandInfo2G
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                    End If
                End If
                
                'MOC
                If CurrentLine(3) = "1" And CurrentLine(4) = "1" And CurrentLine(5) = """RELEASE_COMPLETE""" Then
                    rct = Format(CurrentLine(1), "hh:mm:ss.000")
                    Sheets("GSM_VoiceCST_multimetric_6").Range("Q" & (CurrentRow - 1)) = ChNum
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
'                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo2G, Keep(1), Keep(0)
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    If Keep(1) = "MTC" Then
                    Else
                        If rct <> "" And cmrt <> "" Then
                            endTime = CDate(Left(rct, 8)) + CDbl(Mid(rct, 10)) / 86400000
                            startTime = CDate(Left(cmrt, 8)) + CDbl(Mid(cmrt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MOC").Range("Y" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MOC").Range("P" & CurrentRowSeq).value = BandInfo2G
                            AllCh(Tow) = ChNum
                            Tow = Tow + 1
                        End If
                    End If
                End If
                
                'MTC
                If CurrentLine(3) = "1" And CurrentLine(4) = "2" And CurrentLine(5) = """RELEASE_COMPLETE""" Then
                    rct = Format(CurrentLine(1), "hh:mm:ss.000")
                    If CurrentRowMos = 2 Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos) = Keep(1)
                    Else
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo3G, Keep(1), Keep(0)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    End If
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("S" & CurrentRowMos - 1) = ChNum
'                    WriteMosTermRadioInfo CurrentRowMos - 1, ChNum, BandInfo2G, Keep(1), Keep(0)
'                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("X" & CurrentRowMos - 1) = Keep(1)
                    If Keep(1) = "MOC" Then
                    Else
                        If rct <> "" And prt <> "" Then
                            endTime = CDate(Left(rct, 8)) + CDbl(Mid(rct, 10)) / 86400000
                            startTime = CDate(Left(prt, 8)) + CDbl(Mid(prt, 10)) / 86400000
                            timeDifference = endTime - startTime
                            Sheets("CALL_SEQUENCE_MTC").Range("AA" & CurrentRowSeq) = timeDifference
                            Sheets("CALL_SEQUENCE_MTC").Range("P" & CurrentRowSeq).value = BandInfo2G
                        End If
                    End If
                End If
                
            Case Is = "GPS"
                longitude = CurrentLine(3)
                latitude = CurrentLine(4)
                
            Case Is = "SHO"
                WriteHandoverInspectionEvent CurrentLine, test_date, latitude, longitude, latestCellId, handoverRow
                
            Case Is = "AQDL"
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("K" & CurrentRowMos) = CurrentLine(1)
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("L" & CurrentRowMos) = test_date
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("M" & CurrentRowMos) = latitude
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("N" & CurrentRowMos) = longitude
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("O" & CurrentRowMos) = CurrentLine(4)
                    If RSCPRxlev = "3G" Then
                        Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("P" & CurrentRowMos) = RSCP
    '                    If Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("T" & CurrentRowMos).Value = "" Then
    '                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo3G, Keep(1), Keep(0)
    '                    Else
    '                    End If
                    ElseIf RSCPRxlev = "2G" Then
                        Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("Q" & CurrentRowMos) = RXlev
    '                    If Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("T" & CurrentRowMos).Value = "" Then
    '                    WriteMosOrigRadioInfo CurrentRowMos, ChNum, BandInfo2G, Keep(1), Keep(0)
    '                    Else
    '                    End If
                    End If
                    'Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("P" & CurrentRowMos) = ChNum
                    CurrentRowMos = CurrentRowMos + 1
            Case Is = "CELLMEAS"
                    WriteServingCellInspectionRows CurrentLine, test_date, latitude, longitude, latestCellId, latestRrcState, servingRow
                    If Keep(1) = "MOC" Then
                    Sheets("UMTS_3GIDLE_MultiMetric_5").Select
                        If CurrentLine(3) = "5" Or CurrentLine(3) = "6" Then
                                If UBound(CurrentLine) < 10 Or CurrentLine(5) = "0" Then
                                    CurrentRowRscp = CurrentRowRscp - 1
                                    GoTo CurrentLineIncrement
                                    ElseIf CurrentLine(12) = "" Then
                                    CurrentRowRscp = CurrentRowRscp - 1
                                    GoTo CurrentLineIncrement
                                    ElseIf CurrentLine(10) = "1" And CurrentLine(12) = "0" And (CurrentLine(3) = "5" Or CurrentLine(3) = "6") Then
                                    Range("O" & CurrentRow) = CurrentLine(18)
                                    ElseIf CurrentLine(6) <> 3 Then
                                    CurrentRowRscp = CurrentRowRscp - 1
                                    GoTo CurrentLineIncrement
                                    ElseIf IsNumericAndLessOrEqualZero(CurrentLine(11)) Then
                                    CurrentRowRscp = CurrentRowRscp - 1
                                    GoTo CurrentLineIncrement
                                    ElseIf CurrentLine(3) = "5" Or CurrentLine(3) = "6" Then
                                    Range("O" & CurrentRowRscp) = getmax((ArryLine))
                                    RSCPRxlev = "3G"
                                    RSCP = Sheets("UMTS_3GIDLE_MultiMetric_5").Range("O" & CurrentRowRscp).value
                                End If
                                If Range("O" & CurrentRowRscp).value = "" Then
                                    CurrentRowRscp = CurrentRowRscp - 1
                                    GoTo CurrentLineIncrement
                                    Else
                                    ChNum = CurrentLine(14)
                                    Sheets("UMTS_3GIDLE_MultiMetric_5").Range("P" & CurrentRowRscp) = CurrentLine(14)
                                    Sheets("UMTS_3GIDLE_MultiMetric_5").Range("K" & CurrentRowRscp) = CurrentLine(1)
                                    Sheets("UMTS_3GIDLE_MultiMetric_5").Range("L" & CurrentRowRscp) = test_date
                                    Sheets("UMTS_3GIDLE_MultiMetric_5").Range("M" & CurrentRowRscp) = latitude
                                    Sheets("UMTS_3GIDLE_MultiMetric_5").Range("N" & CurrentRowRscp) = longitude
                                    Sheets("UMTS_3GIDLE_MultiMetric_5").Range("Q" & CurrentRowRscp) = MNC
                                    Sheets("UMTS_3GIDLE_MultiMetric_5").Range("R" & CurrentRowRscp) = MCC
                                    Band = CurrentLine(9)
                                    Sheets("UMTS_3GIDLE_MultiMetric_5").Range("S" & CurrentRowRscp) = getBand((Band))
                                    BandInfo3G = Sheets("UMTS_3GIDLE_MultiMetric_5").Range("S" & CurrentRowRscp).value
                                    Sheets("UMTS_3GIDLE_MultiMetric_5").Range("T" & CurrentRowRscp) = ChNum
                                End If
                        ElseIf CurrentLine(3) = "1" Then
                                    If UBound(CurrentLine) < 11 Then
                                        CurrentRowGSM = CurrentRowGSM - 1
                                        GoTo CurrentLineIncrementGSM
                                    
                                        ElseIf CurrentLine(11) = "" Then
                                            CurrentRowGSM = CurrentRowGSM - 1
                                            GoTo CurrentLineIncrementGSM
                                        
                                        ElseIf CurrentLine(3) <> 1 Or CurrentLine(7) <> 1 Then
                                        
                                            CurrentRowGSM = CurrentRowGSM - 1
                                            GoTo CurrentLineIncrementGSM
                                        ElseIf CurrentLine(3) = 1 And CurrentLine(7) = 1 Then
                                        ChNum = CurrentLine(9)
                                        Sheets("GSM_2G_MultiMetric_1").Range("O" & CurrentRowGSM) = CurrentLine(11)
                                        Sheets("GSM_2G_MultiMetric_1").Range("P" & CurrentRowGSM) = CurrentLine(9)
                                        Sheets("GSM_2G_MultiMetric_1").Range("K" & CurrentRowGSM) = CurrentLine(1)
                                        Sheets("GSM_2G_MultiMetric_1").Range("L" & CurrentRowGSM) = test_date
                                        Sheets("GSM_2G_MultiMetric_1").Range("M" & CurrentRowGSM) = latitude
                                        Sheets("GSM_2G_MultiMetric_1").Range("N" & CurrentRowGSM) = longitude
                                        Sheets("GSM_2G_MultiMetric_1").Range("Q" & CurrentRowGSM) = MNC
                                        Sheets("GSM_2G_MultiMetric_1").Range("R" & CurrentRowGSM) = MCC
                                        RSCPRxlev = "2G"
                                        RXlev = Range("O" & CurrentRowGSM).value
                                        Band = CurrentLine(8)
                                        Sheets("GSM_2G_MultiMetric_1").Range("S" & CurrentRowGSM) = getBand((Band))
                                        BandInfo2G = Range("S" & CurrentRowGSM).value
                        
                                        Else
                                        GoTo CurrentLineIncrementGSM
                                    End If
                        End If
                        
                    
    If CurrentLine(3) = "5" Or CurrentLine(3) = "6" Then
CurrentLineIncrement:
                            CurrentRowRscp = CurrentRowRscp + 1
    
    ElseIf CurrentLine(3) = "1" Then
CurrentLineIncrementGSM:
                            CurrentRowGSM = CurrentRowGSM + 1
    End If
                        
             ElseIf Keep(1) = "MTC" Then
                
                If CurrentLine(3) = "5" Or CurrentLine(3) = "6" Then
                    If UBound(CurrentLine) < 10 Then
                        Else
                        Band = CurrentLine(9)
                        BandInfo3G = getBand((Band))
                    End If
                    
                    If UBound(CurrentLine) < 14 Then
                        Else
                            ChNum = CurrentLine(14)
                    End If
                    
                    If CurrentLine(3) = "5" Or CurrentLine(3) = "6" Then
    '                RSCP = getmax((ArryLine))
    '                RSCPRxlev = "3G"
                    If UBound(CurrentLine) < 10 Or CurrentLine(5) = "0" Then
                    Else
                        If CurrentLine(12) = "" Then
                        ElseIf CurrentLine(10) = "1" And CurrentLine(12) = "0" And (CurrentLine(3) = "5" Or CurrentLine(3) = "6") Then
                            RSCP = CurrentLine(18)
                        ElseIf CurrentLine(6) <> 3 Then
                        ElseIf IsNumericAndLessOrEqualZero(CurrentLine(11)) Then
                        Else
                            RSCP = getmax((ArryLine))
                            RSCPRxlev = "3G"
                        End If
                        End If
                    End If
                
                ElseIf CurrentLine(3) = "1" Then
                        
                    If UBound(CurrentLine) < 11 Then
                    
                    ElseIf CurrentLine(3) = 1 And CurrentLine(7) = 1 Then
                    ChNum = CurrentLine(9)
                    Band = CurrentLine(8)
                    BandInfo2G = getBand((Band))
                    
                    End If
                    If UBound(CurrentLine) < 11 Then
                    ElseIf IsNumericAndLessOrEqualZero(CurrentLine(11)) Then
                    ElseIf CurrentLine(3) <> 1 Or CurrentLine(7) <> 1 Then
                    ElseIf CurrentLine(3) = 1 And CurrentLine(7) = 1 Then
                    RSCPRxlev = "2G"
                    RXlev = CurrentLine(11)
                    End If
                            
                End If
                
            End If
                    
            Case Is = "GPS"
            
            Case Is = "SEI"
                MCC = CurrentLine(6)
                MNC = CurrentLine(7)
            Case "HOA", "HOS", "HOF", "HOI"
                WriteHandoverInspectionEvent CurrentLine, test_date, latitude, longitude, latestCellId, handoverRow
                
            Case Is = "#STOP"
                Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("AO1").value = CurrentRowMos
                Sp = CurrentLine(1)
                If Keep(1) = "MTC" Then
                    Sheets("LTE_POLQA_MOS_MultiMetric_5").Range("AO1").value = 2
                End If
                Exit Do
        
        End Select
        
        
    Loop
    
    'Call Dup Function
    If Keep(1) = "MOC" Then
        outputArray = ArrayRemoveDups(AllCh)
        i = 2
        'Output values to Immediate Window (CTRL + G)
        Sheets("UMTS_3GIDLE_MultiMetric_5").Select
        For Each item In outputArray
            Range("T" & i).value = item
            i = i + 1
        Next item
    End If
    
    KPI.Close
    
    Sheets("LTE_POLQA_MOS_MultiMetric_5").Select
    Sleep 3000  ' 3 second delay
    lRow = Cells.Find(What:="*", _
                        After:=Range("A1"), _
                        LookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).row
    
    Range("W16").value = "AVERAGE"
    
    If IsEmpty(Range("O3").value) Then
        If IsEmpty(Range("O4").value) Then
            Else
            Range("X16").value = Round(WorksheetFunction.Average(Range("O2:O" & lRow)), 2)
        End If
    End If
    
    Sheets("Measurement_Info").Select
    LR = Cells(rows.count, 1).End(xlUp).row
    LR = LR + 1
        
        fileName = Split(fil, "\")
        WordCount = UBound(fileName())
        Worksheets("Measurement_Info").Range("D" & LR).value = test_date
        Worksheets("Measurement_Info").Range("A" & LR).value = fileName(WordCount)
        Worksheets("Measurement_Info").Range("I" & LR).value = Replace(Device, """", "")
        Worksheets("Measurement_Info").Range("B" & LR).value = ST
        Worksheets("Measurement_Info").Range("C" & LR).value = Sp
        Sheets("Measurement_Info").Columns(2).NumberFormat = "hh:mm:ss.000"
        Sheets("Measurement_Info").Columns(3).NumberFormat = "hh:mm:ss.000"
    
    
    Call TurnOnStuff
    excelSettingsOff = False
        
    If count = 0 Then
        AA = fileName(WordCount)
        Dayy = Split(AA, " CST")
        Dayy = Split(Dayy(1), ".")
        fileName = Split(AA, " ")
        fileName = Split(AA, fileName(1))
        fileName = Split(fileName(1), " ")
        Sheets("LTE_POLQA_MOS_MultiMetric_5").Columns(11).NumberFormat = "hh:mm:ss.000"
        Sheets("LTE_POLQA_MOS_MultiMetric_5").Columns(15).NumberFormat = "#.00"
        Sheets("LTE_POLQA_MOS_MultiMetric_5").Select
        TweeT
        Sheets("UMTS_3GIDLE_MultiMetric_5").Select
        Sheets("UMTS_3GIDLE_MultiMetric_5").Columns(11).NumberFormat = "hh:mm:ss.000"
        
        'CodeMos
        'CodeCov
        'Prep
        'ActiveWorkbook.SaveAs GetDesktop & "QoS Automation\" & DirPath & Keep(0) & " " & Filename(1) & " " & Dayy(0) & " CST.xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
        
    End If
    
    If Finish Then
        savePath = CSTOutputWorkbookPath(GetDesktop, DirPath, Keep(0), fileName(1), Dayy(0))
        ActiveWorkbook.SaveAs savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        DirPathGlobal = DirPath
        Sleep 3000  ' 3 second delay
        SaveTelcoWorkbook
        Sleep 3000  ' 3 second delay
        CodeMos
        Sleep 3000  ' 3 second delay
        CodeCov
        Sleep 3000  ' 3 second delay
        BuildHandoverKpiImpact
        Sleep 3000  ' 3 second delay
        Copy_CSTMOS_TO_Results
        Sleep 3000  ' 3 second delay
        Save
        Sleep 3000  ' 3 second delay
        If SingleCSTOutputEnabled Then ClearSingleCSTOutput
        If Last = False Then
            Clear
        Else
            If Last = True And isLastArrayElement = True Then
            For Each wb In Workbooks
            If wb.fullName = savePath Then
                wb.Close SaveChanges:=False
                Exit For
            End If
            Next wb
            End If
        End If
        
    End If

Else

End If
    Exit Sub

FatalError:
    errNumber = Err.Number
    errDescription = Err.Description
    On Error Resume Next
    If Not KPI Is Nothing Then KPI.Close
    If excelSettingsOff Then TurnOnStuff
    MsgBox "CSTCOVMOS failed while processing:" & vbCrLf & CStr(fil) & vbCrLf & _
           "Error " & errNumber & ": " & errDescription, vbExclamation, "CSTCOVMOS"
End Sub

Private Function CSTOutputWorkbookPath(ByVal desktopPath As String, ByVal dirPath As String, ByVal operatorName As String, ByVal fileToken As String, ByVal dayToken As String) As String
    Dim folderPath As String
    Dim outputName As String

    If SingleCSTOutputEnabled And Trim$(SingleCSTOutputFolder) <> "" And Trim$(SingleCSTOutputName) <> "" Then
        folderPath = EnsureTrailingSlash(SingleCSTOutputFolder)
        outputName = CleanWorkbookFileName(SingleCSTOutputName)
        If LCase$(Right$(outputName, 5)) <> ".xlsm" Then outputName = outputName & ".xlsm"
        CSTOutputWorkbookPath = folderPath & outputName
    Else
        CSTOutputWorkbookPath = desktopPath & "QoS Automation\" & dirPath & "\NCA\" & operatorName & "\CST\" & _
                                operatorName & " " & fileToken & " CST" & dayToken & ".xlsm"
    End If
End Function

Private Function EnsureTrailingSlash(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If Right$(folderPath, 1) = "\" Then
        EnsureTrailingSlash = folderPath
    Else
        EnsureTrailingSlash = folderPath & "\"
    End If
End Function

Private Function CleanWorkbookFileName(ByVal fileNameText As String) As String
    Dim badChars As Variant
    Dim item As Variant

    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    CleanWorkbookFileName = Trim$(fileNameText)
    For Each item In badChars
        CleanWorkbookFileName = Replace(CleanWorkbookFileName, CStr(item), "_")
    Next item
    If CleanWorkbookFileName = "" Then CleanWorkbookFileName = "CST Output"
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

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
        GetOrCreateWorksheet.Name = sheetName
    End If
End Function

Private Sub InitializeCellHandoverInspectionSheets(ByRef servingRow As Long, ByRef handoverRow As Long)
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = GetOrCreateWorksheet("SERVING_CELL_TRACKER")
    lastRow = ws.Cells(ws.rows.count, 11).End(xlUp).row
    If lastRow >= 2 Then ws.Range("K2:W" & lastRow).ClearContents
    'ws.Range("K1:W1").Value = Array("Time", "Date", "Latitude", "Longitude", "Full Cell ID", "System", "Band", "Channel", "Scrambling Code", "Ec/N0", "RSCP/RxLev", "RRC State", "Source Event")
    servingRow = 2

    Set ws = GetOrCreateWorksheet("HANDOVER_EVENTS")
    lastRow = ws.Cells(ws.rows.count, 11).End(xlUp).row
    If lastRow >= 2 Then ws.Range("K2:U" & lastRow).ClearContents
    'ws.Range("K1:U1").Value = Array("Time", "Date", "Latitude", "Longitude", "Event ID", "System", "Status/Type", "Detail", "Cells", "Latest Cell ID", "Raw Line")
    handoverRow = 2
End Sub

Private Sub UpdateLatestCellIdentity(ByRef fields() As String, ByRef latestChiTime As String, ByRef latestSystem As String, ByRef latestBand As String, ByRef latestRrcState As String, ByRef latestChannel As String, ByRef latestCellId As String)
    latestChiTime = SafeField(fields, 1)
    latestSystem = SafeField(fields, 3)
    latestBand = SafeField(fields, 4)
    latestRrcState = SafeField(fields, 5)
    latestChannel = SafeField(fields, 6)
    If SafeField(fields, 8) <> "" Then latestCellId = SafeField(fields, 8)
End Sub

Private Sub WriteServingCellInspectionRows(ByRef fields() As String, ByVal testDate As String, ByVal latitude As String, ByVal longitude As String, ByVal latestCellId As String, ByVal latestRrcState As String, ByRef servingRow As Long)
    Dim ws As Worksheet
    Dim systemType As String
    Dim numCells As Long
    Dim paramsPerCell As Long
    Dim i As Long
    Dim baseIndex As Long
    Dim cellType As String

    If UBound(fields) < 3 Then Exit Sub
    systemType = SafeField(fields, 3)
    Set ws = GetOrCreateWorksheet("SERVING_CELL_TRACKER")

    If systemType = "5" Or systemType = "6" Then
        If UBound(fields) < 11 Then Exit Sub
        If Not IsNumeric(SafeField(fields, 10)) Or Not IsNumeric(SafeField(fields, 11)) Then Exit Sub

        numCells = CLng(SafeField(fields, 10))
        paramsPerCell = CLng(SafeField(fields, 11))
        If numCells <= 0 Or paramsPerCell <= 0 Then Exit Sub

        For i = 0 To numCells - 1
            baseIndex = 12 + (i * paramsPerCell)
            If UBound(fields) >= baseIndex + 6 Then
                cellType = SafeField(fields, baseIndex)
                If cellType = "0" Then
                    ws.Range("K" & servingRow).value = SafeField(fields, 1)
                    ws.Range("L" & servingRow).value = testDate
                    ws.Range("M" & servingRow).value = latitude
                    ws.Range("N" & servingRow).value = longitude
                    ws.Range("O" & servingRow).value = latestCellId
                    ws.Range("P" & servingRow).value = systemType
                    ws.Range("Q" & servingRow).value = SafeField(fields, baseIndex + 1)
                    ws.Range("R" & servingRow).value = SafeField(fields, baseIndex + 2)
                    ws.Range("S" & servingRow).value = SafeField(fields, baseIndex + 3)
                    ws.Range("T" & servingRow).value = SafeField(fields, baseIndex + 4)
                    ws.Range("U" & servingRow).value = SafeField(fields, baseIndex + 6)
                    ws.Range("V" & servingRow).value = latestRrcState
                    ws.Range("W" & servingRow).value = "CELLMEAS"
                    servingRow = servingRow + 1
                End If
            End If
        Next i
    ElseIf systemType = "1" Then
        If UBound(fields) < 11 Then Exit Sub
        If SafeField(fields, 7) = "1" Then
            ws.Range("K" & servingRow).value = SafeField(fields, 1)
            ws.Range("L" & servingRow).value = testDate
            ws.Range("M" & servingRow).value = latitude
            ws.Range("N" & servingRow).value = longitude
            ws.Range("O" & servingRow).value = latestCellId
            ws.Range("P" & servingRow).value = systemType
            ws.Range("Q" & servingRow).value = SafeField(fields, 8)
            ws.Range("R" & servingRow).value = SafeField(fields, 9)
            ws.Range("S" & servingRow).value = ""
            ws.Range("T" & servingRow).value = ""
            ws.Range("U" & servingRow).value = SafeField(fields, 11)
            ws.Range("V" & servingRow).value = latestRrcState
            ws.Range("W" & servingRow).value = "CELLMEAS"
            servingRow = servingRow + 1
        End If
    End If
End Sub

Private Function NormalizeBandInfo(ByVal bandInfo As String) As String
    Dim mappedBand As String
    mappedBand = getBand(Trim$(bandInfo))
    If Trim$(mappedBand) <> "" Then
        NormalizeBandInfo = mappedBand
    Else
        NormalizeBandInfo = Trim$(bandInfo)
    End If
End Function

Private Sub WriteHandoverInspectionEvent(ByRef fields() As String, ByVal testDate As String, ByVal latitude As String, ByVal longitude As String, ByVal latestCellId As String, ByRef handoverRow As Long)
    Dim ws As Worksheet
    Dim i As Long
    Dim eventId As String
    Dim cellList As String

    eventId = SafeField(fields, 0)
    Set ws = GetOrCreateWorksheet("HANDOVER_EVENTS")

    For i = 8 To UBound(fields)
        If Trim$(SafeField(fields, i)) <> "" Then
            If cellList = "" Then
                cellList = SafeField(fields, i)
            Else
                cellList = cellList & "," & SafeField(fields, i)
            End If
        End If
    Next i

    ws.Range("K" & handoverRow).value = SafeField(fields, 1)
    ws.Range("L" & handoverRow).value = testDate
    ws.Range("M" & handoverRow).value = latitude
    ws.Range("N" & handoverRow).value = longitude
    ws.Range("O" & handoverRow).value = eventId
    ws.Range("P" & handoverRow).value = SafeField(fields, 3)
    ws.Range("Q" & handoverRow).value = SafeField(fields, 6)
    ws.Range("R" & handoverRow).value = SafeField(fields, 7)
    ws.Range("S" & handoverRow).value = cellList
    ws.Range("T" & handoverRow).value = latestCellId
    ws.Range("U" & handoverRow).value = Join(fields, ",")
    handoverRow = handoverRow + 1
End Sub

Private Function SafeField(ByRef fields() As String, ByVal index As Long) As String
    If index >= LBound(fields) And index <= UBound(fields) Then
        SafeField = Replace(fields(index), """", "")
    Else
        SafeField = ""
    End If
End Function

Public Sub RebuildHandoverKpiImpact()
    BuildHandoverKpiImpact
End Sub

Private Sub BuildHandoverKpiImpact()
    Dim wsHo As Worksheet, wsImpact As Worksheet
    Dim lastHo As Long, lastImpactRow As Long, outRow As Long, r As Long
    Dim hoMs As Double, beforeRow As Long, afterRow As Long, mosRow As Long, cstRow As Long
    Dim rscpBefore As Double, rscpAfter As Double, rscpDelta As Variant
    Dim dlBefore As Variant, ulBefore As Variant, dlAfter As Variant, ulAfter As Variant
    Dim mosValue As String, mosCategory As String, cstValue As String, cstCategory As String
    Dim terrainClass As String, elevationText As String, impactType As String, aiText As String
    Dim latValue As String, lonValue As String, cellBefore As String, cellAfter As String
    Dim scBefore As String, scAfter As String, bandBefore As String, bandAfter As String
    Dim channelBefore As String, channelAfter As String
    Dim nearestLocation As String, locationLat As String, locationLon As String, locationDistance As Double, locationSummary As String

    On Error GoTo Failed
    Set wsHo = ActiveWorkbook.Worksheets("HANDOVER_EVENTS")
    Set wsImpact = ActiveWorkbook.Worksheets("HANDOVER_KPI_IMPACT")
    ResetHandoverKpiImpactCaches
    lastImpactRow = wsImpact.Cells(wsImpact.rows.count, 11).End(xlUp).row
    If lastImpactRow >= 2 Then wsImpact.Range("K2:AY" & lastImpactRow).ClearContents

    'Sheet and headers are maintained manually in the template.
    'Set wsImpact = GetOrCreateSheet("HANDOVER_KPI_IMPACT")
    'wsImpact.Cells.ClearContents
    'wsImpact.Range("K1:AY1").Value = Array("Time", "Date", "Latitude", "Longitude", "Event ID", "Handover Cells", "Cell ID Before", "Cell ID After", "Serving SC Before", "Serving SC After", "Sector/Azimuth Before", "Sector/Azimuth After", "Band Before", "Band After", "Channel Before", "Channel After", "DL Freq Before", "DL Freq After", "UL Freq Before", "UL Freq After", "Frequency Change", "RSCP Before", "RSCP After", "RSCP Delta", "Ec/N0 Before", "Ec/N0 After", "MOS Near Handover", "MOS Category", "CST Near Handover", "CST Category", "Coverage Category", "Terrain Class", "Elevation At Handover", "Terrain Impact", "Impact Type", "AI Interpretation", "Nearest Location", "Location Latitude", "Location Longitude", "Distance To Location (m)", "Impact Location Summary")

    lastHo = wsHo.Cells(wsHo.rows.count, 11).End(xlUp).row
    outRow = 2

    For r = 2 To lastHo
        hoMs = TimeTextToMs(wsHo.Cells(r, 11).value)
        If hoMs >= 0 Then
            rscpDelta = Empty
            dlBefore = Empty
            ulBefore = Empty
            dlAfter = Empty
            ulAfter = Empty
            beforeRow = FindServingCellRow(hoMs, True, 10000)
            afterRow = FindServingCellRow(hoMs, False, 10000)
            mosRow = FindNearestTimeRow("LTE_POLQA_MOS_MultiMetric_5", 11, hoMs, 10000)
            cstRow = FindNearestTimeRow("GSM_VoiceCST_multimetric_6", 11, hoMs, 10000)

            cellBefore = SheetText("SERVING_CELL_TRACKER", beforeRow, 15)
            cellAfter = SheetText("SERVING_CELL_TRACKER", afterRow, 15)
            scBefore = SheetText("SERVING_CELL_TRACKER", beforeRow, 19)
            scAfter = SheetText("SERVING_CELL_TRACKER", afterRow, 19)
            bandBefore = NormalizeBandInfo(SheetText("SERVING_CELL_TRACKER", beforeRow, 17))
            bandAfter = NormalizeBandInfo(SheetText("SERVING_CELL_TRACKER", afterRow, 17))
            channelBefore = SheetText("SERVING_CELL_TRACKER", beforeRow, 18)
            channelAfter = SheetText("SERVING_CELL_TRACKER", afterRow, 18)

            GetChannelFrequency "", channelBefore, bandBefore, dlBefore, ulBefore
            GetChannelFrequency "", channelAfter, bandAfter, dlAfter, ulAfter

            rscpBefore = NumericOrSentinel(SheetText("SERVING_CELL_TRACKER", beforeRow, 21))
            rscpAfter = NumericOrSentinel(SheetText("SERVING_CELL_TRACKER", afterRow, 21))
            If rscpBefore > -99999 And rscpAfter > -99999 Then rscpDelta = rscpAfter - rscpBefore

            mosValue = SheetText("LTE_POLQA_MOS_MultiMetric_5", mosRow, 15)
            mosCategory = MosImpactCategory(mosValue)
            cstValue = SheetText("GSM_VoiceCST_multimetric_6", cstRow, 15)
            cstCategory = CstImpactCategory(cstValue)
            latValue = wsHo.Cells(r, 13).text
            lonValue = wsHo.Cells(r, 14).text
            nearestLocation = FindNearestTownLocation(latValue, lonValue, locationLat, locationLon, locationDistance)
            locationSummary = ImpactLocationSummary(nearestLocation, locationDistance)
            elevationText = ElevationTextForPoint(latValue, lonValue)
            terrainClass = TerrainClassFromContext(elevationText, rscpBefore, rscpAfter)
            impactType = HandoverImpactType(mosCategory, cstCategory, rscpDelta, scBefore, scAfter, CStr(dlBefore), CStr(dlAfter), terrainClass)
            aiText = BuildHandoverImpactSentence(wsHo.Cells(r, 11).text, impactType, scBefore, scAfter, CStr(dlBefore), CStr(dlAfter), rscpDelta, mosCategory, cstCategory, terrainClass, nearestLocation, locationDistance)

            With wsImpact
                .Cells(outRow, 11).value = wsHo.Cells(r, 11).value
                .Cells(outRow, 12).value = wsHo.Cells(r, 12).value
                .Cells(outRow, 13).value = latValue
                .Cells(outRow, 14).value = lonValue
                .Cells(outRow, 15).value = wsHo.Cells(r, 15).value
                .Cells(outRow, 16).value = wsHo.Cells(r, 19).value
                .Cells(outRow, 17).value = cellBefore
                .Cells(outRow, 18).value = cellAfter
                .Cells(outRow, 19).value = scBefore
                .Cells(outRow, 20).value = scAfter
                .Cells(outRow, 21).value = LookupTopologySectorAzimuth(cellBefore, scBefore, latValue, lonValue, SheetText("SERVING_CELL_TRACKER", beforeRow, 16), channelBefore)
                .Cells(outRow, 22).value = LookupTopologySectorAzimuth(cellAfter, scAfter, latValue, lonValue, SheetText("SERVING_CELL_TRACKER", afterRow, 16), channelAfter)
                .Cells(outRow, 23).value = bandBefore
                .Cells(outRow, 24).value = bandAfter
                .Cells(outRow, 25).value = channelBefore
                .Cells(outRow, 26).value = channelAfter
                .Cells(outRow, 27).value = dlBefore
                .Cells(outRow, 28).value = dlAfter
                .Cells(outRow, 29).value = ulBefore
                .Cells(outRow, 30).value = ulAfter
                .Cells(outRow, 31).value = YesNoText(CStr(dlBefore) <> CStr(dlAfter) Or CStr(ulBefore) <> CStr(ulAfter))
                If rscpBefore > -99999 Then .Cells(outRow, 32).value = rscpBefore
                If rscpAfter > -99999 Then .Cells(outRow, 33).value = rscpAfter
                If Not IsEmpty(rscpDelta) Then .Cells(outRow, 34).value = rscpDelta
                .Cells(outRow, 35).value = SheetText("SERVING_CELL_TRACKER", beforeRow, 20)
                .Cells(outRow, 36).value = SheetText("SERVING_CELL_TRACKER", afterRow, 20)
                .Cells(outRow, 37).value = mosValue
                .Cells(outRow, 38).value = mosCategory
                .Cells(outRow, 39).value = cstValue
                .Cells(outRow, 40).value = cstCategory
                .Cells(outRow, 41).value = CoverageImpactCategory(rscpBefore, rscpAfter)
                .Cells(outRow, 42).value = terrainClass
                .Cells(outRow, 43).value = elevationText
                .Cells(outRow, 44).value = TerrainImpactText(terrainClass)
                .Cells(outRow, 45).value = impactType
                .Cells(outRow, 46).value = aiText
                .Cells(outRow, 47).value = nearestLocation
                .Cells(outRow, 48).value = locationLat
                .Cells(outRow, 49).value = locationLon
                If locationDistance >= 0 Then .Cells(outRow, 50).value = Round(locationDistance, 1)
                .Cells(outRow, 51).value = locationSummary
            End With
            outRow = outRow + 1
        End If
    Next r
    Exit Sub
Failed:
    MsgBox "BuildHandoverKpiImpact failed: " & Err.Description, vbExclamation, "CSTCOVMOS"
End Sub

Private Sub ResetHandoverKpiImpactCaches()
    ServingTimeCacheWb = ""
    ServingTimeCount = 0
    Erase ServingTimes
    Erase ServingRows
    MosTimeCacheWb = ""
    MosTimeCount = 0
    Erase MosTimes
    Erase MosRows
    CstTimeCacheWb = ""
    CstTimeCount = 0
    Erase CstTimes
    Erase CstRows
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function FindServingCellRow(ByVal targetMs As Double, ByVal beforeTarget As Boolean, ByVal windowMs As Double) As Long
    Dim r As Long, thisMs As Double, diff As Double, bestDiff As Double
    On Error GoTo Failed
    EnsureServingTimeCache
    bestDiff = windowMs + 1
    For r = 1 To ServingTimeCount
        thisMs = ServingTimes(r)
        If thisMs >= 0 Then
            If (beforeTarget And thisMs <= targetMs) Or ((Not beforeTarget) And thisMs >= targetMs) Then
                diff = Abs(targetMs - thisMs)
                If diff <= windowMs And diff < bestDiff Then
                    bestDiff = diff
                    FindServingCellRow = ServingRows(r)
                End If
            End If
        End If
    Next r
    Exit Function
Failed:
    FindServingCellRow = 0
End Function

Private Function FindNearestTimeRow(ByVal sheetName As String, ByVal timeCol As Long, ByVal targetMs As Double, ByVal windowMs As Double) As Long
    Dim r As Long, thisMs As Double, diff As Double, bestDiff As Double
    On Error GoTo Failed
    EnsureMetricTimeCache sheetName, timeCol
    bestDiff = windowMs + 1
    If sheetName = "LTE_POLQA_MOS_MultiMetric_5" Then
        For r = 1 To MosTimeCount
            thisMs = MosTimes(r)
            If thisMs >= 0 Then
                diff = Abs(targetMs - thisMs)
                If diff <= windowMs And diff < bestDiff Then
                    bestDiff = diff
                    FindNearestTimeRow = MosRows(r)
                End If
            End If
        Next r
    ElseIf sheetName = "GSM_VoiceCST_multimetric_6" Then
        For r = 1 To CstTimeCount
            thisMs = CstTimes(r)
            If thisMs >= 0 Then
                diff = Abs(targetMs - thisMs)
                If diff <= windowMs And diff < bestDiff Then
                    bestDiff = diff
                    FindNearestTimeRow = CstRows(r)
                End If
            End If
        Next r
    Else
        FindNearestTimeRow = FindNearestTimeRowSlow(sheetName, timeCol, targetMs, windowMs)
    End If
    Exit Function
Failed:
    FindNearestTimeRow = 0
End Function

Private Function FindNearestTimeRowSlow(ByVal sheetName As String, ByVal timeCol As Long, ByVal targetMs As Double, ByVal windowMs As Double) As Long
    Dim ws As Worksheet, lastRow As Long, r As Long, thisMs As Double, diff As Double, bestDiff As Double
    On Error GoTo Failed
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.rows.count, timeCol).End(xlUp).row
    bestDiff = windowMs + 1
    For r = 2 To lastRow
        thisMs = TimeTextToMs(ws.Cells(r, timeCol).value)
        If thisMs >= 0 Then
            diff = Abs(targetMs - thisMs)
            If diff <= windowMs And diff < bestDiff Then
                bestDiff = diff
                FindNearestTimeRowSlow = r
            End If
        End If
    Next r
    Exit Function
Failed:
    FindNearestTimeRowSlow = 0
End Function

Private Sub EnsureServingTimeCache()
    If ServingTimeCacheWb = ActiveWorkbook.fullName And ServingTimeCount > 0 Then Exit Sub
    LoadTimeCache "SERVING_CELL_TRACKER", 11, ServingTimes, ServingRows, ServingTimeCount
    ServingTimeCacheWb = ActiveWorkbook.fullName
End Sub

Private Sub EnsureMetricTimeCache(ByVal sheetName As String, ByVal timeCol As Long)
    If sheetName = "LTE_POLQA_MOS_MultiMetric_5" Then
        If MosTimeCacheWb = ActiveWorkbook.fullName And MosTimeCount > 0 Then Exit Sub
        LoadTimeCache sheetName, timeCol, MosTimes, MosRows, MosTimeCount
        MosTimeCacheWb = ActiveWorkbook.fullName
    ElseIf sheetName = "GSM_VoiceCST_multimetric_6" Then
        If CstTimeCacheWb = ActiveWorkbook.fullName And CstTimeCount > 0 Then Exit Sub
        LoadTimeCache sheetName, timeCol, CstTimes, CstRows, CstTimeCount
        CstTimeCacheWb = ActiveWorkbook.fullName
    End If
End Sub

Private Sub LoadTimeCache(ByVal sheetName As String, ByVal timeCol As Long, ByRef times() As Double, ByRef rows() As Long, ByRef count As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim values As Variant
    Dim i As Long
    Dim t As Double

    On Error GoTo Failed
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.rows.count, timeCol).End(xlUp).row
    count = 0
    Erase times
    Erase rows
    If lastRow < 2 Then Exit Sub

    values = ws.Range(ws.Cells(2, timeCol), ws.Cells(lastRow, timeCol)).value
    ReDim times(1 To UBound(values, 1))
    ReDim rows(1 To UBound(values, 1))
    For i = 1 To UBound(values, 1)
        t = TimeTextToMs(values(i, 1))
        If t >= 0 Then
            count = count + 1
            times(count) = t
            rows(count) = i + 1
        End If
    Next i
    Exit Sub
Failed:
    count = 0
    Erase times
    Erase rows
End Sub

Private Function TimeTextToMs(ByVal timeValue As Variant) As Double
    Dim parts() As String, secParts() As String, h As Double, m As Double, s As Double
    Dim timeText As String
    On Error GoTo Failed
    If IsDate(timeValue) Then
        TimeTextToMs = (CDbl(CDate(timeValue)) - Fix(CDbl(CDate(timeValue)))) * 86400000#
        Exit Function
    End If
    If IsNumeric(timeValue) And CDbl(timeValue) > 0 And CDbl(timeValue) < 1 Then
        TimeTextToMs = CDbl(timeValue) * 86400000#
        Exit Function
    End If
    
    timeText = CStr(timeValue)
    timeText = Trim$(timeText)
    If timeText = "" Then GoTo Failed
    parts = Split(timeText, ":")
    If UBound(parts) < 1 Then GoTo Failed
    If UBound(parts) = 1 Then
        h = 0
        m = CDbl(parts(0))
        secParts = Split(parts(1), ".")
    Else
        h = CDbl(parts(0))
        m = CDbl(parts(1))
        secParts = Split(parts(2), ".")
    End If
    s = CDbl(secParts(0))
    If UBound(secParts) >= 1 Then s = s + (CDbl(secParts(1)) / (10 ^ Len(secParts(1))))
    TimeTextToMs = ((h * 3600#) + (m * 60#) + s) * 1000#
    Exit Function
Failed:
    TimeTextToMs = -1
End Function

Private Function SheetText(ByVal sheetName As String, ByVal rowNo As Long, ByVal colNo As Long) As String
    On Error GoTo Failed
    If rowNo <= 0 Or colNo <= 0 Then Exit Function
    SheetText = Trim$(CStr(ActiveWorkbook.Worksheets(sheetName).Cells(rowNo, colNo).value))
    Exit Function
Failed:
    SheetText = ""
End Function

Private Function NumericOrSentinel(ByVal valueText As String) As Double
    If IsNumeric(valueText) Then
        NumericOrSentinel = CDbl(valueText)
    Else
        NumericOrSentinel = -999999#
    End If
End Function

Private Function MosImpactCategory(ByVal mosText As String) As String
    If Not IsNumeric(mosText) Then
        MosImpactCategory = "Unavailable"
    ElseIf CDbl(mosText) < 3# Then
        MosImpactCategory = "Poor"
    ElseIf CDbl(mosText) < 3.6 Then
        MosImpactCategory = "Fair"
    Else
        MosImpactCategory = "Stable"
    End If
End Function

Private Function CstImpactCategory(ByVal cstText As String) As String
    Dim secondsValue As Double
    If Not IsNumeric(cstText) Then
        CstImpactCategory = "Unavailable"
        Exit Function
    End If
    secondsValue = CDbl(cstText)
    If secondsValue > 0 And secondsValue < 1 Then secondsValue = secondsValue * 86400#
    If secondsValue >= 6# Then
        CstImpactCategory = "Extended"
    ElseIf secondsValue >= 4# Then
        CstImpactCategory = "Moderate"
    Else
        CstImpactCategory = "Stable"
    End If
End Function

Private Function CoverageImpactCategory(ByVal rscpBefore As Double, ByVal rscpAfter As Double) As String
    If rscpBefore <= -99999 Or rscpAfter <= -99999 Then
        CoverageImpactCategory = "Unavailable"
    ElseIf rscpAfter <= -95 Or rscpAfter - rscpBefore <= -5 Then
        CoverageImpactCategory = "Poor/Degraded"
    ElseIf rscpAfter <= -85 Then
        CoverageImpactCategory = "Fair"
    Else
        CoverageImpactCategory = "Stable"
    End If
End Function

Private Function HandoverImpactType(ByVal mosCategory As String, ByVal cstCategory As String, ByVal rscpDelta As Variant, ByVal scBefore As String, ByVal scAfter As String, ByVal freqBefore As String, ByVal freqAfter As String, ByVal terrainClass As String) As String
    Dim changedCell As Boolean, changedFreq As Boolean, degradedCoverage As Boolean
    changedCell = (Trim$(scBefore) <> "" And Trim$(scAfter) <> "" And Trim$(scBefore) <> Trim$(scAfter))
    changedFreq = (Trim$(freqBefore) <> "" And Trim$(freqAfter) <> "" And Trim$(freqBefore) <> Trim$(freqAfter))
    If Not IsEmpty(rscpDelta) Then degradedCoverage = (CDbl(rscpDelta) <= -5#)
    If (mosCategory = "Poor" Or cstCategory = "Extended") And (changedCell Or changedFreq Or degradedCoverage) Then
        HandoverImpactType = "Likely Handover Impact"
    ElseIf (mosCategory = "Fair" Or cstCategory = "Moderate" Or degradedCoverage Or InStr(1, terrainClass, "transition", vbTextCompare) > 0 Or InStr(1, terrainClass, "obstruction", vbTextCompare) > 0) And (changedCell Or changedFreq) Then
        HandoverImpactType = "Possible Handover Impact"
    Else
        HandoverImpactType = "Unlikely Handover Impact"
    End If
End Function

Private Function BuildHandoverImpactSentence(ByVal hoTime As String, ByVal impactType As String, ByVal scBefore As String, ByVal scAfter As String, ByVal freqBefore As String, ByVal freqAfter As String, ByVal rscpDelta As Variant, ByVal mosCategory As String, ByVal cstCategory As String, ByVal terrainClass As String, Optional ByVal nearestLocation As String = "", Optional ByVal locationDistance As Double = -1) As String
    BuildHandoverImpactSentence = impactType & ": handover at " & hoTime
    If Trim$(nearestLocation) <> "" Then BuildHandoverImpactSentence = BuildHandoverImpactSentence & " near " & nearestLocation & LocationDistanceText(locationDistance)
    If scBefore <> "" Or scAfter <> "" Then BuildHandoverImpactSentence = BuildHandoverImpactSentence & ", serving SC changed from " & ValueOrUnknown(scBefore) & " to " & ValueOrUnknown(scAfter)
    If freqBefore <> "" Or freqAfter <> "" Then BuildHandoverImpactSentence = BuildHandoverImpactSentence & ", DL frequency changed from " & ValueOrUnknown(freqBefore) & " MHz to " & ValueOrUnknown(freqAfter) & " MHz"
    If Not IsEmpty(rscpDelta) Then BuildHandoverImpactSentence = BuildHandoverImpactSentence & ", RSCP delta was " & Format(CDbl(rscpDelta), "0.0") & " dB"
    BuildHandoverImpactSentence = BuildHandoverImpactSentence & ", MOS was " & mosCategory & ", CST was " & cstCategory & ", terrain context was " & terrainClass & "."
End Function

Private Function LocationDistanceText(ByVal distanceMeters As Double) As String
    If distanceMeters < 0 Then
        LocationDistanceText = ""
    ElseIf distanceMeters >= 1000# Then
        LocationDistanceText = " (~" & Format(distanceMeters / 1000#, "0.00") & " km from town reference)"
    Else
        LocationDistanceText = " (~" & Format(distanceMeters, "0") & " m from town reference)"
    End If
End Function

Private Function ImpactLocationSummary(ByVal nearestLocation As String, ByVal distanceMeters As Double) As String
    If Trim$(nearestLocation) = "" Then
        ImpactLocationSummary = "Town location reference unavailable for this handover point."
    Else
        ImpactLocationSummary = "Impact point is near " & nearestLocation & LocationDistanceText(distanceMeters) & "."
    End If
End Function

Private Function FindNearestTownLocation(ByVal latText As String, ByVal lonText As String, ByRef locationLat As String, ByRef locationLon As String, ByRef distanceMeters As Double) As String
    Static cacheBuilt As Boolean
    Static locNames() As String
    Static locLats() As Double
    Static locLons() As Double
    Static locCount As Long
    Dim wsTown As Worksheet, latCol As Long, lonCol As Long, nameCol As Long, headerRow As Long
    Dim lastRow As Long, r As Long, latValue As Double, lonValue As Double
    Dim testLat As Double, testLon As Double, d As Double, bestDistance As Double, bestIndex As Long

    On Error GoTo Failed
    locationLat = ""
    locationLon = ""
    distanceMeters = -1
    If Not IsNumeric(latText) Or Not IsNumeric(lonText) Then Exit Function
    testLat = CDbl(latText)
    testLon = CDbl(lonText)

    If Not cacheBuilt Then
        Set wsTown = FindTownLocationSheet()
        If Not wsTown Is Nothing Then
            If FindTownLocationColumns(wsTown, headerRow, latCol, lonCol, nameCol) Then
                lastRow = wsTown.Cells(wsTown.rows.count, latCol).End(xlUp).row
                ReDim locNames(1 To lastRow)
                ReDim locLats(1 To lastRow)
                ReDim locLons(1 To lastRow)
                For r = headerRow + 1 To lastRow
                    If IsNumeric(wsTown.Cells(r, latCol).value) And IsNumeric(wsTown.Cells(r, lonCol).value) Then
                        latValue = CDbl(wsTown.Cells(r, latCol).value)
                        lonValue = CDbl(wsTown.Cells(r, lonCol).value)
                        If IsPlausibleCoordinate(latValue, lonValue) Then
                            locCount = locCount + 1
                            If nameCol > 0 Then locNames(locCount) = Trim$(CStr(wsTown.Cells(r, nameCol).value))
                            If locNames(locCount) = "" Then locNames(locCount) = wsTown.Name & " reference row " & CStr(r)
                            locLats(locCount) = latValue
                            locLons(locCount) = lonValue
                        End If
                    End If
                Next r
            End If
        End If
        cacheBuilt = True
    End If

    If locCount = 0 Then Exit Function
    bestDistance = 1E+30
    For r = 1 To locCount
        d = HaversineDistanceMeters(testLat, testLon, locLats(r), locLons(r))
        If d < bestDistance Then
            bestDistance = d
            bestIndex = r
        End If
    Next r

    If bestIndex > 0 Then
        FindNearestTownLocation = locNames(bestIndex)
        locationLat = CStr(locLats(bestIndex))
        locationLon = CStr(locLons(bestIndex))
        distanceMeters = bestDistance
    End If
    Exit Function
Failed:
    FindNearestTownLocation = ""
    locationLat = ""
    locationLon = ""
    distanceMeters = -1
End Function

Private Function FindTownLocationSheet() As Worksheet
    Dim ws As Worksheet, headerRow As Long, latCol As Long, lonCol As Long, nameCol As Long
    For Each ws In ActiveWorkbook.Worksheets
        If Not IsStandardWorkbookSheet(ws.Name) Then
            If FindTownLocationColumns(ws, headerRow, latCol, lonCol, nameCol) Then
                Set FindTownLocationSheet = ws
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function FindTownLocationColumns(ByVal ws As Worksheet, ByRef headerRow As Long, ByRef latCol As Long, ByRef lonCol As Long, ByRef nameCol As Long) As Boolean
    Dim rr As Long, cc As Long, headerText As String
    headerRow = 0
    latCol = 0
    lonCol = 0
    nameCol = 0

    For rr = 1 To 10
        For cc = 1 To 100
            headerText = UCase$(Trim$(CStr(ws.Cells(rr, cc).value)))
            If headerText <> "" Then
                If latCol = 0 And InStr(1, headerText, "LAT", vbTextCompare) > 0 Then latCol = cc
                If lonCol = 0 And (InStr(1, headerText, "LON", vbTextCompare) > 0 Or InStr(1, headerText, "LONG", vbTextCompare) > 0) Then lonCol = cc
                If nameCol = 0 And IsLocationNameHeader(headerText) Then nameCol = cc
            End If
        Next cc
        If latCol > 0 And lonCol > 0 Then
            headerRow = rr
            If nameCol = 0 Then nameCol = FirstTownNameColumn(ws, headerRow, latCol, lonCol)
            FindTownLocationColumns = True
            Exit Function
        End If
        latCol = 0
        lonCol = 0
        nameCol = 0
    Next rr
End Function

Private Function FirstTownNameColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal latCol As Long, ByVal lonCol As Long) As Long
    Dim cc As Long, headerText As String
    For cc = 1 To 100
        If cc <> latCol And cc <> lonCol Then
            headerText = Trim$(CStr(ws.Cells(headerRow, cc).value))
            If headerText <> "" Then
                FirstTownNameColumn = cc
                Exit Function
            End If
        End If
    Next cc
End Function

Private Function IsLocationNameHeader(ByVal headerText As String) As Boolean
    IsLocationNameHeader = (InStr(1, headerText, "LOCATION", vbTextCompare) > 0 Or _
                            InStr(1, headerText, "PLACE", vbTextCompare) > 0 Or _
                            InStr(1, headerText, "NAME", vbTextCompare) > 0 Or _
                            InStr(1, headerText, "DESCRIPTION", vbTextCompare) > 0 Or _
                            InStr(1, headerText, "COMMUNITY", vbTextCompare) > 0 Or _
                            InStr(1, headerText, "AREA", vbTextCompare) > 0 Or _
                            InStr(1, headerText, "LANDMARK", vbTextCompare) > 0)
End Function

Private Function IsPlausibleCoordinate(ByVal latValue As Double, ByVal lonValue As Double) As Boolean
    IsPlausibleCoordinate = (latValue >= -90# And latValue <= 90# And lonValue >= -180# And lonValue <= 180#)
End Function

Private Function HaversineDistanceMeters(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Double
    Const EarthRadiusMeters As Double = 6371000#
    Const PiValue As Double = 3.14159265358979
    Dim dLat As Double, dLon As Double, a As Double, c As Double
    dLat = (lat2 - lat1) * PiValue / 180#
    dLon = (lon2 - lon1) * PiValue / 180#
    lat1 = lat1 * PiValue / 180#
    lat2 = lat2 * PiValue / 180#
    a = Sin(dLat / 2#) * Sin(dLat / 2#) + Cos(lat1) * Cos(lat2) * Sin(dLon / 2#) * Sin(dLon / 2#)
    c = 2# * Atn(Sqr(a) / Sqr(1# - a))
    HaversineDistanceMeters = EarthRadiusMeters * c
End Function

Private Function IsStandardWorkbookSheet(ByVal sheetName As String) As Boolean
    Dim n As String
    n = UCase$(Trim$(sheetName))
    Select Case n
        Case "RAW VOICE", "CST RAW", "MEASUREMENT_INFO", "GSM_VOICEREPORT_MULTIMETRIC_5", _
             "GSM_VOICECST_MULTIMETRIC_6", "LTE_POLQA_MOS_MULTIMETRIC_5", "UMTS_3GIDLE_MULTIMETRIC_5", _
             "SERVING_CELL_TRACKER", "HANDOVER_EVENTS", "HANDOVER_KPI_IMPACT", "GSM_2G_MULTIMETRIC_1", _
             "CALL_SEQUENCE_MOC", "CALL_SEQUENCE_MTC", "MOSPIE", "AVGCALMOS", "COUNT", "AVGCALCOV", _
             "GLO", "QOS MAP VIEW", "NATIVE MAP DATA", "COVERAGE MAP", "MOS MAP", "GOOGLE COVERAGE MAP", _
             "GOOGLE MOS MAP", "GOOGLE MAP CONFIG", "COVERAGE ELEVATION ANALYSIS", "MOS ELEVATION ANALYSIS", _
             "QOS EMBEDDED ASSETS", "MTN", "TELECEL", "AT", "NYAME MIND", "ANALYZE REPORT DESCRIPTION", _
             "REPORT LOG", "POLICY ENGINE", "GOVERNANCE SCHEMA", "PROMPT ROUTER", "ORCHESTRATION", _
             "VALIDATION PIPELINE", "FREQ SHEET", "REF SHEET", "FILTERMOS", "ADVFILTERMOS", "FILTERCOV", _
             "ADVFILTERCOV", "GOOGLE CST MAP", "GOOGLE CALL SEQUENCE MAP"
            IsStandardWorkbookSheet = True
    End Select
End Function

Private Function TerrainClassFromContext(ByVal elevationText As String, ByVal rscpBefore As Double, ByVal rscpAfter As Double) As String
    If Trim$(elevationText) = "" Then
        If rscpBefore > -99999 And rscpAfter > -99999 And rscpAfter - rscpBefore <= -5# Then
            TerrainClassFromContext = "RF transition/possible obstruction; elevation unavailable"
        Else
            TerrainClassFromContext = "Terrain unavailable"
        End If
    ElseIf rscpBefore > -99999 And rscpAfter > -99999 And rscpAfter - rscpBefore <= -5# Then
        TerrainClassFromContext = "Terrain/RF transition zone"
    Else
        TerrainClassFromContext = "No clear terrain-linked degradation"
    End If
End Function

Private Function TerrainImpactText(ByVal terrainClass As String) As String
    If InStr(1, terrainClass, "transition", vbTextCompare) > 0 Or InStr(1, terrainClass, "obstruction", vbTextCompare) > 0 Then
        TerrainImpactText = "Terrain may have contributed to dominance change or post-handover degradation."
    ElseIf terrainClass = "Terrain unavailable" Then
        TerrainImpactText = "Elevation/terrain context unavailable for this handover point."
    Else
        TerrainImpactText = "No strong terrain contribution visible from available data."
    End If
End Function

Private Function ElevationTextForPoint(ByVal latText As String, ByVal lonText As String) As String
    Dim cachePath As String, fsoLocal As Object, cacheText As String, keyText As String
    Dim startPos As Long, elevPos As Long, colonPos As Long, commaPos As Long
    On Error GoTo Failed
    If Not IsNumeric(latText) Or Not IsNumeric(lonText) Then Exit Function
    keyText = """" & Format(CDbl(latText), "0.000000") & "," & Format(CDbl(lonText), "0.000000") & """"
    cachePath = ActiveWorkbook.Path & "\elevation_output\elevation_cache.json"
    Set fsoLocal = CreateObject("Scripting.FileSystemObject")
    If Not fsoLocal.FileExists(cachePath) Then Exit Function
    cacheText = fsoLocal.OpenTextFile(cachePath).ReadAll
    startPos = InStr(1, cacheText, keyText, vbTextCompare)
    If startPos = 0 Then Exit Function
    elevPos = InStr(startPos, cacheText, """elevation""", vbTextCompare)
    colonPos = InStr(elevPos, cacheText, ":")
    commaPos = InStr(colonPos, cacheText, ",")
    If colonPos > 0 And commaPos > colonPos Then ElevationTextForPoint = Trim$(Mid$(cacheText, colonPos + 1, commaPos - colonPos - 1))
    Exit Function
Failed:
    ElevationTextForPoint = ""
End Function

Private Function LookupTopologySectorAzimuth(ByVal cellId As String, ByVal scramblingCode As String, ByVal latText As String, ByVal lonText As String, ByVal systemType As String, ByVal channelNo As String) As String
    Dim networkName As String, techText As String
    Dim i As Long, bestIndex As Long, exactIndex As Long
    Dim pointLat As Double, pointLon As Double, distMeters As Double, bearingValue As Double, azDiff As Double
    Dim bestScore As Double, scoreValue As Double, bestDistance As Double, bestAzDiff As Double

    On Error GoTo Failed
    networkName = DetectOperatorName()
    techText = SystemTypeToTopologyTech(systemType)
    If techText = "" Then techText = "3G"

    EnsureTopologyCache networkName, techText
    If TopologyCount = 0 Then
        LookupTopologySectorAzimuth = ""
        Exit Function
    End If

    If Trim$(cellId) <> "" Then
        For i = 1 To TopologyCount
            If TopologyTechs(i) = techText And Trim$(TopologyCellIds(i)) = Trim$(cellId) Then
                exactIndex = i
                Exit For
            End If
        Next i
    End If

    If exactIndex > 0 Then
        LookupTopologySectorAzimuth = TopologyMatchLabel(exactIndex, -1, -1, "Exact topology CELL ID match")
        Exit Function
    End If

    If Not IsNumeric(latText) Or Not IsNumeric(lonText) Then
        LookupTopologySectorAzimuth = ""
        Exit Function
    End If

    pointLat = CDbl(latText)
    pointLon = CDbl(lonText)
    LookupTopologySectorAzimuth = CachedTopologyGpsMatch(techText, pointLat, pointLon)
    If LookupTopologySectorAzimuth <> "" Then
        If Trim$(channelNo) <> "" Then LookupTopologySectorAzimuth = LookupTopologySectorAzimuth & "; UARFCN " & channelNo
        If Trim$(scramblingCode) <> "" Then LookupTopologySectorAzimuth = LookupTopologySectorAzimuth & "; PSC " & scramblingCode
        Exit Function
    End If

    bestScore = 1E+30
    bestDistance = -1
    bestAzDiff = -1

    For i = 1 To TopologyCount
        If TopologyTechs(i) = techText And IsPlausibleCoordinate(TopologyLats(i), TopologyLons(i)) Then
            distMeters = HaversineDistanceMeters(pointLat, pointLon, TopologyLats(i), TopologyLons(i))
            If distMeters <= 5000# Then
                bearingValue = BearingDegrees(TopologyLats(i), TopologyLons(i), pointLat, pointLon)
                azDiff = AzimuthDifferenceDegrees(bearingValue, TopologyAzimuths(i))
                scoreValue = distMeters + (azDiff * 20#)
                If scoreValue < bestScore Then
                    bestScore = scoreValue
                    bestIndex = i
                    bestDistance = distMeters
                    bestAzDiff = azDiff
                End If
            End If
        End If
    Next i

    If bestIndex > 0 Then
        LookupTopologySectorAzimuth = TopologyMatchLabel(bestIndex, bestDistance, bestAzDiff, "Inferred by GPS + nearest sector azimuth")
        CacheTopologyGpsMatch techText, pointLat, pointLon, LookupTopologySectorAzimuth
        If Trim$(channelNo) <> "" Then LookupTopologySectorAzimuth = LookupTopologySectorAzimuth & "; UARFCN " & channelNo
        If Trim$(scramblingCode) <> "" Then LookupTopologySectorAzimuth = LookupTopologySectorAzimuth & "; PSC " & scramblingCode
    End If
    Exit Function
Failed:
    LookupTopologySectorAzimuth = ""
End Function

Private Sub EnsureTopologyCache(ByVal networkName As String, Optional ByVal targetTech As String = "")
    Dim cacheKey As String, topologyPath As String
    Dim topoWb As Workbook, ws As Worksheet
    Dim openedHere As Boolean

    On Error GoTo Failed
    cacheKey = UCase$(Trim$(networkName)) & "|" & UCase$(Trim$(targetTech))
    If cacheKey = "" Then cacheKey = "UNKNOWN"
    If TopologyCacheBuiltFor = cacheKey Then Exit Sub

    ClearTopologyCache
    TopologyCacheBuiltFor = cacheKey
    topologyPath = TopologyWorkbookPath(cacheKey)
    If Trim$(topologyPath) = "" Or Dir(topologyPath) = "" Then Exit Sub

    Set topoWb = GetOpenWorkbookByFullName(topologyPath)
    If topoWb Is Nothing Then
        Set topoWb = Workbooks.Open(fileName:=topologyPath, ReadOnly:=True, UpdateLinks:=False)
        openedHere = True
    End If

    For Each ws In topoWb.Worksheets
        LoadTopologyRowsFromSheet ws, targetTech
    Next ws

    If openedHere Then topoWb.Close SaveChanges:=False
    Exit Sub
Failed:
    On Error Resume Next
    If openedHere And Not topoWb Is Nothing Then topoWb.Close SaveChanges:=False
End Sub

Private Sub ClearTopologyCache()
    TopologyCount = 0
    Erase TopologySiteNames
    Erase TopologyCellNames
    Erase TopologyCellIds
    Erase TopologyTechs
    Erase TopologyAzimuths
    Erase TopologyLats
    Erase TopologyLons
    Set TopologyGpsMatchCache = Nothing
End Sub

Private Sub LoadTopologyRowsFromSheet(ByVal ws As Worksheet, Optional ByVal targetTech As String = "")
    Dim headerRow As Long, siteCol As Long, cellNameCol As Long, cellIdCol As Long, azCol As Long, latCol As Long, lonCol As Long
    Dim lastRow As Long, r As Long, techText As String

    On Error GoTo Failed
    headerRow = 0
    siteCol = HeaderColumnByText(ws, "SITE NAME", headerRow)
    If siteCol = 0 Then siteCol = HeaderColumnByText(ws, "SITE", headerRow)
    cellNameCol = HeaderColumnByText(ws, "CELL NAME", headerRow)
    cellIdCol = HeaderColumnByText(ws, "CELL ID", headerRow)
    azCol = HeaderColumnByText(ws, "AZIMUTH", headerRow)
    latCol = HeaderColumnByText(ws, "LATITUDE", headerRow)
    lonCol = HeaderColumnByText(ws, "LONGITUDE", headerRow)

    If headerRow = 0 Or siteCol = 0 Or cellNameCol = 0 Or cellIdCol = 0 Or azCol = 0 Or latCol = 0 Or lonCol = 0 Then Exit Sub
    techText = TopologyTechFromSheet(ws.Name)
    If techText = "" Then techText = TopologyTechFromCellName(CStr(ws.Cells(headerRow + 1, cellNameCol).value))
    If techText = "" Then Exit Sub
    If Trim$(targetTech) <> "" And StrComp(techText, targetTech, vbTextCompare) <> 0 Then Exit Sub

    lastRow = ws.Cells(ws.rows.count, siteCol).End(xlUp).row
    For r = headerRow + 1 To lastRow
        If Trim$(CStr(ws.Cells(r, siteCol).value)) <> "" And IsNumeric(ws.Cells(r, latCol).value) And IsNumeric(ws.Cells(r, lonCol).value) Then
            AddTopologyCacheRow techText, CStr(ws.Cells(r, siteCol).value), CStr(ws.Cells(r, cellNameCol).value), CStr(ws.Cells(r, cellIdCol).value), CDbl(Val(ws.Cells(r, azCol).value)), CDbl(ws.Cells(r, latCol).value), CDbl(ws.Cells(r, lonCol).value)
        End If
    Next r
    Exit Sub
Failed:
End Sub

Private Function CachedTopologyGpsMatch(ByVal techText As String, ByVal latValue As Double, ByVal lonValue As Double) As String
    Dim keyText As String
    If TopologyGpsMatchCache Is Nothing Then Set TopologyGpsMatchCache = CreateObject("Scripting.Dictionary")
    keyText = TopologyGpsCacheKey(techText, latValue, lonValue)
    If TopologyGpsMatchCache.Exists(keyText) Then CachedTopologyGpsMatch = CStr(TopologyGpsMatchCache(keyText))
End Function

Private Sub CacheTopologyGpsMatch(ByVal techText As String, ByVal latValue As Double, ByVal lonValue As Double, ByVal labelText As String)
    Dim keyText As String
    If TopologyGpsMatchCache Is Nothing Then Set TopologyGpsMatchCache = CreateObject("Scripting.Dictionary")
    keyText = TopologyGpsCacheKey(techText, latValue, lonValue)
    If Not TopologyGpsMatchCache.Exists(keyText) Then TopologyGpsMatchCache.Add keyText, labelText
End Sub

Private Function TopologyGpsCacheKey(ByVal techText As String, ByVal latValue As Double, ByVal lonValue As Double) As String
    TopologyGpsCacheKey = UCase$(Trim$(techText)) & "|" & Format(latValue, "0.0000") & "|" & Format(lonValue, "0.0000")
End Function

Private Sub AddTopologyCacheRow(ByVal techText As String, ByVal siteName As String, ByVal cellName As String, ByVal cellId As String, ByVal azimuthValue As Double, ByVal latValue As Double, ByVal lonValue As Double)
    If TopologyCount = 0 Then
        ReDim TopologySiteNames(1 To 500)
        ReDim TopologyCellNames(1 To 500)
        ReDim TopologyCellIds(1 To 500)
        ReDim TopologyTechs(1 To 500)
        ReDim TopologyAzimuths(1 To 500)
        ReDim TopologyLats(1 To 500)
        ReDim TopologyLons(1 To 500)
    ElseIf TopologyCount >= UBound(TopologySiteNames) Then
        ReDim Preserve TopologySiteNames(1 To UBound(TopologySiteNames) + 500)
        ReDim Preserve TopologyCellNames(1 To UBound(TopologyCellNames) + 500)
        ReDim Preserve TopologyCellIds(1 To UBound(TopologyCellIds) + 500)
        ReDim Preserve TopologyTechs(1 To UBound(TopologyTechs) + 500)
        ReDim Preserve TopologyAzimuths(1 To UBound(TopologyAzimuths) + 500)
        ReDim Preserve TopologyLats(1 To UBound(TopologyLats) + 500)
        ReDim Preserve TopologyLons(1 To UBound(TopologyLons) + 500)
    End If

    TopologyCount = TopologyCount + 1
    TopologyTechs(TopologyCount) = techText
    TopologySiteNames(TopologyCount) = Trim$(siteName)
    TopologyCellNames(TopologyCount) = Trim$(cellName)
    TopologyCellIds(TopologyCount) = Trim$(cellId)
    TopologyAzimuths(TopologyCount) = azimuthValue
    TopologyLats(TopologyCount) = latValue
    TopologyLons(TopologyCount) = lonValue
End Sub

Private Function TopologyMatchLabel(ByVal indexNo As Long, ByVal distanceMeters As Double, ByVal azimuthDelta As Double, ByVal methodText As String) As String
    If indexNo <= 0 Or indexNo > TopologyCount Then Exit Function
    TopologyMatchLabel = TopologySiteNames(indexNo) & " " & TopologyCellNames(indexNo) & " CellID " & TopologyCellIds(indexNo) & " Az " & Format(TopologyAzimuths(indexNo), "0")
    If distanceMeters >= 0 Or azimuthDelta >= 0 Then
        TopologyMatchLabel = TopologyMatchLabel & " ("
        If distanceMeters >= 0 Then TopologyMatchLabel = TopologyMatchLabel & Format(distanceMeters / 1000#, "0.00") & " km"
        If distanceMeters >= 0 And azimuthDelta >= 0 Then TopologyMatchLabel = TopologyMatchLabel & ", "
        If azimuthDelta >= 0 Then TopologyMatchLabel = TopologyMatchLabel & "az diff " & Format(azimuthDelta, "0") & " deg"
        TopologyMatchLabel = TopologyMatchLabel & ")"
    End If
    TopologyMatchLabel = TopologyMatchLabel & " [" & methodText & "]"
End Function

Private Function DetectOperatorName() As String
    Dim nameText As String
    nameText = UCase$(ActiveWorkbook.Name)
    If InStr(1, nameText, "TELECEL", vbTextCompare) > 0 Then
        DetectOperatorName = "TELECEL"
    ElseIf InStr(1, nameText, "AIRTELTIGO", vbTextCompare) > 0 Or Left$(nameText, 2) = "AT" Or InStr(1, nameText, " AT ", vbTextCompare) > 0 Then
        DetectOperatorName = "AT"
    ElseIf InStr(1, nameText, "MTN", vbTextCompare) > 0 Then
        DetectOperatorName = "MTN"
    Else
        DetectOperatorName = FirstToken(GetSheetFirstText("Measurement_Info", 9))
    End If
End Function

Private Function FirstToken(ByVal valueText As String) As String
    Dim parts() As String
    valueText = Trim$(valueText)
    If valueText = "" Then Exit Function
    parts = Split(valueText, " ")
    FirstToken = UCase$(Trim$(parts(0)))
End Function

Private Function GetSheetFirstText(ByVal sheetName As String, ByVal colNo As Long) As String
    Dim ws As Worksheet, lastRow As Long, r As Long
    On Error GoTo Failed
    Set ws = ActiveWorkbook.Worksheets(sheetName)
    lastRow = ws.Cells(ws.rows.count, colNo).End(xlUp).row
    For r = 2 To lastRow
        If Trim$(CStr(ws.Cells(r, colNo).value)) <> "" Then
            GetSheetFirstText = Trim$(CStr(ws.Cells(r, colNo).value))
            Exit Function
        End If
    Next r
    Exit Function
Failed:
    GetSheetFirstText = ""
End Function

Private Function TopologyWorkbookPath(ByVal networkName As String) As String
    Dim basePath As String, fileName As String
    basePath = Environ$("USERPROFILE") & Application.PathSeparator & "Desktop" & Application.PathSeparator & "QoS Template" & Application.PathSeparator
    Select Case UCase$(Trim$(networkName))
        Case "MTN"
            fileName = "NCA_RPM System_GIS_Topology_MTN_January_2026 1.xlsx"
        Case "TELECEL", "AT", "AIRTELTIGO"
            fileName = "Telecel_RPM System_GIS_Topology_MAY_2025 1.xlsx"
        Case Else
            fileName = ""
    End Select
    If fileName <> "" Then
        TopologyWorkbookPath = basePath & fileName
        If Dir(TopologyWorkbookPath) = "" Then TopologyWorkbookPath = ActiveWorkbook.Path & Application.PathSeparator & fileName
    End If
End Function

Private Function GetOpenWorkbookByFullName(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Workbooks
        If UCase$(wb.fullName) = UCase$(fullPath) Then
            Set GetOpenWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
End Function

Private Function SystemTypeToTopologyTech(ByVal systemType As String) As String
    Select Case Trim$(systemType)
        Case "1"
            SystemTypeToTopologyTech = "2G"
        Case "5", "6"
            SystemTypeToTopologyTech = "3G"
        Case "7"
            SystemTypeToTopologyTech = "4G"
    End Select
End Function

Private Function TopologyTechFromSheet(ByVal sheetName As String) As String
    Dim s As String
    s = UCase$(sheetName)
    If InStr(1, s, "2G", vbTextCompare) > 0 Or InStr(1, s, "GSM", vbTextCompare) > 0 Then
        TopologyTechFromSheet = "2G"
    ElseIf InStr(1, s, "3G", vbTextCompare) > 0 Or InStr(1, s, "UMTS", vbTextCompare) > 0 Then
        TopologyTechFromSheet = "3G"
    ElseIf InStr(1, s, "4G", vbTextCompare) > 0 Or InStr(1, s, "LTE", vbTextCompare) > 0 Then
        TopologyTechFromSheet = "4G"
    End If
End Function

Private Function TopologyTechFromCellName(ByVal cellName As String) As String
    Dim s As String
    s = UCase$(cellName)
    If InStr(1, s, "2G", vbTextCompare) > 0 Or Left$(s, 1) = "G" Then
        TopologyTechFromCellName = "2G"
    ElseIf InStr(1, s, "3G", vbTextCompare) > 0 Or Left$(s, 1) = "U" Then
        TopologyTechFromCellName = "3G"
    ElseIf InStr(1, s, "4G", vbTextCompare) > 0 Or Left$(s, 1) = "L" Then
        TopologyTechFromCellName = "4G"
    End If
End Function

Private Function HeaderColumnByText(ByVal ws As Worksheet, ByVal token As String, ByRef headerRow As Long) As Long
    Dim rr As Long, cc As Long, headerText As String, tokenText As String
    tokenText = UCase$(Trim$(token))
    For rr = 1 To 20
        For cc = 1 To 100
            headerText = UCase$(Trim$(CStr(ws.Cells(rr, cc).value)))
            If headerText = tokenText Then
                headerRow = rr
                HeaderColumnByText = cc
                Exit Function
            End If
        Next cc
    Next rr
    For rr = 1 To 20
        For cc = 1 To 100
            headerText = UCase$(Trim$(CStr(ws.Cells(rr, cc).value)))
            If headerText <> "" And InStr(1, headerText, tokenText, vbTextCompare) > 0 Then
                headerRow = rr
                HeaderColumnByText = cc
                Exit Function
            End If
        Next cc
    Next rr
End Function

Private Function BearingDegrees(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Double
    Const PiValue As Double = 3.14159265358979
    Dim phi1 As Double, phi2 As Double, lambda1 As Double, lambda2 As Double
    Dim y As Double, x As Double, brng As Double
    phi1 = lat1 * PiValue / 180#
    phi2 = lat2 * PiValue / 180#
    lambda1 = lon1 * PiValue / 180#
    lambda2 = lon2 * PiValue / 180#
    y = Sin(lambda2 - lambda1) * Cos(phi2)
    x = Cos(phi1) * Sin(phi2) - Sin(phi1) * Cos(phi2) * Cos(lambda2 - lambda1)
    brng = Atn2(y, x) * 180# / PiValue
    BearingDegrees = NormalizeDegrees(brng)
End Function

Private Function Atn2(ByVal y As Double, ByVal x As Double) As Double
    Const PiValue As Double = 3.14159265358979
    If x > 0 Then
        Atn2 = Atn(y / x)
    ElseIf x < 0 And y >= 0 Then
        Atn2 = Atn(y / x) + PiValue
    ElseIf x < 0 And y < 0 Then
        Atn2 = Atn(y / x) - PiValue
    ElseIf x = 0 And y > 0 Then
        Atn2 = PiValue / 2#
    ElseIf x = 0 And y < 0 Then
        Atn2 = -PiValue / 2#
    Else
        Atn2 = 0
    End If
End Function

Private Function NormalizeDegrees(ByVal degreesValue As Double) As Double
    Do While degreesValue < 0
        degreesValue = degreesValue + 360#
    Loop
    Do While degreesValue >= 360#
        degreesValue = degreesValue - 360#
    Loop
    NormalizeDegrees = degreesValue
End Function

Private Function AzimuthDifferenceDegrees(ByVal bearingValue As Double, ByVal azimuthValue As Double) As Double
    Dim diffValue As Double
    diffValue = Abs(NormalizeDegrees(bearingValue) - NormalizeDegrees(azimuthValue))
    If diffValue > 180# Then diffValue = 360# - diffValue
    AzimuthDifferenceDegrees = diffValue
End Function

Private Function LookupSectorAzimuth(ByVal cellId As String, ByVal scramblingCode As String) As String
    Static azimuthCache As Object
    Static cacheBuilt As Boolean
    Dim ws As Worksheet, headerRow As Long, cellCol As Long, scCol As Long, azCol As Long, sectorCol As Long, siteCol As Long
    Dim lastRow As Long, r As Long, keyText As String, labelText As String, sourceCol As Long
    On Error GoTo Failed

    If Not cacheBuilt Then
        Set azimuthCache = CreateObject("Scripting.Dictionary")
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Name <> "HANDOVER_EVENTS" And ws.Name <> "SERVING_CELL_TRACKER" And ws.Name <> "HANDOVER_KPI_IMPACT" Then
                headerRow = 0
                cellCol = HeaderColumn(ws, "CELL", headerRow)
                scCol = HeaderColumn(ws, "SCRAMBLING", headerRow)
                If scCol = 0 Then scCol = HeaderColumn(ws, "PSC", headerRow)
                azCol = HeaderColumn(ws, "AZIMUTH", headerRow)
                If azCol = 0 Then azCol = HeaderColumn(ws, "AZI", headerRow)
                sectorCol = HeaderColumn(ws, "SECTOR", headerRow)
                siteCol = HeaderColumn(ws, "SITE", headerRow)
                If azCol > 0 And (cellCol > 0 Or scCol > 0) Then
                    If cellCol > 0 Then sourceCol = cellCol Else sourceCol = scCol
                    lastRow = ws.Cells(ws.rows.count, sourceCol).End(xlUp).row
                    For r = headerRow + 1 To lastRow
                        labelText = ""
                        If siteCol > 0 Then labelText = Trim$(CStr(ws.Cells(r, siteCol).value))
                        If sectorCol > 0 And Trim$(CStr(ws.Cells(r, sectorCol).value)) <> "" Then labelText = Trim$(labelText & " Sector " & CStr(ws.Cells(r, sectorCol).value))
                        If azCol > 0 And Trim$(CStr(ws.Cells(r, azCol).value)) <> "" Then labelText = Trim$(labelText & " Az " & CStr(ws.Cells(r, azCol).value))
                        If labelText <> "" Then
                            If cellCol > 0 Then
                                keyText = Trim$(CStr(ws.Cells(r, cellCol).value))
                                If keyText <> "" And Not azimuthCache.Exists("CELL|" & keyText) Then azimuthCache.Add "CELL|" & keyText, labelText
                            End If
                            If scCol > 0 Then
                                keyText = Trim$(CStr(ws.Cells(r, scCol).value))
                                If keyText <> "" And Not azimuthCache.Exists("SC|" & keyText) Then azimuthCache.Add "SC|" & keyText, labelText
                            End If
                        End If
                    Next r
                End If
            End If
        Next ws
        cacheBuilt = True
    End If

    If Not azimuthCache Is Nothing Then
        If Trim$(cellId) <> "" Then
            keyText = "CELL|" & Trim$(cellId)
            If azimuthCache.Exists(keyText) Then
                LookupSectorAzimuth = azimuthCache(keyText)
                Exit Function
            End If
        End If
        If Trim$(scramblingCode) <> "" Then
            keyText = "SC|" & Trim$(scramblingCode)
            If azimuthCache.Exists(keyText) Then
                LookupSectorAzimuth = azimuthCache(keyText)
                Exit Function
            End If
        End If
                    End If
    Exit Function
Failed:
    LookupSectorAzimuth = ""
End Function

Private Function HeaderColumn(ByVal ws As Worksheet, ByVal token As String, ByRef headerRow As Long) As Long
    Dim rr As Long, cc As Long, headerText As String
    For rr = 1 To 10
        For cc = 1 To 100
            headerText = UCase$(Trim$(CStr(ws.Cells(rr, cc).value)))
            If headerText <> "" And InStr(1, headerText, UCase$(token), vbTextCompare) > 0 Then
                headerRow = rr
                HeaderColumn = cc
                Exit Function
            End If
        Next cc
    Next rr
End Function

Private Function YesNoText(ByVal flag As Boolean) As String
    If flag Then YesNoText = "Yes" Else YesNoText = "No"
End Function

Private Function ValueOrUnknown(ByVal valueText As String) As String
    If Trim$(valueText) = "" Then ValueOrUnknown = "unknown" Else ValueOrUnknown = valueText
End Function
Function IsNumericAndLessOrEqualZero(value As Variant) As Boolean
    On Error GoTo ErrorHandler
    If IsNumeric(value) And value <> "" Then
        IsNumericAndLessOrEqualZero = (CDbl(value) <= 0)
    Else
        IsNumericAndLessOrEqualZero = False
    End If
    Exit Function
ErrorHandler:
    IsNumericAndLessOrEqualZero = False
End Function
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
        
    Case Is = "10850"
    
        Result = "GSM 850"
        
    Case Is = "10900"
    
        Result = "GSM 900"
        
    Case Is = "11800"
    
        Result = "GSM 1800"
        
    Case Is = "11900"
    
        Result = "GSM 1900"
        
    Case Is = "19999"
    
        Result = "GSM"
    
    End Select
    
    getBand = Result
    
End Function
'Function ArrayRemoveDups(MyArray As Variant) As Variant
'    Dim nFirst As Long, nLast As Long, i As Long
'    Dim item As String
'
'    Dim arrTemp() As String
'    Dim Coll As New Collection
'
'    'Get First and Last Array Positions
'    nFirst = LBound(MyArray)
'    nLast = UBound(MyArray)
'    ReDim arrTemp(nFirst To nLast)
'
'    'Convert Array to String
'    For i = nFirst To nLast
'        arrTemp(i) = CStr(MyArray(i))
'    Next i
'
'    'Populate Temporary Collection, removing empty entries
'    On Error Resume Next
'    For i = nFirst To nLast
'        If arrTemp(i) <> "" Then ' Check if the entry is not empty
'            Coll.Add arrTemp(i), arrTemp(i)
'        End If
'    Next i
'    Err.Clear
'    On Error GoTo 0
'
'    'Resize Array
'    nLast = Coll.Count + nFirst - 1
'
'    ReDim arrTemp(nFirst To nLast)
'
'    'Populate Array
'    For i = nFirst To nLast
'        arrTemp(i) = Coll(i - nFirst + 1)
'    Next i
'
'    'Output Array
'    ArrayRemoveDups = arrTemp
'
'End Function
Function ArrayRemoveDups(MyArray As Variant) As Variant
    Dim nFirst As Long, nLast As Long, i As Long
    Dim arrTemp() As String
    Dim Coll As New Collection

    ' Validate input is an array
    If Not IsArray(MyArray) Then
        ArrayRemoveDups = Array()
        Exit Function
    End If

    ' Get array bounds
    nFirst = LBound(MyArray)
    nLast = UBound(MyArray)
    
    ' Handle empty input array immediately
    If nFirst > nLast Then
        ArrayRemoveDups = Array()
        Exit Function
    End If

    ReDim arrTemp(nFirst To nLast)

    ' Process elements
    On Error Resume Next
    For i = nFirst To nLast
        If Not IsEmpty(MyArray(i)) And Not IsNull(MyArray(i)) Then
            Dim item As String
            item = CStr(MyArray(i))
            If item <> "" Then
                ' Add with error handling for duplicates
                Coll.Add item, item
                If Err.Number = 457 Then Err.Clear  ' Clear duplicate key error
            End If
        End If
    Next i
    On Error GoTo 0

    ' Handle empty collection (no valid elements)
    If Coll.count = 0 Then
        ArrayRemoveDups = Array()
    Else
        ' Resize and populate output array
        ReDim arrTemp(nFirst To nFirst + Coll.count - 1)
        For i = 1 To Coll.count
            arrTemp(nFirst + i - 1) = Coll(i)
        Next i
    End If

    If Coll.count > 0 Then ArrayRemoveDups = arrTemp
End Function
Private Function FindworkbookPath() As String
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
Function GetFileNameOnly(fgname As String) As String
    ' Returns file name without extension
    Dim fullName As String
    Dim dotPosition As Long
    Dim baseName As String
    Dim partPosition As Long
    
    ' First get the full file name
    fullName = GetFileNameFromPath(fgname)
    
    ' Find the last dot position
    dotPosition = InStrRev(fullName, ".")
    
    If dotPosition > 0 Then
        ' Return name without extension
        baseName = Left(fullName, dotPosition - 1)
        partPosition = InStrRev(baseName, ".")
        If partPosition > 0 And IsNumeric(Mid$(baseName, partPosition + 1)) Then
            baseName = Left$(baseName, partPosition - 1)
        End If
        GetFileNameOnly = baseName
    Else
        ' If no extension found, return full name
        GetFileNameOnly = fullName
    End If
End Function
Function GetFileNameFromPath(fgname As String) As String
    ' Comprehensive function to extract file name from path
    On Error GoTo ErrorHandler
    
    If fgname = "" Then
        GetFileNameFromPath = ""
        Exit Function
    End If
    
    ' Find the last backslash position
    Dim lastBackslash As Integer
    lastBackslash = InStrRev(fgname, "\")
    
    If lastBackslash > 0 Then
        ' Extract file name after the last backslash
        GetFileNameFromPath = Mid(fgname, lastBackslash + 1)
    Else
        ' If no backslash found, return the original string
        GetFileNameFromPath = fgname
    End If
    
    Exit Function
    
ErrorHandler:
    GetFileNameFromPath = ""
End Function

