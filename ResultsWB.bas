Attribute VB_Name = "ResultsWB"
Option Explicit

Public Sub Copy_CSTMOS_TO_Results(Optional ByVal dirPathValue As String = "")
    Dim res() As String
    Dim Pen() As String
    Dim Rop() As String
    Dim adName As String
    Dim distCap As String
    Dim newWorkbook As Workbook
    Dim oWSHShell As Object
    Dim GetDesktop As String
    Dim dayString As String
    Dim resultsPath As String

    dirPathValue = Trim$(dirPathValue)
    If dirPathValue = "" Then
        MsgBox "Copy_CSTMOS_TO_Results needs a campaign path.", vbExclamation, "Results Workbook"
        Exit Sub
    End If

    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop") & "\"
    Set oWSHShell = Nothing

    With ThisWorkbook.Sheets("Measurement_Info")
        adName = Trim$(CStr(.Range("A2").Value))
        Pen = Split(Trim$(CStr(.Range("I2").Value)), " ")
        res = Split(adName, " ")
    End With

    If UBound(Pen) < 0 Or UBound(res) < 3 Then
        MsgBox "Measurement_Info does not contain enough data to update Results.xlsm.", _
               vbExclamation, "Results Workbook"
        Exit Sub
    End If

    Rop = Split(Split(adName, res(3) & " ")(1), ".")
    dayString = Rop(0)
    distCap = WorksheetFunction.Proper(res(2))

    resultsPath = GetDesktop & "QoS Automation\" & dirPathValue & "\Results\Results.xlsm"
    If Dir(resultsPath) = "" Then
        MsgBox "Results workbook was not found:" & vbCrLf & resultsPath, _
               vbExclamation, "Results Workbook"
        Exit Sub
    End If

    Set newWorkbook = Workbooks.Open(resultsPath)

    Select Case UCase$(Pen(0))
        Case "AT"
            ProcessOperatorData newWorkbook, "Count", "B13", "C:F", Array("H16:I16"), 2, dayString, distCap
            ProcessOperatorData newWorkbook, "LTE_POLQA_MOS_MultiMetric_5", "Z16", "C:F", Array("H29:I29"), 2, dayString, distCap
        Case "GLO"
            ProcessOperatorData newWorkbook, "Count", "B13", "L:O", Array("H16:I16"), 2, dayString, distCap
            ProcessOperatorData newWorkbook, "LTE_POLQA_MOS_MultiMetric_5", "Z16", "L:O", Array("H16:I16"), 2, dayString, distCap
        Case "MTN"
            ProcessOperatorData newWorkbook, "Count", "B13", "L:O", Array("Q16:R16"), 11, dayString, distCap
            ProcessOperatorData newWorkbook, "LTE_POLQA_MOS_MultiMetric_5", "Z16", "L:O", Array("Q29:R29"), 11, dayString, distCap
        Case "TELECEL"
            ProcessOperatorData newWorkbook, "Count", "B13", "U:X", Array("Z16:AA16"), 20, dayString, distCap
            ProcessOperatorData newWorkbook, "LTE_POLQA_MOS_MultiMetric_5", "Z16", "U:X", Array("Z29:AA29"), 20, dayString, distCap
        Case Else
            MsgBox "Unsupported operator for Results update: " & Pen(0), _
                   vbExclamation, "Results Workbook"
    End Select

    newWorkbook.Close SaveChanges:=True
End Sub

Private Sub ProcessOperatorData( _
    ByVal targetWorkbook As Workbook, _
    ByVal sourceSheet As String, _
    ByVal sourceCell As String, _
    ByVal dayColumns As String, _
    ByVal rowRanges As Variant, _
    ByVal checkColumn As Long, _
    ByVal dayString As String, _
    ByVal distCap As String)

    Dim dayCols() As String
    Dim colIndex As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim i As Long
    Dim sourceValue As Variant
    Dim dayNumber As Long
    Dim rowRange As Variant

    dayCols = Split(dayColumns, ":")
    sourceValue = ThisWorkbook.Sheets(sourceSheet).Range(sourceCell).Value
    dayNumber = CLng(Replace(UCase$(dayString), "DAY ", ""))

    colIndex = Columns(dayCols(0)).Column + (dayNumber - 1)

    If colIndex > Columns(dayCols(1)).Column Then
        MsgBox "Invalid day configuration for columns: " & dayColumns, vbCritical, "Results Workbook"
        Exit Sub
    End If

    For Each rowRange In rowRanges
        With targetWorkbook.Sheets("Results")
            startRow = .Range(Split(rowRange, ":")(0)).Value
            endRow = .Range(Split(rowRange, ":")(1)).Value
        End With

        If startRow > endRow Then
            For i = startRow To endRow Step -1
                UpdateTargetCell targetWorkbook, i, checkColumn, colIndex, sourceValue, distCap
            Next i
        Else
            For i = startRow To endRow
                UpdateTargetCell targetWorkbook, i, checkColumn, colIndex, sourceValue, distCap
            Next i
        End If
    Next rowRange
End Sub

Private Sub UpdateTargetCell( _
    ByVal targetWorkbook As Workbook, _
    ByVal rowNumber As Long, _
    ByVal checkColumn As Long, _
    ByVal targetColumn As Long, _
    ByVal sourceValue As Variant, _
    ByVal distCap As String)

    With targetWorkbook.Sheets("Results")
        If .Cells(rowNumber, checkColumn).Value <> "" Then
            If .Cells(rowNumber, 2).Value = distCap Then
                .Cells(rowNumber, targetColumn).Value = sourceValue
            End If
        End If
    End With
End Sub
