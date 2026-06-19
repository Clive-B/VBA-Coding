Attribute VB_Name = "EARFCNComp"
Sub AddEARFCNValues()
    Dim ws As Worksheet
    Dim values As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim networkName As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("LTE_4GIDLE_MultiMetric_5")
    networkName = ResolveCoverageNetwork()
    
    If networkName = "MTN" Then
    ' Your values array
    values = Array("6375", "6400", "3200", "3000")
    
    ElseIf networkName = "TELECEL" Then
    values = Array("6225")

    Else
    MsgBox "Cannot determine supported 4G network for EARFCN filtering.", vbExclamation, "EARFCN Values"
    Exit Sub
    
    End If
    
    ' Find the first empty row in column T
    lastRow = ws.Cells(ws.Rows.count, "T").End(xlUp).row + 1
    
    ' If the column is completely empty, start at row 1
    If lastRow = 2 And ws.Range("T1").value = "" Then
        lastRow = 1
    End If
    
    ' Add the values
    For i = LBound(values) To UBound(values)
        ws.Cells(lastRow + i, 20).value = values(i) ' 20 = column T
    Next i
    
    ' Remove duplicates from the entire column
    RemoveDuplicatesFromColumnT_Simple ws
End Sub

Private Function ResolveCoverageNetwork() As String
    Dim networkName As String

    networkName = NetworkFromMeasurementInfo()

    If networkName = "" Then
        On Error Resume Next
        networkName = CStr(Application.Run("'QoS.xlsm'!GetNetwork"))
        On Error GoTo 0
    End If

    ResolveCoverageNetwork = UCase$(Trim$(networkName))
End Function

Private Function NetworkFromMeasurementInfo() As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNo As Long
    Dim deviceText As String
    Dim parts() As String

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Measurement_Info")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
    For rowNo = 2 To lastRow
        deviceText = Trim$(CStr(ws.Cells(rowNo, "I").value))
        If deviceText <> "" Then
            parts = Split(deviceText, " ")
            NetworkFromMeasurementInfo = UCase$(Replace(parts(0), """", ""))
            Exit Function
        End If
    Next rowNo
End Function

Sub RemoveDuplicatesFromColumnT_Simple(ws As Worksheet)
    Dim lastRow As Long
    
    ' Find the last row with data in column T
    lastRow = ws.Cells(ws.Rows.count, "T").End(xlUp).row
    
    ' Only remove duplicates if there's more than 1 row of data
    If lastRow > 1 Then
        ws.Range("T1:T" & lastRow).RemoveDuplicates Columns:=1, Header:=xlNo
    End If
End Sub

