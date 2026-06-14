Attribute VB_Name = "Spots"
Sub CreateSpotSheets()
    Dim wsLocations As Worksheet
    Dim wsSource As Worksheet
    Dim wsSpot As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastRo As Long
    Dim i As Long
    Dim j As Long
    Dim spotName As String
    Dim beginRow As Long
    Dim endRow As Long
    Dim dataRows As Long
    
    ' Set worksheets
    Set wsLocations = ThisWorkbook.Sheets("Locations")
    Set wsSource = ThisWorkbook.Sheets("GSM_DataReport_multimetric_5")
    Set ws = ThisWorkbook.Sheets("Logs")
    
    ' Find last row in Locations sheet
    lastRow = wsLocations.Cells(wsLocations.rows.Count, "A").End(xlUp).row
    
    ' Find last row in Locations sheet
    lastRo = ws.Cells(ws.rows.Count, "A").End(xlUp).row
    
    ' Turn off screen updating for faster execution
    Application.ScreenUpdating = False
    
    ' Loop through each row in Locations sheet (starting from row 2)
    For i = 2 To lastRow
        ' Get values from Locations sheet
        spotName = wsLocations.Cells(i, "A").value
        beginRow = wsLocations.Cells(i, "C").value
        endRow = wsLocations.Cells(i, "D").value
        
        
        ' Validate that we have valid row numbers
        If beginRow <= 0 Or endRow <= 0 Or endRow < beginRow Then
'            MsgBox "Invalid row numbers for " & spotName & _
                   ": Beginning=" & beginRow & ", Ending=" & endRow, vbExclamation
            GoTo ContinueLoop
        End If
        
        ' Calculate number of data rows
        dataRows = endRow - beginRow + 1
        
        ' Check if sheet already exists
        On Error Resume Next
        Set wsSpot = ThisWorkbook.Sheets(spotName)
        On Error GoTo 0
        
        If wsSpot Is Nothing Then
            ' Create new sheet
            Set wsSpot = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsSpot.Name = spotName
        Else
            ' Clear existing sheet if it exists
            wsSpot.Cells.Clear
        End If
        
        ' Copy headers from source sheet (K1:R1 to A1:H1)
        wsSource.Range("K1:R1").Copy Destination:=wsSpot.Range("A1:H1")
        
        ' Copy data from source sheet (K[beginning]:R[ending] to A2:H[dataRows+1])
        wsSource.Range("K" & beginRow & ":R" & endRow).Copy _
            Destination:=wsSpot.Range("A2:H" & (dataRows + 1))
        
        ' Auto-fit columns for better visibility
        wsSpot.Columns("A:H").AutoFit
        
        ' Add sheet name as title
        wsSpot.Range("J1").value = spotName & " - " & wsLocations.Cells(i, "B").value
        wsSpot.Range("J1").Font.Bold = True
        wsSpot.Range("J1").Font.Size = 14
        
        ' Display progress
        Debug.Print "Created sheet: " & spotName & _
                   " (Rows " & beginRow & " to " & endRow & ")"
    
        ' Set the target worksheet
        Set wsSpot = ThisWorkbook.Sheets("Measurement_Info")
        
'        For j = 2 To lastRo
        
        LogfileName = ws.Range("C" & i).value
        If wsLocations.Range("B" & i) = ws.Range("C" & i) Then
        
        ' Write the formula in cell G1
        wsSpot.Range("G" & i).Formula = "=PERCENTILE.INC(" + spotName + "!H:H,0.9)"
'        Exit For

        Else
            j = i
            Do Until wsLocations.Range("B" & i) = ws.Range("C" & j)
            j = j + 1
            Loop
            
            ' Write the formula in cell G1
        wsSpot.Range("G" & j).Formula = "=PERCENTILE.INC(" + spotName + "!H:H,0.9)"
            
        End If
        
'        Next j
        
ContinueLoop:
        ' Reset worksheet object for next iteration
        Set wsSpot = Nothing
    Next i
    
    ' Turn screen updating back on
    Application.ScreenUpdating = True
    
End Sub

