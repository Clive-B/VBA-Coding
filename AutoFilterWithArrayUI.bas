Attribute VB_Name = "AutoFilterWithArrayUI"
Option Base 1
' Add this at the top of your module (before any procedures)
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub AutoFilterWithArray()
Dim ID_range, k As Variant
Dim LR As Long

If Not IsSingleCoverageProcessing() Then GetARFCNValues
AddARFCNValues
Sheets("UMTS_3GIDLE_MultiMetric_5").Select
LR = Cells(Rows.count, 20).End(xlUp).row
If LR < 2 Then Exit Sub
ID_range = Application.Transpose(ActiveSheet.Range("T2:T" & LR))

If LR = 2 Then

Sheets("UMTS_3GIDLE_MultiMetric_5").Range("K1:T1").AutoFilter Field:=6, Criteria1:=Range("T2").value

Else

For k = LBound(ID_range) To UBound(ID_range)
ID_range(k) = CStr(ID_range(k))
Next k

Sleep 3000  ' 3 second delay
Sheets("UMTS_3GIDLE_MultiMetric_5").Range("K1:T1").AutoFilter Field:=6, Operator:=xlFilterValues, _
Criteria1:=ID_range
Sleep 3000  ' 3 second delay

End If

End Sub

Private Function IsSingleCoverageProcessing() As Boolean
    On Error Resume Next
    IsSingleCoverageProcessing = CBool(Single3GCovOutputEnabled)
    If Not IsSingleCoverageProcessing Then IsSingleCoverageProcessing = CBool(Single4GCovOutputEnabled)
    On Error GoTo 0
End Function


