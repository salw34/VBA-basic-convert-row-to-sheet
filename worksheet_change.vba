Private Sub Worksheet_Change(ByVal Target As Range)
' Call main procedure when cell value changed.
On Error GoTo Erorrhandler
Dim KeyCells As Range
Dim targetValue
' The variable KeyCells contains the cells that will.
' Set keycells by name range or cell.
Set KeyCells = Range("order_number")

If Not Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
  ' Set KeyCells value for matching.
  order_number = Range(Target.Address).value
      
  ' Display a message when one of the designated cells has been
  If Len(targetValue) = 9 And IsNumeric(targetValue) Then        
    ' Working on button Caption & worksheet
    Call Convert_Row_to_Sheet
  End If
End If
    
' Enable Event in Excel
Application.EnableEvents = True   
Exit Sub
   
Erorrhandler:
' Re-enable events, even if an error occurs.
Application.EnableEvents = True
End Sub
