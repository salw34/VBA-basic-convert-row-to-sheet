Sub Convert_Row_to_Sheet()
On Error GoTo ErrorHandler
' Disable Application event
Application.EnableEvents = False
Application.ScreenUpdating = False

Dim mainPage As Worksheet
Dim query_sheet As Worksheet
Dim firstOccurrence As Variant
Dim lastOccurrence As Variant
Dim match_order_range As Range
Dim order_number As Variant
Dim time_table As Object

Set mainPage = ThisWorkbook.Sheets("mainPage")
Set query_sheet = ThisWorkbook.Sheets("powerQuery")
Set time_table = mainPage.ListObjects("time_table")

' Order number will set to cell name order_number to use in selection and referance.
order_number = mainPage.Range("order_number").value

' Order number must be number and 9 chars, validation#1.
If Not IsNumeric(order_number) or Len(order_number) <> 9  Then
    Exit Sub
End If

' Clear table date.
time_table.DataBodyRange.ClearContents

' Find the first, last occurrence in powerQueery table by mathing with order_number (Order Number) and set to new range by using excel funtion.
' powerQuery must be sorted by Order Number.
firstOccurrence = Application.Match(order_number, query_sheet.Columns(1), 0)
lastOccurrence = Application.Match(order_number, query_sheet.Columns(1), 1)
Set match_order_range = query_sheet.Range("A" & firstOccurrence & ":A" & lastOccurrence)

' Initializing variable for combine_value_array array.
Dim refering_text1 As String, separator as String
Dim value_join1 As String
Dim combine_range_Value As String
Dim foundIndex As Long, foundCount As Long, counter_value_combine1 As Long, count_array As Long
Dim combine_value_array() As String, value_join_array() As String, counter_value_join_array() As String
Dim queryRow As Range

ReDim combine_value_array(0 To 0)
ReDim value_join_array(0 To 0)
ReDim counter_value_join_array(0 To 0)
separator = "_"

' #Part1# Create array from each row in match_order_range.
For Each queryRow In match_order_range

    ' Assign value to variables.
    value_combine1 = queryRow.cells(1, 1).value
    value_StDate = queryRow.cells(1, 2).value
    value_EnDate = queryRow.cells(1, 3).value    
    value_offset_column = queryRow.cells(1, 4).value
    value_join1 = queryRow.cells(1, 5).value
    refering_text1 = queryRow.cells(1, 6).value
    counter_value_combine1 = queryRow.cells(1, 7).value

    ' Combine text in each column in the same row and using "_" as separator
    combine_range_Value =value_combine1 & separator & value_StDate & separator & value_EnDate separator & value_offset_column
                            
    ' Combine the text and set the initiate array value;
    ' There are 3 arrays for storing necessary data.
    If combine_value_array(0) = "" Then
        combine_value_array(0) = combine_range_Value
        value_join_array(0) = value_join1
        counter_value_join_array(0) = counter_value_combine1
    Else
        ' Check if current combine_range_Value already added to array then return array index  to foundCount
        ' Set foundIndex to -1 for prevent to return array value
        foundIndex = -1
        For foundCount = 0 To count_array
            If combine_value_array(foundCount) = combine_range_Value Then
                foundIndex = foundCount
                Exit For
            End If
        Next foundCount        

        ' Assign values to 3 main arrays.
        ' Add value_join1 to exist array in value_join_array.
        ' Resize array and assign current value to new array index if combine_range_Value not exist in value_join_array
        If foundIndex >= 0 Then
            value_join_array(foundIndex) = value_join_array(foundIndex) & "," & value_join1
            counter_value_join_array(foundIndex) = counter_value_join_array(foundIndex) & "," & counter_value_combine1
        Else
                        count_array = count_array + 1
            ReDim Preserve combine_value_array(0 To count_array)
            ReDim Preserve value_join_array(0 To count_array)
            ReDim Preserve counter_value_join_array(0 To count_array)
            combine_value_array(count_array) = combine_range_Value
            value_join_array(count_array) = value_join1
            counter_value_join_array(count_array) = counter_value_combine1
        End If
    End If
Next queryRow

combine_value_array
value_join_array
counter_value_join_array

' #Part2# filling combined text of array to cells.
Dim combine_array_index As Long
Dim add_new_row As Integer
dim split_combine_arrray
' //Optional variable//
Dim join_array_value As String
Dim counter_value As String

'set initial row in table 
add_new_row = 1

For combine_array_index = LBound(combine_value_array) To UBound(combine_value_array)

    ' Skip empty array refer to combine_value_array.
    If combine_value_array(combine_array_index) <> "" Then

        ' Create the array values to fill into cells.
        split_combine_arrray = Split(combine_value_array(combine_array_index), separator)

        ' set offset column number to fill value, refer to value_offset_column.
        Select Case split_combine_arrray(3)
            Case "Text2"
                set_offset_column_to_fill = 6
            Case "Text3"
                set_offset_column_to_fill = 7,
            Case Else
                set_offset_column_to_fill = 5
        End Select

        ' You can remove join_array_value & counter_value by use array index insetad.
        join_array_value = value_join_array(combine_array_index)
        counter_value = counter_value_join_array(combine_array_index)
        
        ' Fill array value in to each column in row
        ' You can use range instead table object by change "time_table.DataBodyRange.Cells()" to mainPage.range(row, column)
        time_table.DataBodyRange.Cells(add_new_row, 1).value = add_new_row
        time_table.DataBodyRange.Cells(add_new_row, 2).value = split_combine_arrray(0)
        time_table.DataBodyRange.Cells(add_new_row, 3).value = split_combine_arrray(1)
        time_table.DataBodyRange.Cells(add_new_row, 3).NumberFormat = "dd-mm-yyyy"
        time_table.DataBodyRange.Cells(add_new_row, 4).value = split_combine_arrray(1)
        time_table.DataBodyRange.Cells(add_new_row, 4).NumberFormat = "hh:mm"
        time_table.DataBodyRange.Cells(add_new_row, set_offset_column_to_fill).value = split_combine_arrray(3)
        time_table.DataBodyRange.Cells(add_new_row, 6).value = join_array_value
        time_table.DataBodyRange.Cells(add_new_row, 7).value = counter_value
        add_new_row = add_new_row + 1
        
    End If
Next combine_array_index
     
' Release MEM
Set match_order_range = Nothing
Set time_table = Nothing
Set mainPage = Nothing
Set query_sheet = Nothing
Erase combine_value_array
Erase counter_value_join_array
Erase split_combine_arrray
Erase value_join_array

' Reset Application event
Application.ScreenUpdating = True
Application.EnableEvents = True
Exit Sub

ErrorHandler:
    ' Re-enable events, even if an error occurs
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Debug.Print Err.Description
End Sub
