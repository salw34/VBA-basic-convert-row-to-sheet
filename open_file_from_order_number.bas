'////////////////////////
' Usage
'=>> Call open_file_from_order_number(order_number)
'////////////////////////

Sub open_file_from_order_number(ByRef order_number As Variant)
' Call this sub by order number
' order_number is referance value for find file name in your folder.

Dim source_data As Worksheet
Dim PDF_path As String
Dim lastRow As Long, rowNum As Long
Dim sub_system As String, order_type As String
Dim year_two_chars As String
Dim ProgramFilesPath As String, FoxitPhantomPDFPath As String
Dim file_type as String 
Dim file_name_extension as String
Dim main_folder as String
Dim message_popup_box as String

Set source_data = ThisWorkbook.Sheets("source_sheet_for_reference")
lastRow = source_data.Cells(source_data.Rows.Count, "A").End(xlUp).Row

' Find row number that maches with order_number value
On Error Resume Next
    rowNum = Application.Match(order_number, source_data.Range("A1:A" & lastRow), 0)
On Error GoTo 0

' Validate if not found rowNum
If rowNum <= 0 Then
    Set source_data = Nothing
    Exit Sub
End If

' Get folder name from sub-system and order type
' for example, column B is contain sub-system, column C is contain order type.
sub_system = source_data.Range("B" & rowNum).value & "\"
order_type = source_data.Range("C" & rowNum).value & "\"

year_four_chars = Format(Date, "yyyy") & "\"
main_folder = "Dir:\Main folder\"
file_type = ".pdf"
file_name_extension = order_number  & file_type
' To use main_folder as macro file folder
' >> main_folder = ThisWorkbook.path
' order_type is subfolder of sub_system and sub_system is subfolders of main folder.
' Ex: Dir:\Main folder\sub_system\year_four_chars\order_type\.
' Create file path of file that you want to open.
PDF_path = main_folder & sub_system  & year_two_chars & order_type &  & file_name_extension

' This code use FoxitPhantomPDF as main file reader. of you want to use 
If Dir(PDF_path) = file_name_extension Then

    If Environ("ProgramFiles(x86)") <> "" Then
        ' 32-bit program on 64-bit Windows
        ProgramFilesPath = Environ("ProgramFiles(x86)")
    Else
        '32 bit Windows
        ProgramFilesPath = Environ("ProgramFiles")
    End If

    FoxitPhantomPDFPath = ProgramFilesPath & "\Foxit Software\Foxit PhantomPDF\FoxitPhantomPDF.exe"

    'Foxit Directory
    If Dir(FoxitPhantomPDFPath) = "FoxitPhantomPDF.exe" Then
        ' Use Shell function to execute FoxitphantomPDF and open pdf file.
        Shell Chr(34) & FoxitPhantomPDFPath & Chr(34) & " " & Chr(34) & PDF_path & Chr(34), vbNormalFocus
    Else

        ' Use default application for PDF on your windows.
        Dim Shex As Object
        Set Shex = CreateObject("Shell.Application")
        Shex.Open (PDF_path)
        
        ' Reset Object when file already opened.
        Set Shex = Nothing
    End If
Else

    message_popup_box = PDF_path & " was not found." & vbNewLine & vbNewLine & "Find and open it manually"
    MsgBox message_popup_box, vbOKOnly,"Infomation!", "PDF file was not found!"
    
End If

message_popup_box = ""
PDF_path = ""
Set source_data = Nothing

End Sub
