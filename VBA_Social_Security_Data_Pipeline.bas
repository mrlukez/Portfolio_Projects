Attribute VB_Name = "d_ten_eleven_folder"
Option Explicit
Sub file_loop()
'
'Loop through a folder to extract data.
'
'
'
'Make dialog box to assign the path of the folder to a variable.
Dim FileDir As String
Dim FiletoList As String
Dim open_book As Workbook
Application.ScreenUpdating = False

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Please select a folder"
        .ButtonName = "Pick Folder"

        If .Show = 0 Then
            MsgBox "Nothing was selected."
            Exit Sub
        Else
            FileDir = .SelectedItems(1) & "\"
        End If
    End With
    
'Make destination workbook
Workbooks.Add
ActiveWorkbook.SaveAs FileDir & "compiled_data"


'Loop through workbooks in file.
Dim combo As String
    
    combo = FileDir & "*xls*"
    
    FiletoList = VBA.FileSystem.DIR(combo)

    Do Until FiletoList = ""
        Set open_book = Application.Workbooks.Open(FileDir & FiletoList)
'        MsgBox FiletoList
        Call prepare_workbook

        If Not ActiveWorkbook.name Like "*compiled_data*" Then ActiveWorkbook.Close True

        FiletoList = VBA.FileSystem.DIR
    Loop
Application.ScreenUpdating = True

End Sub
Sub prepare_workbook()
'
'Prepare the workbook.
'
'
'
'Identify the string within the worksheet name to condition deletion.
Application.ScreenUpdating = False
Const delete_me As String = "Table 3*"

'Loop through worksheets to delete upon condition.
Application.DisplayAlerts = False
Dim Worksheet As Worksheet
For Each Worksheet In ActiveWorkbook.Worksheets
    
    If ActiveWorkbook.name = "compiled_data.xlsx" Then Exit Sub
    
    If Not Worksheet.name Like delete_me Then
        'Worksheet.range("a2").Interior.Color = vbYellow
        Worksheet.Delete
    End If
    
Next Worksheet
Application.DisplayAlerts = True

'Add sheet to destination workbook here.
Dim workbook_nm As String

    workbook_nm = ActiveWorkbook.name
    
    With Workbooks("compiled_data")
        .Worksheets.Add
        .ActiveSheet.name = workbook_nm
        .ActiveSheet.range("a2").Value = "state"
        .ActiveSheet.range("b2").Value = "county"
        .ActiveSheet.range("c2").Value = "category"
        .ActiveSheet.range("d2").Value = "population"
        .ActiveSheet.range("e2").Value = "workbook"
    End With
'Loop through worksheets and transform data.
For Each Worksheet In ActiveWorkbook.Worksheets
    
    Worksheet.Activate
    
    Call transform_data

Next Worksheet

Application.ScreenUpdating = True

End Sub
Sub transform_data()
'
'Transform original state data.
'
'
'
'Conditionally delete the county/counties line.
If range("c6").Value Like "C*" Then
    Rows(6).Delete
End If

'Conditionally delete the independent city / cities line.
Dim row_city As Long
Dim is_empty_col As Long
    
    'Avoid blank worksheets.
    is_empty_col = range("a7").End(xlDown).row
    If is_empty_col = 1048576 Then Exit Sub
    
    row_city = range("a7").End(xlDown).row + 1
        
    If range("c" & row_city).Value Like "I*" Then
        Rows(row_city).Delete
    End If
    
'Make array for county names.
Dim county_list() As String
Dim row_length As Long

    row_length = range("a6").End(xlDown).row - 5
    
ReDim county_list(1 To row_length)
Dim j As Long
    
    For j = 1 To row_length
        county_list(j) = range("a" & j + 5).Value
    Next j

'Make array for the categories from the original state data.
Dim headers_list(1 To 9) As String

    headers_list(1) = "ANSI code"
    headers_list(2) = "total"
    headers_list(3) = "aged"
    headers_list(4) = "blind and disabled"
    headers_list(5) = "under 18"
    headers_list(6) = "18-64"
    headers_list(7) = "65 or older"
    headers_list(8) = "SSI recipients also receiving OASDI"
    headers_list(9) = "amount of payments"

'Make array for final table.
Dim state_data() As Variant
Dim i As Integer 'navigate final array table rows
Dim r As Integer 'navigate original data rows
Dim c As Byte 'navigate original data columns
Dim end_row_array As Long 'holds array length
Dim state_name As String 'holds state name
Dim end_row_dest As Long 'holds destination array length
Dim workbook_nm As String 'place workbook name in a variable for multiple reference

    end_row_array = row_length * 9
    'For when "b5" has some value other than needed
    If Len(range("b5").Value) <= 7 Then Exit Sub
    
    state_name = Right(range("b5"), Len(range("b5").Value) - 7)
    workbook_nm = ActiveWorkbook.name
    
ReDim state_data(1 To end_row_array, 1 To 5)

            i = 1
    
    For r = 6 To row_length + 5
        For c = 3 To 11
            state_data(i, 1) = state_name 'column 1 will have the state name
            state_data(i, 2) = county_list(r - 5) 'column 2 will have the county name
            state_data(i, 3) = headers_list(c - 2) 'column 3 will have the heading from original data
            state_data(i, 4) = Cells(r, c).Value 'column 4 will be population
            state_data(i, 5) = workbook_nm 'column 5 will have the workbook name
            
            i = i + 1
            
        Next c
    Next r

'Asign final array to destination range.
Dim top_left_dest_row As Long
Dim bottom_right_dest_row As Long

    With Workbooks("compiled_data").Worksheets(workbook_nm)
    
'For the first entry
        If .range("a3").Value = "" Then
            
            top_left_dest_row = .range("a1").End(xlDown).row + 1
            bottom_right_dest_row = .range("e1").End(xlDown).row + end_row_array
            
            .range("a" & top_left_dest_row, "e" & bottom_right_dest_row).Value = state_data
        
        End If

'For every entry after the first
        If Not .range("a3").Value = "" Then
            
            top_left_dest_row = .range("a2").End(xlDown).row + 1
            bottom_right_dest_row = .range("e2").End(xlDown).row + end_row_array
            
            .range("a" & top_left_dest_row, "e" & bottom_right_dest_row).Value = state_data
        
        End If
        
    End With

End Sub
