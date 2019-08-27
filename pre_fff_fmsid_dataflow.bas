Attribute VB_Name = "pre_fff_fmsid_dataflow"
'Module 5 - Handles FMSID dataflow related sheets, ie FMSID_df_input and FMSID_df_output

Sub fill_FMSID_df_input()

    'define workbook and worksheets
    Dim ws_main As Worksheet
    Set ws_main = ThisWorkbook.Sheets("main")
    Dim ws_input As Worksheet
    Set ws_input = ThisWorkbook.Sheets("FMSID_df_input")
    
    'first and last row of main sheet data
    Dim main_fr As Long
    Dim main_lr As Long
    main_fr = 5
    main_lr = ws_main.Range("H2").Value - 1
    
    'first and last row of input sheet data
    Dim fr As Long
    Dim lr As Long
    fr = 2
    lr = ws_input.Cells(Rows.Count, 1).End(xlUp).row
    If lr < fr Then
        lr = fr
    End If
        
    'clear input sheet except first row
    ws_input.Range(ws_input.Cells(fr, 1), ws_input.Cells(lr, 6)).Clear
    
    'copy address data from main sheet
    Dim c1_array As Variant
    c1_array = Array(2, 4, 5, 6, 9) 'column numbers
    Dim c2_array As Variant
    c2_array = Array(1, 3, 4, 5, 6)
    Call copy_table(ws_main, ws_input, c1_array, c2_array, 5, 2)
    
    'clean street name
    lr = ws_input.Cells(Rows.Count, 1).End(xlUp).row
    For i = fr To lr
        ws_input.Cells(i, 5).Value = clean_street_name(ws_input.Cells(i, 5).Value)
        ws_input.Cells(i, 5).Value = remove_ordinal_indicator(ws_input.Cells(i, 5).Value)
    Next i
    
    'fill address column by concatenating suite, civic, street and city
    For i = fr To lr
        ws_main_row = main_fr - fr + i
        ws_input.Cells(i, 2).Value = ws_main.Cells(ws_main_row, 4).Value & ", " & _
            ws_main.Cells(ws_main_row, 5).Value & ", " & _
            ws_main.Cells(ws_main_row, 6).Value
    Next i
    
    'export FMSID_df_input to a new xlsx file for domo
    Dim file_name As String
    file_name = "FMSID_df_input"
    Dim input_range As Range
    Set input_range = ws_input.Range(ws_input.Cells(1, 1), ws_input.Cells(lr, 6))
    Call export_range(input_range, file_name)

End Sub

Sub fill_FMSID_df_output()
    
    'import data from output file
    Dim ws_to_enter As Worksheet
    Set ws_to_enter = ThisWorkbook.Sheets("to_enter")
    Dim ws_output As Worksheet
    Set ws_output = ThisWorkbook.Sheets("FMSID_df_output")
    
    Dim wb_output_file As Workbook
    Dim ws_output_file As Worksheet
    Dim output_fileName As String
    Dim output_path As String 'including file name
    output_fileName = "FMS_LPDS_output.xlsx"
    output_path = ws_to_enter.Range("V4").Value & output_fileName
    Workbooks.Open fileName:=output_path
    Set wb_output_file = Workbooks(output_fileName)
    Set ws_output_file = wb_output_file.Worksheets(1) 'select firts sheet
        
    'copy FMSID_df_output data from domo output file
    Dim c1_array As Variant
    c1_array = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19) 'column numbers
    Dim c2_array As Variant
    c2_array = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)
    Call copy_table(ws_output_file, ws_output, c1_array, c2_array, 2, 2)
                
    wb_output_file.Close SaveChanges:=False

End Sub

Sub update_sort_and_group()

    Dim ws_output As Worksheet
    Set ws_output = ThisWorkbook.Sheets("FMSID_df_output")
    
    Call sort_and_group(ws_output, 1, 1, 19)

End Sub

Private Function clean_street_name(s As String) As String
'This function gets street name as a string and remove unusable words like directions, etc.

    Dim directions_array As Variant
    directions_array = Array("ne", "nw", "se", "sw", "north", "n", "east", "e", "west", "w", "south", "s")
    Dim streettype_array As Variant
    streettype_array = Array("road", "rd", "street", "st", "avenue", "ave", "way", "trail", "highway", "hwy", "drive", "dr", "blvd", "place", "pl", "mt")
    
    s = Replace(s, "-", " ")
    s = Replace(s, ".", " ")
    s = Replace(s, Chr(13), "") 'remove <CR> character, somethimes there is one at the end of the street name

    s = remove_word_array(s, directions_array)
    s = remove_word_array(s, streettype_array)
    
    'get rid of leading and trailing spaces
    s = Trim(s)
    
    clean_street_name = s
       
End Function

Private Function remove_word_array(s As String, a As Variant) As String
'This function gets a string and an array of words, then remove all words in that array from the string
    'add space to begining and end of string
    s = LCase(" " & s & " ")
    For Each element In a
        s = Replace(s, LCase(" " & element & " "), " ") 'case insensitive
    Next element
    remove_word_array = s
End Function

Private Function remove_ordinal_indicator(s As String) As String
    s = LCase(" " & s & " ")
    Dim ordinal_indicator_array As Variant
    ordinal_indicator_array = Array("1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "0th", "1th", "2th", "3th")
    Dim ordinal_indicator_removed_array As Variant
    ordinal_indicator_removed_array = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "1", "2", "3")
    For i = 0 To UBound(ordinal_indicator_array)
        s = Replace(s, LCase(ordinal_indicator_array(i) & " "), ordinal_indicator_removed_array(i)) 'case insensitive
    Next i
    'get rid of leading and trailing spaces
    s = Trim(s)
    remove_ordinal_indicator = s
End Function

Sub copy_table(s1 As Worksheet, s2 As Worksheet, c1 As Variant, c2 As Variant, r1 As Long, r2 As Long)
'copy columns of c1 from sheet s1 starts from row r1 to its last row, to sheet s2
    
    'see if c1 and c2 have same size
    If UBound(c1) <> UBound(c2) Then
        ' Raise the exception
        Err.Raise Number:=513, Description:="Arrays are not same size"
    End If
    
    'first column in source
    Dim source_fc As Long
    source_fc = c1(0)
    'last row in source
    Dim source_lr As Long
    source_lr = s1.Cells(Rows.Count, source_fc).End(xlUp).row
    
    'clear dest
    'todo
    
    'loop to copy
    Dim source_row As Long
    Dim dest_row As Long
    Dim source_col As Long
    Dim dest_col As Long
    Dim dr As Long
    dr = r2 - r1 'row difference
    For source_row = r1 To source_lr
        dest_row = source_row + dr
        For column_index = 0 To UBound(c1)
            source_col = c1(column_index)
            dest_col = c2(column_index)
            s2.Cells(dest_row, dest_col).Value = s1.Cells(source_row, source_col).Value
        Next column_index
    Next source_row
End Sub

Public Sub export_range(input_range As Range, file_name As String)
    'copy data from FMSID_df_input sheet
    input_range.Copy

    file_path = ThisWorkbook.Path
    Dim wb As Workbook
    
    Set wb = Workbooks.Add
    wb.SaveAs fileName:=file_path & "\" & file_name 'complete file path

    'wb.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteValues
    
    input_range.Copy _
    Destination:=wb.Sheets(1).Range("A1")
    
    wb.Close SaveChanges:=True
End Sub

Private Sub sort_and_group(ws As Worksheet, c As Long, fc As Long, lc As Long)

    'fc and lc: first and last column numbers
    'c: column number of id column (column we want to sort and group based on that)
        
    'sort
    ws.Range(ws.Cells(, fc).EntireColumn, ws.Cells(, lc).EntireColumn).Sort Key1:=ws.Cells(1, c), Order1:=xlAscending, Header:=xlYes
    
    'group
    'first and last row of input sheet data
    Dim fr As Long
    Dim lr As Long
    fr = 2
    lr = ws.Cells(Rows.Count, c).End(xlUp).row
    If lr < fr Then
        lr = fr
    End If
    
    'clear all borders
    ws.Range(Cells(fr, 1), Cells(lr, lc)).Borders.LineStyle = xlLineStyleNone
    
    For r = fr + 1 To lr
        If ws.Cells(r, c).Value <> ws.Cells(r - 1, c).Value Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, lc)).Borders(xlEdgeTop).LineStyle = xlContinuous
            ws.Range(ws.Cells(r, 9), ws.Cells(r, 16)).Interior.ColorIndex = xlNone
        Else
            ws.Range(ws.Cells(r - 1, 9), ws.Cells(r, 16)).Interior.ColorIndex = 6
        End If
        
    Next r
    
End Sub
