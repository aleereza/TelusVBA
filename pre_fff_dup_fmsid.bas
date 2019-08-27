Attribute VB_Name = "pre_fff_dup_fmsid"
'Module 4 - Handles duplicate FMSIDs with compare to FFF file
'to do: find missing FMSIDa


Sub find_duplicate_fmsid() 'check all records in "main" sheet

    'define workbook and worksheets
    Dim wb As Workbook
    Set wb = Workbooks("PRE-FFF.xlsm")
    Dim ws As Worksheet
    Set ws = wb.Sheets("DUP FMSID")
    Dim main_ws As Worksheet
    Set main_ws = wb.Sheets("main")
    
    Dim fff_wb As Workbook
    Set fff_wb = Workbooks("FFF Data.xlsx")
    Dim fff_ws As Worksheet
    Set fff_ws = fff_wb.Worksheets(1)
    
    'first and last row of main
    Dim fr As Long
    Dim lr As Long
    fr = 5
    lr = main_ws.Cells(Rows.Count, 10).End(xlUp).row
    
    Dim r1 As Long
    Dim row As Long 'current row in DUP FMSID sheet
    row = 2
    Dim fmsid As String
    Dim foundcell As Range
        
    'copy and count if
    For r1 = fr To lr
        fmsid = main_ws.Cells(r1, 3).Value
        Call copy_fff2dup(main_ws, r1, ws, row)
        If fmsid <> "" Then
            ws.Cells(row, "O").Value = Application.WorksheetFunction.CountIf(fff_ws.Range("C:C"), fmsid)
            If ws.Cells(row, "O").Value >= 1 Then 'if FMSID founded in FFF
                row = row + 1
                Set foundcell = fff_ws.Range("C:C").Find(What:=fmsid)
                Call copy_fff2dup(fff_ws, foundcell.row, ws, row)
            End If
        Else
            ws.Cells(row, "O").Value = 0
        End If
        Range(Cells(row, 1), Cells(row, 15)).Borders(xlEdgeBottom).Weight = xlThin
        row = row + 1
    Next r1
    
    'get rows from fff
    
    

End Sub
Sub copy_allfff2dup()

    'define workbook and worksheets
    Dim wb As Workbook
    Set wb = Workbooks("PRE-FFF.xlsm")
    Dim ws As Worksheet
    Set ws = wb.Sheets("DUP FMSID")
    
    Dim fff_wb As Workbook
    Set fff_wb = Workbooks("FFF Data.xlsx")
    Dim fff_ws As Worksheet
    Set fff_ws = fff_wb.Worksheets(1)
    
    'first and last row of fff
    Dim fff_fr As Long
    Dim fff_lr As Long
    fff_fr = 2
    fff_lr = fff_ws.Cells(Rows.Count, 2).End(xlUp).row
    
    Dim r1 As Long
    For r1 = fff_fr To fff_lr
        Call copy_fff2dup(fff_ws, r1, ws, r1)
    Next r1
    
    MsgBox ("Copy Completed!")

End Sub
Sub copy_fff2dup(fff_ws As Worksheet, r1 As Long, ws As Worksheet, r2 As Long)
    ws.Cells(r2, "A").Value = fff_ws.Cells(r1, "A").Value 'Note
    ws.Cells(r2, "B").Value = fff_ws.Cells(r1, "B").Value 'FIPUID
    ws.Cells(r2, "C").Value = fff_ws.Cells(r1, "C").Value 'FMSID
    ws.Cells(r2, "D").Value = fff_ws.Cells(r1, "D").Value 'Suite
    ws.Cells(r2, "E").Value = fff_ws.Cells(r1, "E").Value 'Civic
    ws.Cells(r2, "F").Value = fff_ws.Cells(r1, "F").Value 'Street
    ws.Cells(r2, "G").Value = fff_ws.Cells(r1, "I").Value 'City
    ws.Cells(r2, "H").Value = fff_ws.Cells(r1, "J").Value 'Network
    ws.Cells(r2, "I").Value = fff_ws.Cells(r1, "M").Value 'Date
    ws.Cells(r2, "J").Value = fff_ws.Cells(r1, "O").Value 'Partner
    ws.Cells(r2, "K").Value = fff_ws.Cells(r1, "Q").Value 'Type
    ws.Cells(r2, "L").Value = fff_ws.Cells(r1, "U").Value 'B Name
    ws.Cells(r2, "M").Value = fff_ws.Cells(r1, "V").Value 'B Class
    ws.Cells(r2, "N").Value = fff_ws.Cells(r1, "W").Value 'H class
End Sub

Function isfound(fff_ws As Worksheet, r As Long, ws As Worksheet) As Boolean 'check if FMSID in row r of fff is already found as duplicate
    
    'first and last row of ws (DUP FMSID)
    Dim fr As Long
    Dim lr As Long
    fr = 2
    lr = ws.Cells(Rows.Count, 2).End(xlUp).row
    If lr < 2 Then
        isfound = False
        Exit Function
    End If
    For row = fr To lr
        If fff_ws.Cells(r, "C").Value = ws.Cells(row, "C").Value Then
            isfound = True
            Exit Function
        End If
    Next row
    isfound = False
    
End Function

Sub place_dup_in_fff()

    'define workbook and worksheets
    Dim wb As Workbook
    Set wb = Workbooks("PRE-FFF.xlsm")
    Dim ws As Worksheet
    Set ws = wb.Sheets("DUP FMSID")
    
    Dim fff_wb As Workbook
    Set fff_wb = Workbooks("FFF Data.xlsx")
    Dim fff_ws As Worksheet
    Set fff_ws = fff_wb.Worksheets(1)
    
    'first and last row of fff
    Dim fff_fr As Long
    Dim fff_lr As Long
    fff_fr = 2
    fff_lr = fff_ws.Cells(Rows.Count, 2).End(xlUp).row
    
    'first and last row of dup fmsid
    Dim fr As Long
    Dim lr As Long
    fr = 2
    lr = ws.Cells(Rows.Count, 2).End(xlUp).row
    
    Dim r1 As Long
    Dim fipuid As String
    Dim foundcell As Range
    
    For r1 = fr To lr
        If ws.Cells(r1, "A").Value = "DUP" Then
            fipuid = ws.Cells(r1, "B").Value
            Set foundcell = fff_ws.Range("B:B").Find(What:=fipuid)
            If Not foundcell Is Nothing Then
                fff_ws.Cells(foundcell.row, "A").Value = "DUP"
            Else
                MsgBox ("FIPUID not found")
            End If
        End If
    Next r1
       
End Sub

Sub clear_dup_sheet()

    'define workbook and worksheets
    Dim wb As Workbook
    Set wb = Workbooks("PRE-FFF.xlsm")
    Dim ws As Worksheet
    Set ws = wb.Sheets("DUP FMSID")
    
    'first and last row of dup fmsid
    Dim fr As Long
    Dim lr As Long
    Dim fc As Long
    Dim lc As Long
    fr = 2
    lr = ws.Cells(Rows.Count, 3).End(xlUp).row
    fc = 1
    lc = 15
    
    ws.Range(Cells(fr, fc), Cells(lr, lc)).Clear
    
End Sub
