Attribute VB_Name = "pre_fff_main"
'version 5.2 - 19082019 (Remove <CR> char added)
'Module 1 - main functions

Sub selectFDS()
    Dim TxtRng  As Range
    Set TxtRng = Range("D1")
    Call selectFile(TxtRng)
End Sub

Sub selectEnter()
    Dim TxtRng  As Range
    Set TxtRng = Range("F1")
    Call selectFile(TxtRng)
End Sub

Sub selectFMSID()
    Dim TxtRng  As Range
    Set TxtRng = Range("R1")
    Call selectFile(TxtRng)
End Sub

Sub selectFFF()
    Dim TxtRng  As Range
    Set TxtRng = Range("T1")
    Call selectFile(TxtRng)
End Sub

Sub selectUnassigned() 'sheet "healthcare"
    Dim TxtRng  As Range
    Set TxtRng = Range("A2")
    Call selectFile(TxtRng)
End Sub
Sub selectM2N() 'sheet "M2N"
    Dim TxtRng  As Range
    Set TxtRng = Range("A2")
    Call selectFile(TxtRng)
End Sub

Private Sub selectFile(TxtRng As Range)

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks("PRE-FFF.xlsm")
    Set ws = wb.Sheets("main")
    
    Dim fd As Office.FileDialog
    Dim fileName As String

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

        .AllowMultiSelect = False

        ' Set the title of the dialog box.
        .Title = "Please select the File."

        ' Clear out the current filters, and add our own.
        .Filters.Clear
        .Filters.Add "All Files", "*.*"

        ' Show the dialog box. If the .Show method returns True, the
        ' user picked at least one file. If the .Show method returns
        ' False, the user clicked Cancel.
        If .Show = True Then
        fileName = .SelectedItems(1) 'replace txtFileName with your textbox
        TxtRng.Value = fileName
        End If
   End With
End Sub

Sub fillFFF(to_enter_row As Integer)
    'at the time of running:
    'Missing FMSID file should be open
    
    'PRE-FFF
    Dim wb As Workbook
    Set wb = Workbooks("PRE-FFF.xlsm")
    Dim ws As Worksheet
    Set ws = wb.Sheets("main")
    Dim to_enter_ws As Worksheet
    Set to_enter_ws = wb.Sheets("to_enter")
    Dim m2n_ws As Worksheet
    Set m2n_ws = wb.Sheets("M2N")
    
    Dim main_fr As Integer 'first row and last row of PRE-FFF
    Dim main_lr As Integer
    main_fr = ws.Range("H2").Value
    
    'constant parameters
    Dim fds_folder_path As String
    fds_folder_path = to_enter_ws.Range("V4").Value
    Dim fds_filename_col_num As Integer
    Dim type_col_num As Integer
    fds_filename_col_num = 7
    type_col_num = 13
    
    'web fds constants
    Dim webfds_suite_col_num As Integer
    Dim webfds_civic_col_num As Integer
    Dim webfds_street_col_num As Integer
    Dim webfds_class_col_num As Integer
    Dim webfds_business_col_num As Integer
    Dim webfds_fip_col_num As Integer
    Dim webfds_demarc_col_num As Integer
    Dim webfds_comment_col_num As Integer
    Dim webfds_firstRow As Integer
    Dim webfds_lastRow As Integer
    
    webfds_suite_col_num = 1
    webfds_civic_col_num = 2
    webfds_street_col_num = 3
    webfds_class_col_num = 5
    webfds_business_col_num = 6
    webfds_fip_col_num = 12
    webfds_demarc_col_num = 11
    webfds_comment_col_num = 16
    webfds_firstRow = 2
    
    'm2n constants
    Dim m2n_suite_col_num As Integer
    Dim m2n_civic_col_num As Integer
    Dim m2n_street_col_num As Integer
    Dim m2n_business_col_num As Integer
    Dim m2n_premtype_col_num As Integer
    Dim m2n_key_col_num As Integer
    Dim m2n_firstRow As Integer
    Dim m2n_lastRow As Integer
    Dim m2n_key_firstRow As Integer
    Dim m2n_key_lastRow As Integer
    Dim m2n_key As String
    
    m2n_suite_col_num = 1
    m2n_civic_col_num = 4
    m2n_street_col_num = 5
    m2n_business_col_num = 8
    m2n_premtype_col_num = 7
    m2n_key_col_num = 6
    m2n_firstRow = 3
    m2n_lastRow = m2n_ws.Cells(Rows.Count, m2n_street_col_num).End(xlUp).row
    
    Dim fds_wb As Workbook
    
    Dim fds_ws As Worksheet
    
    Dim TxtRng  As Range
    Dim fds_fileName As String 'FDS sheet excel file
    Dim fds_path As String
    Dim fms_fileName As String 'MTU-Missing-FMSID excel file
    Dim fms_path As String
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim findResult As Range
    
    'If an error occurs, use the error handling routine at the end of this file.
    On Error GoTo ErrorHandler
    
    '===== Fill the Pre FFF =====
    Dim destRow As Integer 'destination row in PRE-FFF
    destRow = main_fr 'start row to write in PRE-FFF
    Dim comment1 As String
    Dim comment2 As String
    Dim comment3 As String
        
    'fill for NGM
    If to_enter_ws.Cells(to_enter_row, type_col_num).Value = 2 Then
        'last row of m2n
        
        m2n_key = to_enter_ws.Cells(to_enter_row, "K").Value
        m2n_key_firstRow = m2n_ws.Range(m2n_ws.Cells(m2n_firstRow, m2n_key_col_num), m2n_ws.Cells(m2n_lastRow, m2n_key_col_num)).Find(What:=m2n_key).row
        'MsgBox ("m2n_key_firstRow: " & Str(m2n_key_firstRow))
        m2n_key_lastRow = m2n_ws.Range(m2n_ws.Cells(m2n_firstRow, m2n_key_col_num), m2n_ws.Cells(m2n_lastRow, m2n_key_col_num)).Find(What:=m2n_key, searchdirection:=xlPrevious).row
        'MsgBox ("m2n_key_lastRow: " & Str(m2n_key_lastRow))
        
        For r = m2n_key_firstRow To m2n_key_lastRow 'firs and last row for the key
            'copy apartment #
            ws.Cells(destRow, 4).Value = m2n_ws.Cells(r, m2n_suite_col_num).Value
            'copy house #
            ws.Cells(destRow, 5).Value = m2n_ws.Cells(r, m2n_civic_col_num).Value
            'copy street
            ws.Cells(destRow, 6).Value = m2n_ws.Cells(r, m2n_street_col_num).Value
            'copy business name
            If m2n_ws.Cells(r, m2n_business_col_num).Value = "" Then
                ws.Cells(destRow, 21).Value = "Unknown"
            Else
                ws.Cells(destRow, 21).Value = m2n_ws.Cells(r, m2n_business_col_num).Value
            End If
            'Premis Type
            ws.Cells(destRow, 17).Value = m2n_ws.Cells(r, m2n_premtype_col_num).Value
            'Partner
            ws.Cells(destRow, 15).Value = "FNGM"
            
            'copy health suite
            'If fds_ws.Cells(r, webfds_class_col_num).Value = "health" Then
            '    ws.Cells(destRow, 22).Value = "healthcare"
            '    ws.Cells(destRow, 23).Value = "?"
            'End If
            'copy fiber in prem
'            If fds_ws.Cells(r, webfds_fip_col_num).Value <> "Yes" Then
'                ws.Cells(destRow, 18).Value = 1
'                'copy comments
'                comment1 = fds_ws.Cells(r, webfds_fip_col_num).Value
'                comment2 = fds_ws.Cells(r, webfds_demarc_col_num).Value
'                comment3 = fds_ws.Cells(r, webfds_comment_col_num).Value
'                ws.Cells(destRow, 19).Value = comment1
'                If comment2 <> "" Then
'                    ws.Cells(destRow, 19).Value = ws.Cells(destRow, 19).Value & "|" & comment2
'                End If
'                If comment3 <> "" Then
'                    ws.Cells(destRow, 19).Value = ws.Cells(destRow, 19).Value & "|" & comment3
'                End If
'            End If

            'Conv Date
            ws.Cells(destRow, 13).Value = ws.Cells(1, 12).Value
            'Conv Source
            ws.Cells(destRow, 14).Value = Year(ws.Cells(destRow, 13).Value)
            
            
            'NW#
            ws.Cells(destRow, 10).Value = ws.Cells(1, 2).Value
            'City
            ws.Cells(destRow, 9).Value = ws.Cells(1, 8).Value
            
            destRow = destRow + 1
        Next r
    Else
        'fill for MxU
        
        'Open FDS file
        Dim fds_found_records As Integer 'number of found records in FDS
        'fds_path = ws.Range("D1").Value
        fds_path = fds_folder_path & to_enter_ws.Cells(to_enter_row, fds_filename_col_num).Value
        fds_fileName = fso.GetFileName(fds_path)
        Workbooks.Open fileName:=fds_path
        Set fds_wb = Workbooks(fds_fileName)
        'Set fds_ws = fds_wb.Sheets("POST BUILD")
        Set fds_ws = fds_wb.Worksheets(1) 'select firts sheet
        
        
        
        
        
        'two different approaches for two different type of FDS
        If to_enter_ws.Cells(to_enter_row, type_col_num).Value = 1 Then
            '#1 Web FDS
            
            webfds_lastRow = fds_ws.Cells(Rows.Count, webfds_street_col_num).End(xlUp).row 'last row of FDS
            
            For r = webfds_firstRow To webfds_lastRow 'firs and last row of FDS sheet
                If fds_ws.Cells(r, webfds_class_col_num).Value = "business" Or fds_ws.Cells(r, webfds_class_col_num).Value = "health" Or fds_ws.Cells(r, webfds_class_col_num).Value = "utility/spare" Then
                    'copy apartment #
                    If fds_ws.Cells(r, webfds_suite_col_num).Value <> "-" Then
                        ws.Cells(destRow, 4).Value = fds_ws.Cells(r, webfds_suite_col_num).Value
                    End If
                    'copy house #
                    ws.Cells(destRow, 5).Value = fds_ws.Cells(r, webfds_civic_col_num).Value
                    'copy street
                    ws.Cells(destRow, 6).Value = fds_ws.Cells(r, webfds_street_col_num).Value
                    'copy business name
                    ws.Cells(destRow, 21).Value = fds_ws.Cells(r, webfds_business_col_num).Value
                    'copy health suite
                    If fds_ws.Cells(r, webfds_class_col_num).Value = "health" Then
                        ws.Cells(destRow, 22).Value = "healthcare"
                        ws.Cells(destRow, 23).Value = "?"
                    End If
                    'copy fiber in prem
                    If fds_ws.Cells(r, webfds_fip_col_num).Value <> "Yes" Then
                        ws.Cells(destRow, 18).Value = 1
                        'copy comments
                        comment1 = fds_ws.Cells(r, webfds_fip_col_num).Value
                        comment2 = fds_ws.Cells(r, webfds_demarc_col_num).Value
                        comment3 = fds_ws.Cells(r, webfds_comment_col_num).Value
                        ws.Cells(destRow, 19).Value = comment1
                        If comment2 <> "" Then
                            ws.Cells(destRow, 19).Value = ws.Cells(destRow, 19).Value & "|" & comment2
                        End If
                        If comment3 <> "" Then
                            ws.Cells(destRow, 19).Value = ws.Cells(destRow, 19).Value & "|" & comment3
                        End If
                    End If
                    'Conv Date
                    ws.Cells(destRow, 13).Value = ws.Cells(1, 12).Value
                    'Conv Source
                    ws.Cells(destRow, 14).Value = Year(ws.Cells(destRow, 13).Value)
                    'Partner
                    ws.Cells(destRow, 15).Value = ws.Cells(1, 14).Value
                    'Premis Type
                    ws.Cells(destRow, 17).Value = ws.Cells(1, 16).Value
                    'NW#
                    ws.Cells(destRow, 10).Value = ws.Cells(1, 2).Value
                    'City
                    ws.Cells(destRow, 9).Value = ws.Cells(1, 8).Value
                    
                    destRow = destRow + 1
                End If
            Next r
               
        ElseIf to_enter_ws.Cells(to_enter_row, type_col_num).Value = 0 Then
            '#2 Excel FDS
            
            On Error Resume Next
            Set findResult = fds_ws.Cells.Find("street", Lookat:=xlWhole)
            If findResult Is Nothing Then
                'MsgBox "Couldn't find 'street', looking for 'street/avenue'"
                On Error GoTo ErrorHandler
                Set findResult = fds_ws.Cells.Find("street/avenue", Lookat:=xlWhole)
            End If
            'MsgBox "after search for street"
            'Set findResult = fds_ws.Cells.Find("street*")
            Dim rangeAddress As String
            Dim firstRow, lastRow, headerRow, col As Long
            Dim apt_col As Integer
            Dim mtu_col, hus_col, str_col, bus_col, hlt_col, fip_col As Integer
            str_col = findResult.Column
            rangeAddress = findResult.Address
            
            'MsgBox findResult.Column
            headerRow = findResult.row
            firstRow = findResult.row + 1 'first row of FDS
            lastRow = fds_ws.Cells(Rows.Count, findResult.Column).End(xlUp).row 'last row of FDS
            'MsgBox lastRow & firstRow
            
            'find MTU column
            On Error Resume Next
            Set findResult = fds_ws.Rows(headerRow).Find(What:="MTU*")
            If findResult Is Nothing Then
                MsgBox "Couldn't find 'MTU*'"
                'Exit Sub
            End If
            mtu_col = findResult.Column
            Set findResult = fds_ws.Rows(headerRow).Find(What:="APARTMENT*")
            If findResult Is Nothing Then
                Set findResult = fds_ws.Rows(headerRow).Find(What:="UNIT*")
                If findResult Is Nothing Then
                    MsgBox "Couldn't find 'APARTMENT*' or 'UNIT*'"
                End If
                'Exit Sub
            End If
            apt_col = findResult.Column
            Set findResult = fds_ws.Rows(headerRow).Find(What:="HOUSE*")
            If findResult Is Nothing Then
                Set findResult = fds_ws.Rows(headerRow).Find(What:="BUILDING*")
                If findResult Is Nothing Then
                    MsgBox "Couldn't find 'HOUSE*' or 'BUILDING*'"
                End If
                'Exit Sub
            End If
            hus_col = findResult.Column
            Set findResult = fds_ws.Rows(headerRow).Find(What:="Business*")
            If findResult Is Nothing Then
                MsgBox "Couldn't find 'Business*'"
                'Exit Sub
            End If
            bus_col = findResult.Column
            Set findResult = fds_ws.Rows(headerRow).Find(What:="Health*")
            If findResult Is Nothing Then
                MsgBox "Couldn't find 'Health*'"
                'Exit Sub
            End If
            hlt_col = findResult.Column
            Set findResult = fds_ws.Rows(headerRow).Find(What:="fibre in prem*")
            If findResult Is Nothing Then
                MsgBox "Couldn't find 'fibre in prem*'"
                'Exit Sub
            End If
            fip_col = findResult.Column
            Set findResult = fds_ws.Rows(headerRow).Find(What:="Fibre Demarc*")
            If findResult Is Nothing Then
                MsgBox "Couldn't find 'Fibre Demarc*'"
                'Exit Sub
            End If
            co1_col = findResult.Column
            Set findResult = fds_ws.Rows(headerRow).Find(What:="Comments (*")
            If findResult Is Nothing Then
                MsgBox "Couldn't find 'Comments (*'"
                'Exit Sub
            End If
            co2_col = findResult.Column
            On Error GoTo ErrorHandler
            'MsgBox "co1:" & co1_col & "co2:" & co2_col
            
            
            'Dim destRow, destCol As Integer 'destination row and column
            
            For r = firstRow To lastRow 'firs and last row of FDS sheet
                If fds_ws.Cells(r, mtu_col).Value = 1 Then
                    'copy apartment #
                    If fds_ws.Cells(r, apt_col).Value <> "-" Then
                        ws.Cells(destRow, 4).Value = fds_ws.Cells(r, apt_col).Value
                    End If
                    'copy house #
                    ws.Cells(destRow, 5).Value = fds_ws.Cells(r, hus_col).Value
                    'copy street
                    ws.Cells(destRow, 6).Value = fds_ws.Cells(r, str_col).Value
                    'copy business name
                    ws.Cells(destRow, 21).Value = fds_ws.Cells(r, bus_col).Value
                    'copy health suite
                    If LCase(fds_ws.Cells(r, hlt_col).Value) = "y" Or LCase(fds_ws.Cells(r, hlt_col).Value) = "yes" Then
                        ws.Cells(destRow, 22).Value = "healthcare"
                        ws.Cells(destRow, 23).Value = "?"
                    End If
                    'copy fiber in prem
                    If LCase(fds_ws.Cells(r, fip_col).Value) = "n" Or LCase(fds_ws.Cells(r, fip_col).Value) = "no" Then
                        ws.Cells(destRow, 18).Value = 1
                        'copy comments
                        comment1 = ""
                        comment2 = ""
                        If fds_ws.Cells(r, co1_col).Value <> "-" And fds_ws.Cells(r, co1_col).Value <> "OTHER ""Comments Required""" And fds_ws.Cells(r, co1_col).Value <> "" Then
                            comment1 = fds_ws.Cells(r, co1_col).Value
                            If fds_ws.Cells(r, co2_col).Value <> "-" Then
                                comment2 = fds_ws.Cells(r, co2_col).Value
                                ws.Cells(destRow, 19).Value = comment1 & "|" & comment2
                            Else
                                ws.Cells(destRow, 19).Value = comment1
                            End If
                        Else
                            If fds_ws.Cells(r, co2_col).Value <> "-" Then
                                comment2 = fds_ws.Cells(r, co2_col).Value
                                ws.Cells(destRow, 19).Value = comment2
                            End If
                        End If
                    End If
                    'FalconName (City)
                    'ws.Cells(destRow, 11).Value = ws.Cells(1, 8).Value
                    'FSA
                    'ws.Cells(destRow, 12).Value = ws.Cells(1, 10).Value
                    'Conv Date
                    ws.Cells(destRow, 13).Value = ws.Cells(1, 12).Value
                    'Conv Source
                    ws.Cells(destRow, 14).Value = Year(ws.Cells(destRow, 13).Value)
                    'Partner
                    ws.Cells(destRow, 15).Value = ws.Cells(1, 14).Value
                    'Premis Type
                    ws.Cells(destRow, 17).Value = ws.Cells(1, 16).Value
                    'NW#
                    ws.Cells(destRow, 10).Value = ws.Cells(1, 2).Value
                    'City
                    ws.Cells(destRow, 9).Value = ws.Cells(1, 8).Value
                    
                    
                    destRow = destRow + 1
                End If
            Next r
    
        End If
    
    fds_wb.Close SaveChanges:=False
    
    End If
        
    main_lr = destRow - 1
    fds_found_records = main_lr - main_fr + 1
    ws.Range("B2").Value = fds_found_records
    
    'write number of found records in FDS in to_enter sheet for later check
    to_enter_ws.Cells(to_enter_row, 9).Value = fds_found_records
    
    'write te main_lr as the new main_fr in Range("H2")
    ws.Range("H2").Value = main_lr + 1
    
    'Error Handling section.
ErrorHandler:
    Select Case Err.Number
        'Common error #1: file path or workbook name is wrong.
        Case 1004
            Application.ScreenUpdating = True
            MsgBox "The workbook could not be found in the path"
        Exit Sub
    
        'Common error #2: the specified text wasn't in the target workbook.
        Case 9, 91
            Application.ScreenUpdating = True
            MsgBox "The value was not found."
        Exit Sub
    
        'General case: turn screenupdating back on, and exit.
        Case Else
            Application.ScreenUpdating = True
        Exit Sub
    End Select

End Sub

Sub clear_records()
    'clear current data in main sheet and clear formattings
    'from row 5 to row Range("h2").value
    
    'definition of workbooks and worksheets
    Dim wb As Workbook
    Dim ws As Worksheet
    'PRE-FFF
    Set wb = Workbooks("PRE-FFF.xlsm")
    Set ws = wb.Sheets("main")
    
    Dim main_fr As Integer 'first row and last row of PRE-FFF
    Dim main_lr As Integer
    main_fr = 5
    main_lr = main_fr + ws.Range("H2").Value - 1
    With ws.Range(Cells(main_fr, 1), Cells(main_lr, 26))
        .Clear
    End With
    
    ws.Range("H2").Value = main_fr
    
End Sub

Sub find_existing(to_enter_row As Integer)
    'FFF Data should be open
        
    'definition of workbooks and worksheets
    Dim wb As Workbook
    Dim fff_wb As Workbook
    Dim ws As Worksheet
    Dim fff_ws As Worksheet
    Dim to_enter_ws As Worksheet
    'PRE-FFF
    Set wb = Workbooks("PRE-FFF.xlsm")
    Set ws = wb.Sheets("main")
    Set to_enter_ws = wb.Sheets("to_enter")
    'FFF
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fff_path = ws.Range("T1").Value
    fff_fileName = fso.GetFileName(fff_path)
    On Error Resume Next
    Set fff_wb = Workbooks(fff_fileName)
    On Error GoTo 0
    If fff_wb Is Nothing Then
        Workbooks.Open fileName:=fff_path
        Set fff_wb = Workbooks(fff_fileName)
    End If
    Set fff_ws = fff_wb.Worksheets(1)
    
    

    'first and last row of fff
    Dim fff_fr As Long
    Dim fff_lr As Long
    fff_fr = 2
    fff_lr = fff_ws.Cells(Rows.Count, 2).End(xlUp).row
    Dim num_found_records As Integer
    num_found_records = 0
    'Dim arrRows() As Integer
    For r1 = fff_fr To fff_lr
        If fff_ws.Cells(r1, 10).Value = ws.Range("B1").Value Then
            num_found_records = num_found_records + 1
        End If
    Next r1
    to_enter_ws.Cells(to_enter_row, 8).Value = num_found_records
    
    

End Sub


Sub fillall()

    'define workbook and worksheets
    Dim wb As Workbook
    Set wb = Workbooks("PRE-FFF.xlsm")
    Dim ws As Worksheet
    Set ws = wb.Sheets("main")
    Dim to_enter_ws As Worksheet
    Set to_enter_ws = wb.Sheets("to_enter")
    
    'going throug to_enter list
    Dim fr As Integer
    Dim lr As Integer
    fr = to_enter_ws.Range("V2").Value
    lr = to_enter_ws.Range("V3").Value
    
    Dim comunity As String
    
    Dim r As Integer
    For r = fr To lr
        
        'put the NW# in place
        ws.Range("B1").Value = to_enter_ws.Cells(r, 2).Value
        
        'put ISW partner in place
        If to_enter_ws.Cells(r, "M").Value = 2 Then
            ws.Range("N1").Value = "FNGM"
        Else
            ws.Range("N1").Value = to_enter_ws.Cells(r, 6).Value
        End If
        
        'put conversion date
        ws.Range("L1").Value = to_enter_ws.Cells(r, 1).Value
        
        'put city in place
        ws.Range("H1").Value = to_enter_ws.Cells(r, 12).Value
        
        Call find_existing(r)
        
        Call fillFFF(r)
        
        'change colors
        If to_enter_ws.Cells(r, 8).Value <> 0 Then
            to_enter_ws.Cells(r, 8).Interior.ColorIndex = 3 'red
        Else
            to_enter_ws.Cells(r, 8).Interior.ColorIndex = 4 'green
        End If
        
        If to_enter_ws.Cells(r, 5).Value = to_enter_ws.Cells(r, 9).Value Then
            to_enter_ws.Cells(r, 9).Interior.ColorIndex = 4
        Else
            to_enter_ws.Cells(r, 9).Interior.ColorIndex = 3
        End If
    Next r
    
    Call health_list_link
    Call remove_cr_char
       
End Sub

Sub compare()
'when there are already records in FFF with same NW#

    'PRE-FFF
    Set wb = Workbooks("PRE-FFF.xlsm")
    Set ws = wb.Sheets("main")
    Dim main_fr As Integer 'first row and last row of PRE-FFF
    Dim main_lr As Integer
    main_fr = 5
    main_lr = main_fr + ws.Range("B2").Value - 1
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'FFF
    fff_path = ws.Range("T1").Value
    fff_fileName = fso.GetFileName(fff_path)
    On Error Resume Next
    Set fff_wb = Workbooks(fff_fileName)
    On Error GoTo 0
    If fff_wb Is Nothing Then
        Workbooks.Open fileName:=fff_path
        Set fff_wb = Workbooks(fff_fileName)
    End If
    Set fff_ws = fff_wb.Worksheets(1)
    
    'first and last row
    fff_fr = 2
    fff_lr = fff_ws.Cells(Rows.Count, 2).End(xlUp).row
    Dim num_found_records As Integer
    num_found_records = 0
    'Dim arrRows() As Integer
    For r1 = fff_fr To fff_lr
        If fff_ws.Cells(r1, 10).Value = ws.Range("B1").Value Then
            num_found_records = num_found_records + 1
            For r2 = main_fr To main_lr
                If fff_ws.Cells(r1, 4).Value = ws.Cells(r2, 4).Value And fff_ws.Cells(r1, 5).Value = ws.Cells(r2, 5).Value Then
                    For c1 = 3 To 22
                        If fff_ws.Cells(r1, c1).Value = ws.Cells(r2, c1).Value Then
                            ws.Cells(r2, c1).Interior.ColorIndex = 4
                        Else
                            ws.Cells(r2, c1).Interior.ColorIndex = 3
                        End If
                    Next c1
                End If
            Next r2
        End If
    Next r1




End Sub

Sub dupfms()
    'check FMSID column and highlight duplicates
    Set wb = Workbooks("PRE-FFF.xlsm")
    Set ws_main = wb.Sheets("main")
    Dim FMSID_Range As Range
    Dim FMSID_Cell As Range
    Dim lr As Integer
    lr = ws_main.Range("B2").Value + 4
    Set FMSID_Range = ws_main.Range(Cells(5, 3), Cells(lr, 3))
    For Each FMSID_Cell In FMSID_Range
        If WorksheetFunction.CountIf(FMSID_Range, FMSID_Cell.Value) > 1 Then
            FMSID_Cell.Interior.ColorIndex = 3
        Else
            FMSID_Cell.Interior.ColorIndex = 4
        End If
    Next
End Sub


Sub updateHealth()

    'PRE-FFF
    Set wb = Workbooks("PRE-FFF.xlsm")
    Set ws_main = wb.Sheets("main")
    Set ws_hc = wb.Sheets("healthcare")
    Dim main_fr, main_lr As Integer 'first row and last row of PRE-FFF/main
    Dim hc_fr, hc_lr As Integer 'first row and last row of PRE-FFF/healthcare
    main_fr = 5
    'main_lr = main_fr + ws.Range("B2").Value - 1
    hc_fr = 6
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'FFF
    'FFF Data.xlsx should be open
    Set fff_wb = Workbooks("FFF Data.xlsx")
    Set fff_ws = fff_wb.Worksheets(1)
    Dim fff_fr, fff_lr As Long 'first row and last row of FFF Data
    fff_fr = 2
    fff_lr = fff_ws.Cells(Rows.Count, 2).End(xlUp).row
    
    'update file
    upd_path = ws_hc.Range("A2").Value
    upd_fileName = fso.GetFileName(upd_path)
    On Error Resume Next
    Set upd_wb = Workbooks(upd_fileName)
    On Error GoTo 0
    If upd_wb Is Nothing Then
        Workbooks.Open fileName:=upd_path
        Set upd_wb = Workbooks(upd_fileName)
    End If
    Set upd_ws = upd_wb.Worksheets(1)
    
    
    Dim upd_fr, upd_lr As Integer 'first row and last row of Unassigned
    upd_fr = 2
    upd_lr = upd_ws.Cells(Rows.Count, 1).End(xlUp).row
    
    
    Dim destRow As Integer 'where write in healthcare sheet
    destRow = hc_fr
    Dim fmsid As Long
    Dim found As Boolean
    Dim fmsid_col, subclass_col As Integer
    fmsid_col = ws_hc.Range("D2").Value
    subclass_col = ws_hc.Range("E2").Value
    For r1 = upd_fr To upd_lr
        ws_hc.Cells(destRow, 1).Value = upd_ws.Cells(r1, fmsid_col).Value
        ws_hc.Cells(destRow, 3).Value = upd_ws.Cells(r1, subclass_col).Value
        
        r2 = fff_fr
        found = False
        Do While r2 <= fff_lr And found = False
            If upd_ws.Cells(r1, fmsid_col).Value = fff_ws.Cells(r2, 3).Value Then
                found = True
                ws_hc.Cells(destRow, 4).Value = fff_ws.Cells(r2, 23).Value
            End If
            r2 = r2 + 1
        Loop
        destRow = destRow + 1
    Next r1
    
    
    
    
        'Error Handling section.
ErrorHandler:
    Select Case Err.Number
        'Common error #1: file path or workbook name is wrong.
        Case 1004
            Application.ScreenUpdating = True
            MsgBox "The workbook could not be found in the path"
        Exit Sub
    
        'Common error #2: the specified text wasn't in the target workbook.
        Case 9, 91
            Application.ScreenUpdating = True
            MsgBox "The value was not found."
        Exit Sub
    
        'General case: turn screenupdating back on, and exit.
        Case Else
            Application.ScreenUpdating = True
        Exit Sub
    End Select

End Sub

Sub health_list_link()  'add health list and google map link in main sheet for health units

    'define workbook and worksheets
    Dim wb As Workbook
    Set wb = Workbooks("PRE-FFF.xlsm")
    Dim ws_main As Worksheet
    Set ws_main = wb.Sheets("main")
    
    'find first/last row
    Dim main_fr As Long
    Dim main_lr As Long
    main_fr = 5
    main_lr = ws_main.Range("H2").Value - 1
    
    For r = main_fr To main_lr
        If ws_main.Cells(r, 23).Value = "?" Then
            With ws_main.Cells(r, 23).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                Formula1:="Acupuncturist,Association/Support,Audiologist,Chiropractor,Dentist/Orthodontist/Denturist,Diagnostic Imaging,EXTENDED CARE FACILITY,MEDICAL CLINIC,Massage Therapist,Medical Laboratory,Naturopath,Occupational Therapist,Optician/Optometrist,Other,Pharmacy,Physician,Physiotherapist,Podiatrist,Psychologist,REHABILITION MEDICINE,RESPIROLOGIST,Speech Therapist"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = True
            End With
            With ws_main
                .Hyperlinks.Add _
                Anchor:=.Cells(r, 24), _
                Address:="http://maps.google.com/?q=" & .Cells(r, 5).Value & " " & .Cells(r, 6).Value & " " & .Cells(r, 9).Value & " " & .Cells(r, 21).Value, _
                TextToDisplay:=.Cells(r, 21).Value
            End With
        End If
    Next r
    
End Sub

Sub remove_cr_char()
' to remove <CR> character, somethimes there is one at the end of the street name or civic
' it causes problem in FMSID search

    'define worksheet
    Set ws_main = ThisWorkbook.Sheets("main")
    
    'find first/last row
    Dim fr As Long
    Dim lr As Long
    Dim fc As Long
    Dim lc As Long
    fr = 5
    lr = ws_main.Range("H2").Value - 1
    fc = 1
    lc = 23
    
    'remove <CR> from every cell
    ws_main.Range(ws_main.Cells(fr, fc), ws_main.Cells(lr, lc)).Replace _
    What:=Chr(13), Replacement:=""

End Sub

Sub prepare_to_enter()
'to put data which copied from domo "MTU records to enter_with link" in the format of "to_enter" sheet

    'define workbook and worksheets
    Dim wb As Workbook
    Set wb = Workbooks("PRE-FFF.xlsm")
    Dim ws As Worksheet
    Set ws = wb.Sheets("to_enter")
    
    'find first/last row
    Dim main_lr As Long
    Dim raw_lr As Long
    Dim main_fr As Integer
    Dim raw_fr As Integer
    
    main_fr = 2
    raw_fr = 12
    main_lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    raw_lr = ws.Cells(ws.Rows.Count, "U").End(xlUp).row
    If main_lr < 2 Then
        main_lr = 2
    End If
    
    'clear main table
    With ws.Range(ws.Cells(main_fr, 1), ws.Cells(main_lr, 15))
        .Clear
    End With
    
    'fill main table
    Dim r As Integer
    Dim main_row As Integer
    main_row = main_fr
    For r = raw_fr To raw_lr
        ws.Cells(main_row, "A").Value = ws.Cells(r, "U").Value
        ws.Cells(main_row, "B").Value = ws.Cells(r, "V").Value
        ws.Cells(main_row, "C").Value = ws.Cells(r, "W").Value
        ws.Cells(main_row, "D").Value = 0
        ws.Cells(main_row, "E").Value = ws.Cells(r, "W").Value
        ws.Cells(main_row, "F").Value = ws.Cells(r, "X").Value
        ws.Cells(main_row, "K").Value = ws.Cells(r, "Y").Value 'Jira key
        ws.Cells(main_row, "L").Value = ws.Cells(r, "Z").Value
        
        With ws
            .Hyperlinks.Add _
            Anchor:=.Cells(main_row, "J"), _
            Address:=.Range("V5").Value & .Cells(main_row, "K").Value, _
            TextToDisplay:=.Cells(main_row, "K").Value
        End With
        
        With ws
            .Hyperlinks.Add _
            Anchor:=.Cells(main_row, "N"), _
            Address:=.Range("V6").Value & .Cells(main_row, "K").Value, _
            TextToDisplay:=.Cells(main_row, "K").Value
        End With
        
        main_row = main_row + 1
    Next r
       
End Sub

