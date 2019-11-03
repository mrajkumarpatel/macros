Option Explicit

Sub combine_spread_sheets()
 
    Dim fso As Object, folder As Object, file As Object
    Dim wb As Workbook, o_wb As Workbook, ws As Worksheet, o_ws As Worksheet
    Dim files_path As String, wb_day1 As Workbook, wb_day2 As Workbook
    Dim sheets_in_day1 As Collection, sheets_in_day2 As Collection
    Dim o_file As String
    Dim file1 As String, file2 As String
    Dim no_of_rows As Long, o_no_of_rows As Long
    Dim no_of_columns As Long, last_column As String

    
    files_path = "C:\Users\abc\python\projects\dev\combine_xls\resources"
    o_file = files_path & Application.PathSeparator & "Data Collated.xlsx"
    
    Call delete_file(o_file)
    
    Application.Visible = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(files_path)
    Set o_wb = Workbooks.Add

    For Each file In folder.Files
     
        If file.name Like "*.xls?" Then
            Set wb = Workbooks.Open(files_path & Application.PathSeparator & file.name, ReadOnly:=True)
            For Each ws In wb.Worksheets
                If sheet_in_workbook(o_wb, ws) Then
                    Set o_ws = o_wb.Sheets(ws.name)
                    o_no_of_rows = row_count(o_ws)
                    If o_no_of_rows = 1 Then
                        ws.UsedRange.Copy o_ws.Range("A1")
                    Else
                        no_of_rows = row_count(ws)
                        no_of_columns = column_count(ws, ws.UsedRange.row) ' To DO Remove the row Parameter
                        last_column = column_index_to_name(no_of_columns)
                        ws.Range("A2:" & last_column & no_of_rows).Copy o_ws.Range("A" & CStr(o_no_of_rows + 1))
                    End If
                Else
                    ws.Copy Before:=o_wb.Sheets(1)
                End If
            Next
            wb.Close
        End If
    Next file
    
    o_wb.SaveAs o_file, AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    o_wb.Close (True)
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    
End Sub
Function sheet_in_workbook(wb As Workbook, sh As Worksheet) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.name = sh.name Then
            sheet_in_workbook = True
            Exit Function
        End If
    Next
    sheet_in_workbook = False
End Function

Function item_in_collection(c As Collection, item As String)
    Dim value
    For Each value In c
        If value = item Then
            item_in_collection = True
            Exit Function
        End If
    Next
    item_in_collection = False
End Function
Function delete_sheets_in_workbook(wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ws.Delete
    Next
End Function
Function delete_file(file As String)
    If Len(Dir(file)) > 0 Then
        Call Kill(file)
    End If
End Function
Function row_count(ws As Worksheet, Optional column As String = "A") As Long
    With ws
        row_count = .Range(column & .Rows.Count).End(xlUp).row
    End With
End Function
Function column_count(ws As Worksheet, Optional row As String = 1) As Long
    With ws
        column_count = .Cells(row, .Columns.Count).End(xlToLeft).column
    End With
End Function
Function column_name_to_index(name As String)
     With ActiveSheet
        column_name_to_index = .Range(name & 1).column
    End With
End Function
Function column_index_to_name(index As Long)
     With ActiveSheet
        column_index_to_name = Split(.Cells(1, index).Address(True, False), "$")(0)
    End With
End Function
