
Sub loop_through_all_worksheets()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    'do whatever you need
    ws.Cells(1, 1) = 1 'this sets cell A1 of each sheet to "1"

    Call stocks
Next

starting_ws.Activate 'activate the worksheet that was originally active


End Sub




