Sub cleanBlanksSelection()


' set a variable for range selection
Dim WorkRng As Range




   On Error GoTo ErrorHandler
   ' select range, if nothing is selected, then we go to ErrorHandler
   ' create input box
    xTitleId = "Select Cells To Clean"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)

    ' final range select statement
   WorkRng.Select

    With Selection

    ' copy cells to themselves using the NumberFormat method

    Selection.NumberFormat = "General"
    .Value = .Value

    End With

    ' turns on screen updating - turn off if the system is slow or crashing

    Application.ScreenUpdating = True

    Exit Sub




ErrorHandler:

   ' Tell the user there is an error with the range
   MsgBox "Either you didn't select some cells, or the sheet may have another error."

   Resume Next


End Sub
