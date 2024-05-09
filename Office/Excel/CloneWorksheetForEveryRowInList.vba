' insert module execute and then remove module or save without macro

Sub CopyListAndRename()

  ' Define the source table name and list names
  Dim tblName As String, listName As String
  tblName = "Vloge" ' worksheet name where to iterate through first column
  listName = "OL1" ' worksheet to clone for every row

  ' Get the last row of data in the table
  Dim lastRow As Long
  lastRow = Sheets(tblName).Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through each row in the table (skip the header) - from row 3 because 1 = header, 2 = first row for which worksheet already exists
  For Row = 3 To lastRow
    ' Get the value from the current row
    Dim newName As String
    newName = Sheets(tblName).Cells(Row, 1).Value

    ' Check if value is not empty
    If newName <> "" Then
      ' Copy the list
      Sheets(listName).Copy After:=Worksheets(Sheets.Count)

      ' Get the newly created sheet
      Dim newSheet As Worksheet
      Set newSheet = ActiveSheet

      ' Rename the sheet with the value
      newSheet.Name = "OL" & newName

      ' Clear the data in the newly created sheet (optional)
      ' newSheet.Cells.ClearContents
    End If
  Next Row

End Sub
