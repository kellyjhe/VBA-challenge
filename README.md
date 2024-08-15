Module 2 Challenge

Code sources:

Used help from Xpert Learning code:
For percentage formatting
  Sub FormatAsPercentage()
      ' Define the range you want to format as a percentage
      Dim rng As Range
      Set rng = Range("A1")  ' Change "A1" to your desired cell or range

      ' Apply the percentage format to the range
      rng.NumberFormat = "0.00%"  ' This will format the number with two decimal places and a percentage sign

      ' Optionally, you can also set the value as a percentage directly
      rng.Value = 0.75  ' This will display 75.00% in the cell

  End Sub


To find max
  Sub FindMaxNumber()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxNumber As Double

    ' Set the worksheet where your data is located
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Find the last row in the column (assuming it's column A)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Find the maximum number in the column
    maxNumber = Application.WorksheetFunction.Max(ws.Range("A1:A" & lastRow))

    ' Display the maximum number found
    MsgBox "The maximum number in column A is: " & maxNumber
End Sub
