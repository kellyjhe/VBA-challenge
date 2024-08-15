Module 2 Challenge

Screenshots of multiple_year_stock_data:
<img width="1100" alt="Q1 results" src="https://github.com/user-attachments/assets/79d85316-2c8f-4186-a502-b9d2c6f557fa">
<img width="1143" alt="Q2 results" src="https://github.com/user-attachments/assets/ba332410-f478-425d-91e6-6daf212495f3">

Screenshots of alphabetical_testing:
<img width="1138" alt="Screenshot 2024-08-15 at 4 27 30 PM" src="https://github.com/user-attachments/assets/8474c02c-6f8d-462d-a8e2-daca7703d311">


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
