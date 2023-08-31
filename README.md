# VBA-challenge

## Multi-Year Stock Data Analysis

This script analyzes stock data across multiple worksheets. It will calculate the yearly change, percent change, and total volume for each stock. Additionally, it will identify stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume from the analyzed data.

## Process

### Step 1.
Define the variables used to store data and metrics for analysis using dim.

### Step 2.
Clear previously analyzed data before looping through each worksheet in the workbook.
```
For Each ws In ThisWorkbook.Worksheets
 ws.Activate
 Columns("I:Q").Delete
 ...
Next was
```

### Step 3:
Loop through each row of data to calculate data and populate the summary table.
```
For I = 2 To lastRow
…
Next I
```

### Step 4:
Add conditioning format to the yearly change column 
```
If yearlyChange > 0 Then
ws.Cells(summaryPointer, "J").Interior.Color = RGB(0, 255, 0)
ElseIf yearlyChange < 0 Then
ws.Cells(summaryPointer, "J").Interior.Color = RGB(255, 0, 0)
End If
```

### Step 5:
Use the generated summary table to find the greatest percentage increase, greatest percentage decrease, and greatest total volume.
```
For Row = 2 To summary_count
If Cells(Row, "COL").Value > VALUE Then
VALUE = Cells(Row, "COL").Value
VALUE = Cells(Row, "COL").Value
End If
```
 
### Step 6:
Create tables for the analyzed data
```
ws.Cells(1, "COL").Value = "TITLE"
```

### Step 7:
For formatting and ease of use purposes: 1) percentage data reformatted, 2) columns will auto fit to display all data, 3) display a message showing the analysis is complete. 
```
Range("P2") = maxIncrease & "%"
Columns("A:Q").AutoFit
MsgBox ("Analysis Done")
```

#### Note:
The final code was simplified from redundant lines and reformatted for ease of use based on suggestions during the tutoring session.
