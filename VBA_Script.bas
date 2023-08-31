Attribute VB_Name = "Module1"
Sub StockAnalysis()

    'define the variables that will be used to store the values
    'long for integer, string for text, double for noninteger
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    'loop thru each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        'recommended by tutor to add in
        ws.Activate
        'clear previously analyzed data before each script run
        Columns("I:Q").Delete
        
        'find the last row with data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'set initial value for variables
        ticker = ws.Cells(2, "A").Value
        openPrice = ws.Cells(2, "C").Value
        totalVolume = 0
        
        'suggested by tutor to create the new variable calc purpose
        summaryPointer = 2
        
        'loop thru rows starting from the second row, where the data starts
        For I = 2 To lastRow
        
            'if ticker on next row is different, then data for current ticker is complete
            If ws.Cells(I + 1, 1).Value <> ticker Then
            
                'pull closing price for the current stock for analysis
                closePrice = ws.Cells(I, "F").Value
            
                'calculates the change in stock price
                yearlyChange = closePrice - openPrice
            
                    'if openPrice is not equal to zero then calculate percentChange
                    If openPrice <> 0 Then
                        percentChange = (yearlyChange / openPrice) * 100
                    Else
                        percentChange = 0
                    End If
            
                'calculate the total volume of the current stock
                totalVolume = totalVolume + ws.Cells(I, "G").Value
            
                'enter value from analysis into these cells
                ws.Cells(summaryPointer, "I").Value = ticker
                ws.Cells(summaryPointer, "J").Value = yearlyChange
                ws.Cells(summaryPointer, "K").Value = "%" & percentChange
                ws.Cells(summaryPointer, "L").Value = totalVolume
            
                'conditioning format yearly change to green for positive and red for negative
                If yearlyChange > 0 Then
                    ws.Cells(summaryPointer, "J").Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryPointer, "J").Interior.Color = RGB(255, 0, 0)
                End If
                
                'if the ticker symbol in the next row is different, then complete and move to next ticker
                ticker = ws.Cells(I + 1, "A").Value
                openPrice = ws.Cells(I + 1, "C").Value
                'restart the new ticker
                totalVolume = 0
                
                summaryPointer = summaryPointer + 1
                
            Else
            
                'calculate total volume for the same ticker
                totalVolume = totalVolume + ws.Cells(I, "G").Value
            
            End If
        
        Next I
        
        'find how many row in the new summary date
        summary_count = Cells(Rows.Count, "I").End(xlUp).Row
        
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVolume As Double
        Dim increaseTicker As String
        Dim decreaseTicker As String
        Dim volumeTicker As String
        
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        
        'find the bonus summary with sample code from tutor
        For Row = 2 To summary_count
            If Cells(Row, "K").Value > maxIncrease Then
                maxIncrease = Cells(Row, "K").Value
                increaseTicker = Cells(Row, "I").Value
            End If
            
        Next Row
        
        Range("O2") = increaseTicker
        Range("P2") = maxIncrease & "%"
        
        For Row = 2 To summary_count
            If Cells(Row, "K").Value < maxDecrease Then
                maxDecrease = Cells(Row, "K").Value
                decreaseTicker = Cells(Row, "I").Value
            End If
            
        Next Row
        
        Range("O3") = decreaseTicker
        Range("P3") = maxDecrease & "%"
        
        For Row = 2 To summary_count
            If Cells(Row, "L").Value > maxVolume Then
                maxVolume = Cells(Row, "L").Value
                volumeTicker = Cells(Row, "I").Value
            End If
            
        Next Row
        
        Range("O4") = volumeTicker
        Range("P4") = maxVolume
        
        'headers for analyzed tables
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "yearlyChange"
        ws.Cells(1, "K").Value = "percentChange"
        ws.Cells(1, "L").Value = "totalVolume"
        ws.Cells(2, "N").Value = "Greatest % Increase"
        ws.Cells(3, "N").Value = "Greatest % Decrease"
        ws.Cells(4, "N").Value = "Greatest Total Volume"
        ws.Cells(1, "O").Value = "Ticker"
        ws.Cells(1, "P").Value = "Value"
        
        'autofit col
        Columns("A:Q").AutoFit
        
    Next ws
    
    MsgBox ("Analysis Done")
    
End Sub
