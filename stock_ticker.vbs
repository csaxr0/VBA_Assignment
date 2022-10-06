Sub Stock_Ticker()

'This code will calculate stock ticker Variables for this assignment

'Set Variables
Dim Row As Integer
Dim lastRow As Long
Dim Percent_Change As Double
Dim yearly_change As Double
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Ticker_Table_Row As Double
Dim Ticker_Name As String
Dim Stock_Volume As Double



Range("I1").Value = "Ticker Name"
Range("L1").Value = "Total"
Range("J1").Value = "yearly change"
Range("k1").Value = "Percent Change"


'Setup for loop

            'Ticker_Table_Row = 2
            lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
'Opening Price
            Opening_Price = Cells(2, 3).Value
                      
            Ticker_Table_Row = 2
            
'Loop and calculations for Ticker, Closing Price, Yearly Change, Percent Change & Volume
    For Row = 2 To lastRow
        If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
            Closing_Price = Cells(Row, 6).Value
            Stock_Volume = Stock_Volume + Cells(Row, 7).Value
            Ticker_Name = Cells(Row, 1).Value
            yearly_change = (Closing_Price - Opening_Price)
            Percent_Change = (yearly_change / Opening_Price)
            Range("I" & Ticker_Table_Row).Value = Ticker_Name
            Range("L" & Ticker_Table_Row).Value = Stock_Volume
            Range("J" & Ticker_Table_Row).Value = yearly_change
            If Range("J" & Ticker_Table_Row).Value >= 0 Then
                Range("J" & Ticker_Table_Row).Interior.ColorIndex = 4 'Green
            ElseIf Range("J" & Ticker_Table_Row).Value < 0 Then
                Range("J" & Ticker_Table_Row).Interior.ColorIndex = 3 'Red
            End If
            Range("k" & Ticker_Table_Row).Value = FormatPercent(Percent_Change)
            Ticker_Table_Row = Ticker_Table_Row + 1
            Stock_Volume = 0
            Opening_Price = Cells(Row, 3).Value
        Else
            Stock_Volume = Stock_Volume + Cells(Row, 7).Value
        End If

    Next Row
     
     
End Sub



