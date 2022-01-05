Attribute VB_Name = "Module1"

'Create a script that will loop through all the stocks for one year and output the following information:
        'The ticker symbol
        'yearly change from opening price at the beginning of a given year to the closing price at the end of the year
        'The percent change from opening price at the beginning of a given yar to the closing price at the end of that year
        'The stock volume of the stock
        
Sub Stock_Analysis():
    
    'Declaring the variables
    Dim i As Long
    Dim x As Integer
    Dim LastRow1 As Long
    Dim Ticker As String
    Dim Stock_Ticker As Integer
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Total_Volume As Double
    Dim Percent_Change As Range
    Dim Min_Value As Double
    Dim Max_Value As Double
    Dim ws As Worksheet
    Dim Ticker_e As String
    Dim Ticker_c As String
    Dim Ticker_f As String
    

    For Each ws In Worksheets
        'Creating the name for each column in my stock analysis table
        ws.Cells(1, 9).Value = "Stock_Ticker"
        ws.Cells(1, 10).Value = "Yearly_Change"
        ws.Cells(1, 11).Value = "Percentage_Change"
        ws.Cells(1, 12).Value = "Total_Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'Setting the initial row count to start on second row
            Stock_Ticker = 2
            Ticker = ws.Cells(2, 1).Value
            ws.Cells(Stock_Ticker, 9).Value = Ticker
            Counter = 0
            
        'Setting the opening price to start at C2
            Opening_Price = ws.Cells(2, 3).Value
            
        'Defining my last row
            LastRow1 = ws.Cells(Rows.Count, "A").End(xlUp).Row
                
                'Creating a loop through each row for all the stocks for one year
                For i = 2 To 800000
                
                        'Calculating the Total Volume for each Ticker in the loop
                        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
             
                                'Checking to make sure that ticker is different from the previous
                                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                                
                                        'Return new Ticker in the next cell
                                        'Stock_Ticker = ws.Cells(i, 1).Value
                                        ws.Cells(Stock_Ticker, 9).Value = ws.Cells(i, 1)
                    
                                        'Return the value of the closing price which is different from opening price
                                        Closing_Price = ws.Cells(i, 6).Value
        
                                        'Calculating the Yearly Change
                                        Yearly_Change = Closing_Price - Opening_Price
                                        
                                        'Storing the Yearly Change
                                        ws.Cells(Stock_Ticker, 10).Value = Yearly_Change
                                        
                                        'Calculating the Percentage Change
                                        'Need to avoid division by 0 when opening price is 0
                                                    If Opening_Price <> 0 Then
                                                            'Percentage_Change = 0
                                                            'Yearly_Change = 0
                                                    
                                                            Percentage_Change = Yearly_Change / Opening_Price
                                                            
                                                    End If

                                        'Storing the Percentage Change
                                        ws.Cells(Stock_Ticker, 11).Value = Percentage_Change
        
                                        'Storing the Total Volume for each Ticker in the loop
                                        ws.Cells(Stock_Ticker, 12).Value = Total_Volume
                                        
                                        'Enabling a newer Ticker to be picked if different from the previous one
                                            Stock_Ticker = Stock_Ticker + 1
                                        
                                             Opening_Price = Cells(i + 1, 3).Value
           
                                            Total_Volume = 0
                                            
                                         Dim y As Integer
                                         y = 0
                                         If i <= LastRow Then
                                            Do While ws.Cells(i + y, 1).Value = Ticker
                                                If ws.Cells(i + y, 3).Value <> 0 Then
                                                    Opening_Price = ws.Cells(i + y, 3).Value
                                                    Exit Do
                                                Else
                                                    Opening_Price = 1
                                                    y = y + 1
                                                End If
                                            Loop
                                End If
                            End If
                            
                Next i

                         'Defining the lastrow for column J to go through and format color
                                LastRow2 = ws.Cells(Rows.Count, "J").End(xlUp).Row
                         'Creating a loop for color formatting the Yearly Change
                         For x = 2 To LastRow2
            
                                        'Creating the conditional formatting for the Yearly Change
                                        If ws.Cells(x, 10).Value > 0 Then
            
                                                'Setting the first conditional formating to green for any positive change
                                                ws.Cells(x, 10).Interior.Color = RGB(0, 255, 0)
                                                
                                        ElseIf ws.Cells(x, 10).Value < 0 Then
                                               'Setting the second conditional formating to red for any negative change
                                                ws.Cells(x, 10).Interior.Color = RGB(255, 0, 0)
                                               
                                        End If
                                          ' Fomatting column K to percentage
                                            ws.Cells(x, 11).NumberFormat = "0.00%"
                        Next x
                        
                        
                      
                        
                       'Defining the lastrow for column K to go through to find the maximum and minimum yearly change
                                LastRow3 = ws.Cells(Rows.Count, "K").End(xlUp).Row
                                Max_Value = ws.Cells(2, 11).Value
                                Min_Value = ws.Cells(2, 11).Value

                                  For k = 3 To LastRow3
                                        If ws.Cells(k, 11) > Max_Value Then
                                               'ws.Cells(2, 16).Value = MaxValue
                                                Max_Value = ws.Cells(k, 11).Value
                                                Ticker_e = ws.Cells(k, 9).Value
                                               'Fomatting cell to percentage
                                                ws.Cells(2, 16).NumberFormat = "0.00%"
                                                  
                                        ElseIf ws.Cells(k, 11) < Min_Value Then
                                                'ws.Cells(3, 16).Value = MinValue
                                                Min_Value = ws.Cells(k, 11).Value
                                                Ticker_c = ws.Cells(k, 9).Value
                                               'Fomatting cell to percentage
                                                ws.Cells(3, 16).NumberFormat = "0.00%"
                                                
                                        End If
                                        
                                Next k
                                
                                'Populating the values of their respective cells
                                ws.Cells(2, 16).Value = Max_Value
                                ws.Cells(2, 15).Value = Ticker_e
                                ws.Cells(3, 16).Value = Min_Value
                                ws.Cells(3, 15).Value = Ticker_c
                                
                                'Finding the greatest total volume
                                LastRow4 = ws.Cells(Rows.Count, "L").End(xlUp).Row
                                Greatest_Volume = ws.Cells(2, 12).Value
                                Ticker_f = ws.Cells(2, 9).Value
                                
                                For l = 3 To LastRow4
                                    If ws.Cells(l, 12) > Greatest_Volume Then
                                    Greatest_Volume = ws.Cells(l, 12).Value
                                     Ticker_f = ws.Cells(l, 9).Value
                                     
                                    End If

                                Next l
                                
                                'Populating the values of their respective cells
                                ws.Cells(4, 16).Value = Greatest_Volume
                                ws.Cells(4, 15).Value = Ticker_f
                                
                                'Fomatting cell to number
                                ws.Cells(4, 16).NumberFormat = "0"
    Next ws
    
End Sub



