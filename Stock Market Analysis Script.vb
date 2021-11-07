Sub StockMarket_Analysis()

'Define the variables that you are working with
'--------------------------------------------------

'Variable for Ticker

Dim Ticker As String

'Variable for year open

Dim year_open As Double

'Variable for year close

Dim year_close As Double

'Variable for yearly change

Dim Yearly_Change As Double

'Variable for total stock volume

Dim Total_Stock_Vol As Double

'Variable for percent change

Dim Percent_Change As Double

'Variable to allow a row to start

Dim data_start As Integer

'Variable to excute the code in all work sheet at once within workbook

Dim wks As Worksheet

'Loop for worksheet to excute the code once
'--------------------------------------------------

For Each wks In Worksheets

    'Assign a column header for every task we are going perform
    
    wks.Range("I1").Value = "Ticker"
    wks.Range("J1").Value = "Yearly Change"
    wks.Range("K1").Value = "Percent Change"
    wks.Range("L1").Value = "Total Stock Volume"

    'Assign intiger for the loop to start
    data_start = 2
    previous_i = 1
    Total_Stock_Vol = 0
    
    EndRow = wks.Cells(Rows.Count, "A").End(xlUp).Row

        'Summarize and loop the yearly change, percent change, and total stock volume
        
        For i = 2 To EndRow
            
            'If Ticker alphabet changes or is not equal to the previous
            
            If wks.Cells(i + 1, 1).Value <> wks.Cells(i, 1).Value Then
            
            'Get the Ticker alphabet
            
            Ticker = wks.Cells(i, 1).Value
            
            'Intiate the variable to move to next ticker
            
            previous_i = previous_i + 1
            
            ' Get the value first day open and last day close form the column 3 or "C" and on column 6 or "F" repectively
            
            year_open = wks.Cells(previous_i, 3).Value
            year_close = wks.Cells(i, 6).Value
            
            ' A for loop to sum the total stock volume
            
            For j = previous_i To i
            
                Total_Stock_Vol = Total_Stock_Vol + wks.Cells(j, 7).Value
                
            Next j
            
            'Open the data when the value reaches zero
            
            If year_open = 0 Then
            
                Percent_Change = year_close
                
            Else
                Yearly_Change = year_close - year_open
                
                Percent_Change = Yearly_Change / year_open
                
            End If
         '--------------------------------------------------
         
            'Get the worksheet summary table
            
            wks.Cells(data_start, 9).Value = Ticker
            wks.Cells(data_start, 10).Value = Yearly_Change
            wks.Cells(data_start, 11).Value = Percent_Change
            
            'Formatting values to percentage
        
            wks.Cells(data_start, 11).NumberFormat = "0.00%"
            wks.Cells(data_start, 12).Value = Total_Stock_Vol
            
            'go to the next row when summary is completed
            
            data_start = data_start + 1
            
            'Return the variable to zero
            
            Total_Stock_Vol = 0
            Yearly_Change = 0
            Percent_Change = 0
            
            'Change i number to variable previous_i
            previous_i = i
        
        End If
    
    Next i
    
'The bonus summery table
  '--------------------------------------------------
    
    'Go to the last row of column k
    
    kEndRow = wks.Cells(Rows.Count, "K").End(xlUp).Row
    
    'Define variable to initiate the bonus summery table values

    Increase = 0
    Decrease = 0
    Greatest = 0
    
        'find max/min for percentage change and the max volume Loop
        For k = 3 To kEndRow
        
            'Define previous increment to check
            last_k = k - 1
                        
            'Define current row for percentage
            current_k = wks.Cells(k, 11).Value
            
            'Define Previous row for percentage
            prevous_k = wks.Cells(last_k, 11).Value
            
            'greatest total volume row
            volume = wks.Cells(k, 12).Value
            
            'Prevous greatest volume row
            prevous_vol = wks.Cells(last_k, 12).Value
            
   '--------------------------------------------------
            
            'Find the increase by defining increase as increase
            If Increase > current_k And Increase > prevous_k Then
                
                Increase = Increase
                
            ElseIf current_k > Increase And current_k > prevous_k Then
                
                Increase = current_k
                
                'define name for increase percentage in current
                increase_name = wks.Cells(k, 9).Value
                
            ElseIf prevous_k > Increase And prevous_k > current_k Then
            
                Increase = prevous_k
                
                'define name for increase percentage in previous
                increase_name = wks.Cells(last_k, 9).Value
                
            End If
                
       '--------------------------------------------------
            'Find the decrease by defining the decrease as decrease
            
            If Decrease < current_k And Decrease < prevous_k Then
                
                Decrease = Decrease
    
            ElseIf current_k < Increase And current_k < prevous_k Then
                
                Decrease = current_k
                
              
                decrease_name = wks.Cells(k, 9).Value
                
            ElseIf prevous_k < Increase And prevous_k < current_k Then
            
                Decrease = prevous_k

                decrease_name = wks.Cells(last_k, 9).Value
                
            End If
            
       '--------------------------------------------------
           'Find the greatest volume
           
            If Greatest > volume And Greatest > prevous_vol Then
            
                Greatest = Greatest
            
            ElseIf volume > Greatest And volume > prevous_vol Then
            
                Greatest = volume
                
                'define name for greatest volume
                greatest_name = wks.Cells(k, 9).Value
                
            ElseIf prevous_vol > Greatest And prevous_vol > volume Then
                
                Greatest = prevous_vol
                
                'define name for greatest volume
                greatest_name = wks.Cells(last_k, 9).Value
                
            End If
            
        Next k
  '--------------------------------------------------
    ' Assigning names
    
    wks.Range("N1").Value = "Column Name"
    wks.Range("N2").Value = "Greatest % Increase"
    wks.Range("N3").Value = "Greatest % Decrease"
    wks.Range("N4").Value = "Greatest Total Volume"
    wks.Range("O1").Value = "Ticker Name"
    wks.Range("P1").Value = "Value"
    
    'Get values for greatest increase, greatest increase, and  greatest volume Ticker name
    wks.Range("O2").Value = increase_name
    wks.Range("O3").Value = decrease_name
    wks.Range("O4").Value = greatest_name
    wks.Range("P2").Value = Increase
    wks.Range("P3").Value = Decrease
    wks.Range("P4").Value = Greatest
    
    'Formatting of the greatest increase and decrease in percentage
    
    wks.Range("P2").NumberFormat = "0.00%"
    wks.Range("P3").NumberFormat = "0.00%"


'--------------------------------------------------
' Conditional formatting columns of the colors

'Ending the row for column J

    jEndRow = wks.Cells(Rows.Count, "J").End(xlUp).Row
    

        For j = 2 To jEndRow
            
            'if values are greater than or less than zero
            If wks.Cells(j, 10) > 0 Then
            
                wks.Cells(j, 10).Interior.ColorIndex = 4
                
            Else
            
                wks.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
        Next j
    
'Excute code to the next worksheet
Next wks
'--------------------------------------------------
End Sub
