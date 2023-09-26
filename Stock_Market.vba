Sub Stock_Market_A()

'choose worksheet
Dim Current_Sheet As Worksheet

'Check for summary table headers
Create_Summary_Headers = True

'loop through all worksheets
For Each Current_Sheet In Worksheets

    'set variables in each worksheet
    Dim Summary_Table_Row As Long
    Dim Last_Row As Long
    Dim Ticker_Symbol As String
    Dim Yearly_Change As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim Bonus_EndRow As Long
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double



    'Find last row in worksheet
    Last_Row = Current_Sheet.Cells(Rows.Count, 1).End(xlUp).Row

    'Create summary headers
    If Create_Summary_Headers Then
        Current_Sheet.Range("I1").Value = "Ticker"
        Current_Sheet.Range("J1").Value = "Yearly Change"
        Current_Sheet.Range("K1").Value = "Percent Change"
        Current_Sheet.Range("L1").Value = "Total Volume"

        Current_Sheet.Range("N2").Value = "Greatest % Increase"
        Current_Sheet.Range("N3").Value = "Greatest % Decrease"
        Current_Sheet.Range("N4").Value = "Greatest Total Volume"
        Current_Sheet.Range("O1").Value = "Ticker"
        Current_Sheet.Range("P1").Value = "Value"
        
    Else
        Create_Summary_Headers = True
    End If

    'For loop starting points
    Summary_Table_Row = 2
    Total_Volume = 0
    
    'prev i checks first day opening
    prev_i = 1


    'sort stocks by ticker symbol
    For i = 2 To Last_Row

        'check for ticket symbol change
        If Current_Sheet.Cells(i + 1, 1).Value <> Current_Sheet.Cells(i, 1).Value Then
            
            'set ticker value & increase iteration
            Ticker_Symbol = Current_Sheet.Cells(i, 1).Value
            
            prev_i = prev_i + 1
            
            'get first day opening price and last day closing price
            Opening_Price = Current_Sheet.Cells(prev_i, 3).Value
            Closing_Price = Current_Sheet.Cells(i, 6).Value
            
            'calculate yearly change and percent change
            If (Opening_Price <> 0) Then
                Yearly_Change = Closing_Price - Opening_Price
                Percent_Change = Yearly_Change / Opening_Price
            Else
                Percent_Change = Closing_Price
            
            End If
            
            'condition format yearly changed whether greater than zero or less than equal to 0
            If (Yearly_Change > 0) Then
                Current_Sheet.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf (Yearly_Change <= 0) Then
                Current_Sheet.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If

            'sum total stock volume
            For j = prev_i To i
            
                Total_Volume = Total_Volume + Current_Sheet.Cells(j, 7).Value
            
            Next j
            
            'set values to summary table & format percentages
            Current_Sheet.Range("J" & Summary_Table_Row).NumberFormat = "0.00%"
            
            Current_Sheet.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
            Current_Sheet.Range("J" & Summary_Table_Row).Value = Yearly_Change
            Current_Sheet.Range("K" & Summary_Table_Row).Value = Percent_Change
            Current_Sheet.Range("L" & Summary_Table_Row).Value = Total_Volume
            
            Summary_Table_Row = Summary_Table_Row + 1
        
            'Reset variables
            prev_i = i
            Yearly_Change = 0
            Percent_Change = 0
            Total_Volume = 0
        
        
        End If

    Next i

    'set initial variables to 0
    Greatest_Increase = WorksheetFunction.Max(Current_Sheet.Range("K:K"))
    Greatest_Decrease = WorksheetFunction.Min(Current_Sheet.Range("K:K"))
    Greatest_Volume = WorksheetFunction.Max(Current_Sheet.Range("L:L"))
    
    EndRow_2 = Current_Sheet.Cells(Rows.Count, "K").End(xlUp).Row
    
    'Find max and min for percentage change
    For Summary_Table_2 = 2 To EndRow_2
        
        If (Current_Sheet.Cells(Summary_Table_2, 11) = Greatest_Increase) Then
            Increase_Val = Current_Sheet.Cells(Summary_Table_2, 9).Value
        End If
        
        If (Current_Sheet.Cells(Summary_Table_2, 11) = Greatest_Decrease) Then
            Decrease_Val = Current_Sheet.Cells(Summary_Table_2, 9).Value
        End If
        
        If (Current_Sheet.Cells(Summary_Table_2, 12) = Greatest_Volume) Then
            Volume_Val = Current_Sheet.Cells(Summary_Table_2, 9).Value
        End If
        
    Next Summary_Table_2
    
    Current_Sheet.Range("P2").NumberFormat = "0.00%"
    Current_Sheet.Range("P3").NumberFormat = "0.00%"

    Current_Sheet.Range("O2").Value = Increase_Val
    Current_Sheet.Range("O3").Value = Decrease_Val
    Current_Sheet.Range("O4").Value = Volume_Val
    Current_Sheet.Range("P2").Value = Greatest_Increase
    Current_Sheet.Range("P3").Value = Greatest_Decrease
    Current_Sheet.Range("P4").Value = Greatest_Volume
    
Next Current_Sheet

End Sub