Attribute VB_Name = "Module1"
Sub Ticker_Data()

For Each ws In Worksheets

' Set initial variable for Ticker Symbol, Yearly Change, Percent Change, and Total Stock Volume
Dim Ticker_Symbol As String
Dim Yearly_Change As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Percent_Change As Double
Dim Total_Volume As LongLong

' Set variables to 0
Total_Volume = 0
Open_Price = "0.00"
Close_Price = "0.00"

' Keep track of the location for each Ticker Symbol in the output table
Dim Output_Table_Row As Integer
Output_Table_Row = 2

' Determine the last row of ticker data
LastRow_Ticker = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Label the new columns
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"
ws.Cells(1, 12) = "Total Stock Volume"
ws.Cells(2, 14) = "Greatest % Increase"
ws.Cells(3, 14) = "Greatest % Decrease"
ws.Cells(4, 14) = "Greatest Total Volume"
ws.Cells(1, 15) = "Ticker"
ws.Cells(1, 16) = "Value"

    ' Loop through all data in the column
    For i = 2 To LastRow_Ticker
    
         ' Identify Starting Value for Percent Change
         If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1) Then
            Open_Price = ws.Cells(i, 3).Value

         ' Identify when the ticker symbol changes
         ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
             ' Output the Ticker Symbol
            Ticker_Symbol = ws.Cells(i, 1).Value
        
             ' Print the Ticker Symbol in the Output Table
            ws.Range("I" & Output_Table_Row).Value = Ticker_Symbol
        
            ' Output Yearly Change
            Close_Price = ws.Cells(i, 6).Value
            Yearly_Change = Close_Price - Open_Price
       
             ' Print Yearly Change in the Output Table
             ws.Range("J" & Output_Table_Row).Value = Yearly_Change
             
                'Conditional formatting color change for Yearly Change
                If Yearly_Change = 0 Then
                    ws.Range("J" & Output_Table_Row).Interior.ColorIndex = 0
                ElseIf Yearly_Change > 0 Then
                    ws.Range("J" & Output_Table_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Output_Table_Row).Interior.ColorIndex = 3
                End If
         
             ' Output Percent Change
                    If Yearly_Change <> "0" And Open_Price <> "0" Then
                        Percent_Change = (Yearly_Change / Open_Price)
                        ' Print Percent Change in the Output Table
                        ws.Range("K" & Output_Table_Row).Value = Percent_Change
                        ws.Range("K" & Output_Table_Row).NumberFormat = "0.00%"
                    ElseIf Yearly_Change <> "0" And Open_Price = "0" Then
                        Percent_Change = Yearly_Change
                        ' Print Percent Change in the Output Table
                        ws.Range("K" & Output_Table_Row).Value = Percent_Change
                        ws.Range("K" & Output_Table_Row).NumberFormat = "0.00%"
                    Else
                        ws.Range("K" & Output_Table_Row).Value = "0.00%"
                    End If
       
            ' Output the Total Stock Volume
            Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
            ' Print the Total Stock Volume
             ws.Range("L" & Output_Table_Row).Value = Total_Volume
           
            ' Add one to the summary table row
             Output_Table_Row = Output_Table_Row + 1
        
             ' Reset variables
            Total_Volume = 0
            Open_Price = "0.00"
            Close_Price = "0.00"
        
        ' If the following cell has the same Ticker Symbol
        Else
    
         ' Add to the Total Stock Volume
         Total_Volume = Total_Volume + ws.Cells(i, 7).Value
        
        End If
    Next i

' Determine the greatest % increase and decrease
Dim per_rng As Range
Dim Percent_Max As Double
Dim Percent_Min As Double

' Determine the last row of ticker data
LastRow_Percent = ws.Cells(Rows.Count, 11).End(xlUp).Row

' Look through the Percent Change column for the Max
Set per_rng = ws.Range("K:K")
Percent_Max = Application.WorksheetFunction.Max(per_rng)
ws.Range("P2") = Percent_Max

' Determine Ticker Symbol for Greatest % Max
For i = 2 To LastRow_Percent
    If ws.Cells(i, 11) = Percent_Max Then
        ws.Range("O2") = ws.Cells(i, 9)
        Exit For
    End If
Next i
    
' Look through the Percent Change column for the Min
Set per_rng = ws.Range("K:K")
Percent_Min = Application.WorksheetFunction.Min(per_rng)
ws.Range("P3") = Percent_Min

' Determine Ticker Symbol for Greatest % Min
For i = 2 To LastRow_Percent
    If ws.Cells(i, 11) = Percent_Min Then
        ws.Range("O3") = ws.Cells(i, 9)
        Exit For
    End If
Next i
    
' Determine the Greatest Total Volume
Dim vol_rng As Range
Dim Greatest_Volume As LongLong

Set vol_rng = ws.Range("L:L")
Greatest_Volume = Application.WorksheetFunction.Max(vol_rng)
ws.Range("P4") = Greatest_Volume

' Determine the last row of ticker data
LastRow_Volume = ws.Cells(Rows.Count, 12).End(xlUp).Row

' Determine Ticker Symbol for Total Volume
For i = 2 To LastRow_Volume
    If ws.Cells(i, 12) = Greatest_Volume Then
        ws.Range("O4") = ws.Cells(i, 9)
        Exit For
    End If
Next i

' Format values to percent
ws.Range("P2:P3").NumberFormat = "0.00%"

' Autofit data
ws.Columns("I:P").AutoFit

Next ws
End Sub
