Sub VOlume():

'Make formula loop through each worksheet in the workbook
Dim xsheet As Worksheet
For Each xsheet In ThisWorkbook.Worksheets
xsheet.Select

   'Declare variables For Ticket Symbol
    Dim Ticker As String
    
    'Declare variables for Stock Volume
    Dim Stock_Volume As Double
    
    'Declare Stock_Volume starting value
    Stock_Volume = 0
    
    'Keep Track of each ticker in the summary table
    Dim Stock_Table_Row As Integer
    Stock_Table_Row = 2
    
    'Declare Titles for Summary Table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"

    'Loop through all the ticket symbols
    For I = 2 To 75000
    
        'Check to see if we are still within same Ticker
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
            'Set Ticker Name
            Ticker = Cells(I, 1).Value
            
            'Add to the Ticker Total
            Stock_Volume = Stock_Volume + Cells(I, 7).Value
            
            'Print Ticker in the Summary Table
            Range("I" & Stock_Table_Row).Value = Ticker
            
            'Print Stock Volume in the Summary Table
            Range("J" & Stock_Table_Row).Value = Stock_Volume
            
            'Add one to the stock_Table_Row
            Stock_Table_Row = Stock_Table_Row + 1
            
            'Reset Stock_Volume
            Stock_Volume = 0
        
        'If cell is the same
        Else
        
            'Add to the Ticker
            Stock_Volume = Stock_Volume + Cells(I, 7).Value
            
        End If
        
    Next I

 
 Next xsheet
 
    
End Sub

