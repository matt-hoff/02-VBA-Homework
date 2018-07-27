Sub Volume():

'Make formula loop through each worksheet in the workbook
Dim xsheet As Worksheet
For Each xsheet In ThisWorkbook.Worksheets
xsheet.Select

   'Declare variables for ticker symbol
    Dim Ticker As String
    
    'Declare variables for stock volume
    Dim Stock_Volume As Double
    
    'Declare stock volume starting value
    Stock_Volume = 0
    
    'Keep track of each ticker in the summary table
    Dim Stock_Table_Row As Integer
    Stock_Table_Row = 2
    
    'Declare titles for summary table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"

    'Loop through all the ticket symbols
    For I = 2 To 800000
    
        'Check to see if we are still within same ticker
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
            'Set ticker name
            Ticker = Cells(I, 1).Value
            
            'Add to the ticker total
            Stock_Volume = Stock_Volume + Cells(I, 7).Value
            
            'Print ticker in the summary table
            Range("I" & Stock_Table_Row).Value = Ticker
            
            'Print stock volume in the summary table
            Range("J" & Stock_Table_Row).Value = Stock_Volume
            
            'Add one to the stock table row
            Stock_Table_Row = Stock_Table_Row + 1
            
            'Reset stock volume
            Stock_Volume = 0
        
        'If cell is the same
        Else
        
            'Add to the ticker
            Stock_Volume = Stock_Volume + Cells(I, 7).Value
            
        End If
        
    Next I
 
 Next xsheet
    
End Sub