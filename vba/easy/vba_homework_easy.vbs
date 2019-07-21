Function CreateCombinedStocksWorksheet() As Object
    
    'ƒ() intent: to create the combined stocks worksheet with the "Ticker" and "Total Stock Volume" column headers
    
    Dim CombineStocksWorksheet As Object
    Dim ColumnNames(1) As String
    
    ColumnNames(0) = "Ticker"
    ColumnNames(1) = "Total Stock Volume"
    
    Worksheets.Add Before:=Worksheets(1), Count:=1
    
    Set CombineStocksWorksheet = Worksheets(1)
    CombineStocksWorksheet.Name = "Combined_Stocks"
    
    CombineStocksWorksheet.Activate
    
    Cells(1, 1).Value = ColumnNames(0)
    Cells(1, 2).Value = ColumnNames(1)
    
    Set CreateCombinedStocksWorksheet = CombineStocksWorksheet

End Function

Function GetLastColumn() As Range

    'ƒ() intent: find the last column (based on row 1) that contains data and return the range object

            Dim LastColumn As Range
            
            Set LastColumn = Cells(1, Columns.Count).End(xlToLeft)
            Set GetLastColumn = LastColumn

End Function

Function GetLastRow() As Range

    'ƒ() intent: find the last row (based on column A) that contains data and return the range object
    
            Dim LastRow As Range
            
            Set LastRow = Cells(Rows.Count, 1).End(xlUp)
            Set GetLastRow = LastRow

End Function
Function GetStockTicker(LastColumnIndex As Long) As String

        'ƒ() intent: find the column labeled "<ticker>", then return the stock ticker value
        
        Dim ActiveCell As Range
        Dim StockTicker As String

        For i = 1 To LastColumnIndex
        
            Set ActiveCell = Cells(1, i)
            
            'activating the current range object to access the value it contains
            
            ActiveCell.Activate
        
            If (ActiveCell.Value = "<ticker>") Then

                StockTicker = ActiveCell.Offset(1, 0).Value
                
                'exiting the for loop to prevent unnecessary iterations
                
                Exit For

            End If
            
        Next i
            
        GetStockTicker = StockTicker
            
End Function

Function GetVolumeTotal(LastColumnIndex As Long, LastRowIndex As Long) As Double

        'ƒ() intent: find the column labeled "<vol>", sum all the values in the column, then return the result
        
        Dim ActiveCell As Range
        Dim VolumeTotal As Double
        
        VolumeTotal = 0
        
        For i = 1 To LastColumnIndex

            Set ActiveCell = Cells(1, i)

            'activating the current range object to access the value it contains

            ActiveCell.Activate

            If (ActiveCell.Value = "<vol>") Then
            
                Dim FirstCellAddress, LastCellAddress As String
                Dim VolumeRange As Range
                
                'getting the cell address for the first and last cells to define a dynamic range
                
                FirstCellAddress = ActiveCell.Offset(1, 0).Address
                LastCellAddress = Cells(LastRowIndex, LastColumnIndex).Address
                
                Set VolumeRange = Range(FirstCellAddress & ":" & LastCellAddress)
                
                'changing the data type of the spreadsheets cells to decimal so I can use the sum function on the dynamic range
                
                VolumeRange.NumberFormat = "0.00"
                
                VolumeTotal = Application.Sum(VolumeRange)
                
                'starting the iteration at 2 because the first row is the column header
                
                
'                For j = 2 To LastRowIndex
'
'                    'converting the string to a double
'
'                    VolumeTotal = VolumeTotal + CDbl(Cells(j, LastColumnIndex).Value)
'
'                Next j

                'exiting the for loop to prevent unnecessary iterations

                Exit For

            End If

        Next i
        
        GetVolumeTotal = VolumeTotal
            
End Function

Sub CombineStockData()

    'ƒ() intent: to produce the results of the homework assignment

    'this counter will be used to place the ticker and volume in the correct row
    
    Dim Counter As Integer
    Dim WSCombinedStocks As Object
    Set WSCombinedStocks = CreateCombinedStocksWorksheet()
                
    Counter = 2

    For Each s In Worksheets
    
        Dim SheetName As String
        SheetName = s.Name
        
        If Not (SheetName = "Combined_Stocks") Then
        
            Dim ActiveCell, LastColumn, LastRow As Range
            Dim LastColumnIndex, VolumeColumnIndex As Long
            Dim StockTicker As String
            Dim VolumeTotal As Double
            
            'activate the worksheet that is currently being referenced in this iteration of the loop
                    
            s.Activate
        
            Set LastColumn = GetLastColumn()
            Set LastRow = GetLastRow()
              
            'get the stock ticker and volume total for the active worksheet
            
            StockTicker = GetStockTicker(LastColumn.Column)
            VolumeTotal = GetVolumeTotal(LastColumn.Column, LastRow.Row)
            
            'activating the Combined_Stocks worksheet to make the "Cells" object point to it
            
            WSCombinedStocks.Activate
            
            'outputting the stock ticker and volume total
                        
            Cells(Counter, 1).Value = StockTicker
            Cells(Counter, 2).Value = VolumeTotal
            
            Counter = (Counter + 1)
        End If
    
    Next s
    
    'after outputting all the stock data:
    'I am using autofit to ensure the column length is wide enough to display all the contents inside the cell
    
    WSCombinedStocks.Activate
    Range("A:B").Columns.AutoFit
    MsgBox ("Operation Complete")
    
End Sub
