Function CreateCombinedStocksWorksheet() As Object
    
    'ƒ() intent: to create the combined stocks worksheet with the "Ticker" and "Total Stock Volume" column headers
    
    Application.DisplayAlerts = False
    
    'if the "Combined_Stocks" worksheet already exists, delete it
    
    For Each s In Worksheets
    
        If (s.Name = "Combined_Stocks") Then
        
            Worksheets("Combined_Stocks").Delete
        
            Exit For
            
        End If
        
    Next s
    
    Application.DisplayAlerts = True

    Dim CombineStocksWorksheet As Object
    Dim ColumnNames(2) As String
    
    ColumnNames(0) = "Year"
    ColumnNames(1) = "Ticker"
    ColumnNames(2) = "Total Stock Volume"
    
    Worksheets.Add Before:=Worksheets(1), Count:=1
    
    Set CombineStocksWorksheet = Worksheets(1)
    CombineStocksWorksheet.Name = "Combined_Stocks"
    
    CombineStocksWorksheet.Activate
    
    Cells(1, 1).Value = ColumnNames(0)
    Cells(1, 2).Value = ColumnNames(1)
    Cells(1, 3).Value = ColumnNames(2)
    
    Set CreateCombinedStocksWorksheet = CombineStocksWorksheet
    
End Function
Function GetLastRow() As Range

    'ƒ() intent: find the last row (based on column A) that contains data and return the range object
    
            Dim LastRow As Range
            
            Set LastRow = Cells(Rows.Count, 1).End(xlUp)
            Set GetLastRow = LastRow

End Function

Sub CombineStocksPerSheet()
    
    Dim Ticker, Year As String
    Dim ActiveSheet_LastRowIndex, CombinedStocks_LastRowIndex As Long
    Dim Total As Double
    
    Set ws_CombinedStocks = CreateCombinedStocksWorksheet()
    Total = 0
    
    For Each s In Worksheets
    
        Year = s.Name
        s.Activate
    
        ActiveSheet_LastRowIndex = GetLastRow().Row
        
        Ticker = Cells(2, 1).Value
        
            For i = 2 To ActiveSheet_LastRowIndex + 1
        
                'code to execute when the ticker matches the value
        
                If (Ticker = Cells(i, 1).Value) Then
                
                    'changing the data type of the cell to decimal
                    Cells(i, 7).NumberFormat = "0.00"
                    
                    Total = Total + Cells(i, 7).Value
                
                ElseIf (i = ActiveSheet_LastRowIndex) Then
                
                    ws_CombinedStocks.Activate
                    CombinedStocks_LastRowIndex = GetLastRow().Row
                    'Cells((CombinedStocks_LastRowIndex), 2).AutoFit
                    
                Else
                    
                    ws_CombinedStocks.Activate
                    CombinedStocks_LastRowIndex = GetLastRow().Row
                    Cells((CombinedStocks_LastRowIndex + 1), 1).Activate
                    
                    Cells((CombinedStocks_LastRowIndex + 1), 1).Value = Year
                    Cells((CombinedStocks_LastRowIndex + 1), 2).Value = Ticker
                    Cells((CombinedStocks_LastRowIndex + 1), 3).NumberFormat = "0.00"
                    Cells((CombinedStocks_LastRowIndex + 1), 3).Value = Total
                    
                    s.Activate
                    Total = 0
                    Ticker = Cells(i, 1).Value
        
                End If
        Next i
        
        Year = ""
    
    Next s
    
    ws_CombinedStocks.Activate
    ws_CombinedStocks.Columns("A:C").AutoFit
End Sub
