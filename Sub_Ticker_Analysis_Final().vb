Sub Ticker_Analysis_Final()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
 
        ' Set an initial variable to store the last row in the main dataset
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set an initial vairable for the MaxClose and MinOpen values
        Dim MaxClose As Double
        Dim MinOpen As Double
        
        ' Set initial variables for Total Volume
        Dim Total_Vol As Double
        Total_Vol = 0
        
        ' Set initial variables for Yearly Changes and Ticker Symbol
        Dim YearlyChange As Double
        Dim YearlyChangePercent As Double
        Dim Ticker As String
        
        ' Set an initial variable for the Summary_Table row location
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' Loop through all ticker symbols
        For i = 2 To LastRow
        
            ' Check if cells above and below each other match
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Store Ticker Symbol and MinOpen Values
                Ticker = ws.Cells(i, 1).Value
                MinOpen = ws.Cells(i, 3).Value
            End If
            
            ' Update MaxClose for each record
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
               MaxClose = ws.Cells(i, 6).Value
            End If
                    
            'Calculate the Total_vol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
            Total_Vol = Total_Vol + ws.Cells(i, 7).Value
                    
            Else
        
                Total_Vol = Total_Vol + ws.Cells(i, 7).Value
        
            End If
                            
            ' Check if the next row has a different Ticker Symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Calculate Yearly Changes
                YearlyChange = MaxClose - MinOpen
                YearlyChangePercent = (MaxClose - MinOpen) / MinOpen
                
                ' Output values to the Summary Table
                
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                ws.Range("K" & Summary_Table_Row).Value = YearlyChangePercent
                ws.Range("L" & Summary_Table_Row).Value = Total_Vol
                                       
                            
                ' Move to the next row in the Summary Table
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset Total Volume MinOpen and MaxClose for the next Ticker
                Total_Vol = 0
                MinOpen = 0
                MaxClose = 0
            End If
            
            
        Next i
        
        ' Set initial varaibles for bottom of summary table
        Dim j As Integer
        Dim LastRow2 As Long
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Set initial variables for the outliers table
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Volume As Double
        Dim Ticker_Greatest_Increase As String
        Dim Ticker_Greatest_Decrease As String
        Dim Ticker_Greatest_Volume As String
        Dim Ticker2 As String
        Dim PercentRange As Range
        Set PercentRange = ws.Range("K2:K" & LastRow2)
        Dim VolumeRange As Range
        Set VolumeRange = ws.Range("L2:L" & LastRow2)
        Dim Greatest_Increase_Row As Long
        Dim Greatest_Decrease_Row As Long
        Dim Greatest_Volume_Row As Long
        
        
        'Loop through the summary table and apply conditional formatting on Yearly Change
        For j = 2 To LastRow2
            
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            Else: ws.Cells(j, 10).Interior.ColorIndex = 3
        
            End If
            
            'Format Percent Change column as Percentage
            ws.Cells(j, 11).NumberFormat = "0.00%"
            
        Next j
        
        'Assign headers for summary table and outliers table
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        ws.Range("P1:Q1").Value = Array("Ticker", "Value")
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Assign calculations for outliers table
        Greatest_Increase = ws.Application.WorksheetFunction.Max(PercentRange)
        Greatest_Decrease = ws.Application.WorksheetFunction.Min(PercentRange)
        Greatest_Volume = ws.Application.WorksheetFunction.Max(VolumeRange)
        
        'Fill in outliers table with variables and apply formatting
        ws.Range("Q2").Value = Greatest_Increase
        ws.Range("Q3").Value = Greatest_Decrease
        ws.Range("Q4").Value = Greatest_Volume
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'Loop through cells to determine ticker symbol for outliers table
        For j = 2 To LastRow2
            
            If ws.Cells(j, 11) = Greatest_Increase Then
                ws.Range("P2") = ws.Cells(j, 9)
                            
            End If
            
            If ws.Cells(j, 11) = Greatest_Decrease Then
                ws.Range("P3") = ws.Cells(j, 9)
                
            End If
            
            If ws.Cells(j, 12) = Greatest_Volume Then
                ws.Range("P4") = ws.Cells(j, 9)
                
            End If
            
        Next j
    
        'Autofit columns on sheets
        ActiveSheet.Columns.AutoFit
    
    Next ws
        
End Sub