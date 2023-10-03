# Module-2
Sub cmdStart()
  Dim sCompany As String
  Dim iWorksheet As Integer
  Dim lRow As Long
  Dim lTotalRow As Long
  Dim oSheet As Variant
  Dim lOpen As Double
  Dim lClose As Double
  Dim dblTotalVol As Double
  Dim sPrevCompany As String
  Dim iGreatestIncrease As Double
  Dim iGreatestDecrease As Double
  Dim lGreatestTotalVolume As Double
  Dim sGreatestIncresaeTicker As String
  Dim sGreatestDecreaseTicker As String
  Dim sGreatestTotalVolumeTicker As String

  Const kiCompany As Integer = 1
  Const kiOpen As Integer = 3
  Const kiClose As Integer = 6
  Const kiVolumen As Integer = 7
  Const kiTicker As Integer = 9
  Const kiYearChange As Integer = 10
  Const kiPercentChange As Integer = 11
  Const kiTotalStockVolume As Integer = 12
  Const kiGreatestTicker As Integer = 16
  Const kiGreatestValue As Integer = 17
  Const kiGreatestLabelCol As Integer = 15
  
     
  For Each oSheet In ThisWorkbook.Sheets
    'Paint Summary titles
    lTotalRow = 1
    iGreatestIncrease = 0
    iGreatestDecrease = 0
    lGreatestTotalVolume = 0

    oSheet.Cells(lTotalRow, kiTicker).Value = "Ticker"
    oSheet.Cells(lTotalRow, kiYearChange) = "Year Change"
    oSheet.Cells(lTotalRow, kiPercentChange) = "Percent Change"
    oSheet.Cells(lTotalRow, kiTotalStockVolume) = "Total Stock Volume"
    
    lRow = 2
    lTotalRow = 2
    dblTotalVol = 0
    sPrevCompany = oSheet.Cells(lRow, kiCompany)
      
    Do While oSheet.Cells(lRow, kiCompany) <> ""
        
        sCompany = oSheet.Cells(lRow, kiCompany)
        If lRow = 2 Then
            lOpen = oSheet.Cells(lRow, kiOpen)
        End If
                
        If sCompany <> sPrevCompany Then 'change of ticker. its time to total ticker
            'Paint Totals.
            oSheet.Cells(lTotalRow, kiTicker) = sPrevCompany
            oSheet.Cells(lTotalRow, kiYearChange) = lClose - lOpen
            If oSheet.Cells(lTotalRow, kiYearChange) < 0 Then
                'Red
                oSheet.Cells(lTotalRow, kiYearChange).Interior.ColorIndex = 3
            Else
                'Green
                oSheet.Cells(lTotalRow, kiYearChange).Interior.ColorIndex = 10
            End If
            oSheet.Cells(lTotalRow, kiPercentChange) = (((lClose - lOpen) * 100) / lOpen) / 100
            
            oSheet.Cells(lTotalRow, kiPercentChange).NumberFormat = "0.00%"
            
            oSheet.Cells(lTotalRow, kiTotalStockVolume) = dblTotalVol
            
            
            If lTotalRow = 2 Then 'first company totals
                iGreatestIncrease = oSheet.Cells(lTotalRow, kiPercentChange)
                iGreatestDecrease = oSheet.Cells(lTotalRow, kiPercentChange)
                lGreatestTotalVolume = oSheet.Cells(lTotalRow, kiTotalStockVolume)
                'unique ticker for first row for now
                sGreatestIncresaeTicker = sPrevCompany
                sGreatestDecreaseTicker = sPrevCompany
                sGreatestTotalVolumeTicker = sPrevCompany
                
            Else
          
                If iGreatestIncrease < oSheet.Cells(lTotalRow, kiPercentChange) Then
                    iGreatestIncrease = oSheet.Cells(lTotalRow, kiPercentChange)
                    sGreatestIncresaeTicker = oSheet.Cells(lTotalRow, kiTicker)
                End If
                
                If iGreatestDecrease > oSheet.Cells(lTotalRow, kiPercentChange) Then
                    iGreatestDecrease = oSheet.Cells(lTotalRow, kiPercentChange)
                    sGreatestDecreaseTicker = oSheet.Cells(lTotalRow, kiTicker)
                End If
                
                If lGreatestTotalVolume < oSheet.Cells(lTotalRow, kiTotalStockVolume) Then
                    lGreatestTotalVolume = oSheet.Cells(lTotalRow, kiTotalStockVolume)
                    sGreatestTotalVolumeTicker = oSheet.Cells(lTotalRow, kiTicker)
                End If
            
            End If
            
            dblTotalVol = 0
            lTotalRow = lTotalRow + 1
            lOpen = oSheet.Cells(lRow, kiOpen)
        
        End If

        dblTotalVol = dblTotalVol + oSheet.Cells(lRow, kiVolumen)
        
        lClose = oSheet.Cells(lRow, kiClose)
        lRow = lRow + 1
        sPrevCompany = sCompany
    Loop
    
    'print greatest titles summary
    oSheet.Cells(1, kiGreatestTicker) = "Ticker"
    oSheet.Cells(1, kiGreatestValue) = "Value"
    oSheet.Cells(2, kiGreatestLabelCol) = "Greatest % Increase"
    oSheet.Cells(3, kiGreatestLabelCol) = "Greatest % Decrease"
    oSheet.Cells(4, kiGreatestLabelCol) = "Greatest Total Volume"
    
    oSheet.Cells(2, kiGreatestTicker) = sGreatestIncresaeTicker
    oSheet.Cells(3, kiGreatestTicker) = sGreatestDecreaseTicker
    oSheet.Cells(4, kiGreatestTicker) = sGreatestTotalVolumeTicker
    
    oSheet.Cells(2, kiGreatestValue) = iGreatestIncrease
    oSheet.Cells(2, kiGreatestValue).NumberFormat = "0.00%"
    
    oSheet.Cells(3, kiGreatestValue) = iGreatestDecrease
    oSheet.Cells(3, kiGreatestValue).NumberFormat = "0.00%"
    
    oSheet.Cells(4, kiGreatestValue) = lGreatestTotalVolume
    
  
  Next
 

End Sub
