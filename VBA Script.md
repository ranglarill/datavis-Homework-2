Sub WallStreetData():


   'Define variables
      
       Dim Ticker As String
       Dim Volume As Double
       Dim InitialVolume As Double
        Dim LastRowAdj As Double
        Dim SummaryTableRow As Integer
        Dim LastRow As Double
        Dim OpenPrice As Double
        Dim PercentChange As Double
        Dim Change As Double

      

        
'Dim i As Long

        SummaryTableRow = 2
       InitialVolume = 0
       Volume = InitialVolume


    'Set the headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total_Stock_Volume"
 


    'Define Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    LastRowAdj = LastRow + 1


    'Loop through year of data for each ticker
       For i = 3 To LastRowAdj

        
            'See if ticker in current row matches ticker in previous row

            If Cells(i, 1).Value = Cells(i - 1, 1).Value Then

            ' Add current row volume to stored volume Summary Table and move to the next

                Volume = Volume + Cells(i, 7).Value

            'If ticker doesn't match
            Else

            Ticker = Cells(i - 1, 1).Value
            Cells(SummaryTableRow, 9).Value = Ticker

          'Put stored total for prior ticket in Summary table
            Cells(SummaryTableRow, 10).Value = Volume

            'Reset stored volume to current row value
            Volume = Cells(i, 7).Value

            'Reset summary table row
            SummaryTableRow = SummaryTableRow + 1



            End If

           

    Next i



End Sub




