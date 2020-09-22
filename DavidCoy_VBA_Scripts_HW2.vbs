Sub idklol()

    Dim Current As Worksheet
    Dim Ticker_Name As String
    Dim Ticker_Total As Double
    Dim Summary_table As Integer
    Dim Ticker_Count As Integer
    Dim Ticker_Open As Double
    Dim Ticker_Close As Double
    Dim Ticker_Rows As Integer
    Dim Ticker_Volume As Double
    
    For Each Current In Worksheets
        Current.Cells(1, 9) = "Ticker"
        Current.Cells(1, 10) = "Yearly Change"
        Current.Cells(1, 11) = "Percent Change"
        Current.Cells(1, 12) = "Total Stock Volume"

        Ticker_Count = 0 + 2
        Summary_Table_Row = 2
        Row = 2
        Ticker_Rows = 0
        Ticker_Volume = 0
    
        Do Until IsEmpty(Current.Cells(Row, 1)) ' Do this because we do not know the last row with data for each sheet
            Ticker_Volume = Current.Cells(Row, 7).Value + Ticker_Volume
            If Current.Cells(Row + 1, 1).Value <> Current.Cells(Row, 1).Value Then
                Ticker_Name = Current.Cells(Row, 1).Value 'Get the Ticker's Name
                Current.Cells(Ticker_Count, 9) = Ticker_Name 'Insert the Ticker's Name
                Ticker_Close = Current.Cells(Row, 6).Value 'Get closing value
                Ticker_Open = Current.Cells(Row - Ticker_Rows, 3).Value 'Get opening value
                Current.Cells(Ticker_Count, 10) = Ticker_Close - Ticker_Open ' Yearly Change
                If Current.Cells(Ticker_Count, 10).Value <= 0 Then
                    Current.Cells(Ticker_Count, 10).Interior.ColorIndex = 3
                Else
                    Current.Cells(Ticker_Count, 10).Interior.ColorIndex = 4
                End If
                Current.Cells(Ticker_Count, 12) = Ticker_Volume
                
                If Ticker_Open = 0 Or Ticker_Close = 0 Then ' For some reason "PLNT" is all zeroes
                    Current.Cells(Ticker_Count, 11) = 0 'Percent Change
                Else
                    Current.Cells(Ticker_Count, 11) = 1 - Ticker_Close / Ticker_Open 'Percent Change
                End If
                Current.Cells(Ticker_Count, 11).Style = "Percent" 'Make the style a rounded percentage
                Current.Cells(Ticker_Count, 11).NumberFormat = "0.00%"
                Ticker_Rows = -1
                Ticker_Count = Ticker_Count + 1
                Ticker_Volume = 0
            End If
        Ticker_Rows = Ticker_Rows + 1
        Row = Row + 1
        Loop
    'MsgBox Current.Name
    Summary Current
    Next
End Sub

Sub Summary(Current)
    Current.Cells(2, 15) = "Greatest % Increase"
    Current.Cells(3, 15) = "Greatest % Decrease"
    Current.Cells(4, 15) = "Greatest Total Volume"
    Current.Cells(1, 16) = "Ticker"
    Current.Cells(1, 17) = "Value"
    
    Dim Greatest_Inc As Double
    Dim Greatest_Dec As Double
    Dim Greatest_Total_Vol As Double
    Dim Greatest_Inc_Ticker As String
    Dim Greatest_Dec_Ticker As String
    Dim Greatest_Total_Vol_Ticker As String
    
    Greatest_Inc = 0
    Greatest_Dec = 0
    Greatest_Total_Vol = 0
    
    Row = 2
    Do Until IsEmpty(Current.Cells(Row, 11))
        If Current.Cells(Row, 11).Value < Greatest_Inc Then
            Greatest_Inc = Current.Cells(Row, 11).Value
            Greatest_Inc_Ticker = Current.Cells(Row, 9).Value
        End If
        
        If Current.Cells(Row, 11).Value > Greatest_Dec Then
            Greatest_Dec = Current.Cells(Row, 11).Value
            Greatest_Dec_Ticker = Current.Cells(Row, 9).Value
        End If
        
        If Current.Cells(Row, 12).Value > Greatest_Total_Vol Then
            Greatest_Total_Vol = Current.Cells(Row, 12).Value
        End If
        
        Row = Row + 1
    Loop
    
    'Greatest_Inc = Application.WorksheetFunction.Max(Range("K:K"))
    Current.Cells(2, 16).Value = Greatest_Inc_Ticker
    Current.Cells(2, 17).Value = Greatest_Inc * -1
    Current.Cells(2, 17).NumberFormat = "0.00%"
    Current.Cells(2, 17).Style = "Percent" 'Make the style a rounded percentage
    
    'Greatest_Inc = Application.WorksheetFunction.Min(Range("K:K"))
    Current.Cells(3, 16).Value = Greatest_Dec_Ticker
    Current.Cells(3, 17).Value = Greatest_Dec * -1
    Current.Cells(3, 17).NumberFormat = "0.00%"
    Current.Cells(3, 17).Style = "Percent" 'Make the style a rounded percentage
    
    Current.Cells(4, 17).Value = Greatest_Total_Vol
End Sub

