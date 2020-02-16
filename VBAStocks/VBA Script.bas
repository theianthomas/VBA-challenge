Attribute VB_Name = "Module1"
Sub StocksWs()

' loop through all sheets

For Each Ws In Worksheets

' set headers in each worksheet

    Ws.Range("I1").Value = "Ticker"
    Ws.Range("J1").Value = "Yearly Change"
    Ws.Range("K1").Value = "Percent Change"
    Ws.Range("L1").Value = "Total Stock Volume"
    Ws.Range("O2").Value = "Greatest % Increase"
    Ws.Range("O3").Value = "Greatest % Decrease"
    Ws.Range("O4").Value = "Greatest Total Volume"
    Ws.Range("P1").Value = "Ticker"
    Ws.Range("Q1").Value = "Value"

' set initial variable for holding the ticker name
    Dim Ticker_Name As String
    Ticker_Name = " "
        
' set an initial variable for holding the total per ticker name
    Dim Total_Ticker_Volume As Double
    Total_Ticker_Volume = 0
        
' set variables for all price calculations
    Dim Open_Price As Double
    Open_Price = 0
    Dim Close_Price As Double
    Close_Price = 0
    Dim Change_Price As Double
    Change_Price = 0
    Dim Change_Percent As Double
    Change_Percent = 0
    Dim MAX_TICKER_NAME As String
    MAX_TICKER_NAME = " "
    Dim MIN_TICKER_NAME As String
    MIN_TICKER_NAME = " "
    Dim MAX_PERCENT As Double
    MAX_PERCENT = 0
    Dim MIN_PERCENT As Double
    MIN_PERCENT = 0
    Dim MAX_VOLUME_TICKER As String
    MAX_VOLUME_TICKER = " "
    Dim MAX_VOLUME As Double
    MAX_VOLUME = 0


' keep track of location for each ticker name
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
' set initial row count
    Dim Lastrow As Long
    Dim i As Long

Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
    
' set initial value of Open Price
    Open_Price = Ws.Cells(2, 3).Value
    
For i = 2 To Lastrow

' set initial value of Close Price
    Close_Price = Ws.Cells(i, 6).Value
    
' loop through ticker
    
If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            
' set the ticker name
    Ticker_Name = Ws.Cells(i, 1).Value
    

' calculate change in price and percent
    Change_Price = Close_Price - Open_Price
' condition if there is a 0
    If Open_Price <> 0 Then
        Change_Percent = (Change_Price / Open_Price) * 100
    End If
    
' add to the Ticker name total volume
    Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value
    
' print in the summary table
    Ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    Ws.Range("J" & Summary_Table_Row).Value = Change_Price
    Ws.Range("K" & Summary_Table_Row).Value = (CStr(Change_Percent) & "%")
    Ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
    
' Fill change in price with Green and Red colors
    If (Change_Price > 0) Then
' Fill column with GREEN color
    Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    ElseIf (Change_Price <= 0) Then
' Fill column with RED color
    Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
    
' Add 1 to the summary table row count
    Summary_Table_Row = Summary_Table_Row + 1
' Reset change in price and close price
    Change_Price = 0
    Close_Price = 0
' next open price per ticker
    Open_Price = Ws.Cells(i + 1, 3).Value
    
' percent change conditional
                If (Change_Percent > MAX_PERCENT) Then
                    MAX_PERCENT = Change_Percent
                    MAX_TICKER_NAME = Ticker_Name
                ElseIf (Change_Percent < MIN_PERCENT) Then
                    MIN_PERCENT = Change_Percent
                    MIN_TICKER_NAME = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > MAX_VOLUME) Then
                    MAX_VOLUME = Total_Ticker_Volume
                    MAX_VOLUME_TICKER = Ticker_Name
                End If
                ' reset counters
                Total_Ticker_Volume = 0
                
Else

' increase Total Ticker Volume
    Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value
    

End If
        
Next i

Ws.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
Ws.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
Ws.Range("P2").Value = MAX_TICKER_NAME
Ws.Range("P3").Value = MIN_TICKER_NAME
Ws.Range("Q4").Value = MAX_VOLUME
Ws.Range("P4").Value = MAX_VOLUME_TICKER

Next Ws

End Sub

