Sub stocks()

' Loop through all worksheets

Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
WS.Activate

' Declare new variables

Dim open_price As Double
Dim close_price As Double
Dim ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim volume As Double
volume = 0
Dim row As Integer
row = 2

' Add headers for main variables in summary table

Cells(1, 9).Value = "ticker"
Cells(1, 10).Value = "yearly_change"
Cells(1, 11).Value = "percent_change"
Cells(1, 12).Value = "volume"

' Find the last row, aka non-blank cell in column A(1)

LstRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

' Set the opening price to be the beginning year price, otherwise the loop uses the open price from the end of the year

open_price =  Cells(row, 3).Value

'Start loop through ticker symbols

	For i = 2 To LstRow

		' Check if we are still within same ticker and assign ticker symbol

		If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

		' Set the ticker symbol

		ticker = Cells(i, 1).Value
		Cells(row, 9).Value = ticker

		' Set the closing price

		close_price = Cells(i, 6).Value

		' Calculate the yearly change (closing price minus opening price)

		yearly_change = close_price - open_price
		Cells(row, 10) = yearly_change

		' Calculate the percent change (yearly change / opening price)

       ***IMPORTANT NOTE: This must be written as an if/else statement with three conditions in order to account for cases that have zero as the denominator, which causes an error. Thus, cases where the opening price is zero are assigned an automatic value so that the division within the percent change formula does not occur for these cases. All other cases that have a true denominator follow the percent change forumla.***

		If (open_price = 0 And close_price = 0) Then
                    percent_change = 0
                ElseIf (open_price = 0 And close_price <> 0) Then
                    percent_change = 1
                Else
                    percent_change = yearly_change / open_price
                    Cells(row, 11).Value = percent_change
                    Cells(row, 11).NumberFormat = "0.00%"

		End if

		' Calculate the volume (summing volume for all rows within the same ticker)

		volume = volume + Cells(i, 7).Value
		cells(row, 12).Value = volume

		' Set the next open price for new ticker symbol

		open_price = Cells(i + 1, 3).Value

		' Add new row to summary table'
		row = row + 1

		' Set the volume back to zero if there is a new ticker symbol

		volume = 0

		' Othwerise, if it's still the same ticker, keep summing volume values

            Else
            	Volume = Volume + Cells(i, 7).Value
           
        End If

 	Next i

 		' Start next looop to assign cell color: green for positive yearly chance and red for negative change

     	' First determine the last row for yearly change'
     	 
     	 LstRow_yearly_change = Cells(Rows.Count, 10).End(xlUp).Row
         
         For j = 2 To LstRow_yearly_change
            
            If Cells(j, 10).Value > 0 Then
            	Cells(j, 10).Interior.ColorIndex = 4
            Else
            	Cells(j, 10).Interior.ColorIndex = 3

            End If

    Next j

    ' Challenges: return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". 

    ' Headers for new variables

     Cells(2, 14).Value = "greatest % increase"
     Cells(3, 14).Value = "greatest % decrease"
     Cells(4, 14).Value = "greatest volume"
     Cells(1, 15).Value = "ticker"
     Cells(1, 16).Value = "value"

     For k = 2 To LstRow_yearly_change

       ' Caluclate the greatest % increase

        If Cells(k, 11).Value =Application.WorksheetFunction.Max(WS.Range("K2:K" & LstRow_yearly_change)) Then
        	
        	' Assign the ticker symbol
        	Cells(2, 15).Value = Cells(k, 9).Value
        	' Assign value
         	Cells(2, 16).Value = Cells(k, 11).Value
         	Cells(2, 16).NumberFormat = "0.00%"
        
        ' Calculate the greatest % decrease

        ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & LstRow_yearly_change)) Then
         	' Assign the ticker symbol
         	Cells(3, 15).Value = Cells(k, 9).Value
         	' Assign value
         	Cells(3, 16).Value = Cells(k, 11).Value
        	Cells(3, 16).NumberFormat = "0.00%"
         
         ' Calculate greatest total volume

         ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & LstRow_yearly_change)) Then
         	' Assign the ticker symbol
         	Cells(4, 15).Value = Cells(k, 9).Value
         	' Assign value
         	Cells(4, 16).Value = Cells(k, 12).Value
       
        End If

    Next k

Next WS

End Sub
