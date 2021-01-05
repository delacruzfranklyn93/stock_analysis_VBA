Attribute VB_Name = "Module4"
Sub final()

    Dim Current As Worksheet
    Application.ScreenUpdating = False
    
    For Each Current In Worksheets
        Current.Select
        Call test
    Next
    Application.ScreenUpdating = True
End Sub

Sub test()

'Initalize the variables

    Dim count As Double
    Dim column As Integer
    Dim ticker As Integer
    Dim total_row As Integer
    Dim op As Double
    Dim cl As Double

'Assign the variables and first ticker cell

    count = 0
    Range("I2").Value = Cells(2, 1).Value
    column = 1
    Row = 2
    total_row = 2
    op = Range("C2").Value
    cl = 0
    
'Assign the proper column names

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % increase"
    Range("N3").Value = "Greatest % decrease"
    Range("N4").Value = "Greatest Total Volume"
    

    
'Begin for loop to iterate over the rows
    
    For i = 2 To Cells(Rows.count, 1).End(xlUp).Row
        
        count = Cells(i, 7).Value + count

'Check to see if there is a new ticker classifiaction
        
        If Cells(i, column).Value <> Cells(i + 1, column).Value Then

'Set the first new ticker under the ticker column and each new ticker thereafter in the row directly beneath
            
            Cells(Row + 1, 9).Value = Cells(i + 1, column).Value

'Calculate the yearly change using the opening and closing price
            
            cl = Cells(i, 6).Value
            Cells(Row, 10).Value = Format(cl - op, ".00")

'Calculate the percent change using the opening and closing price
            
            If op = 0 Then
                Cells(Row, 11).Value = Format(0, "Percent")
            Else
                Cells(Row, 11).Value = Format((cl - op) / op, "Percent")
            End If

'Create the conditional formating to change to red or green depending if yearly change is positive or negative
            
            If Cells(Row, 10).Value >= 0 Then
                Cells(Row, 10).Interior.ColorIndex = 4
            Else
                Cells(Row, 10).Interior.ColorIndex = 3
            End If
            
'Update op to the opening price of the new ticker

            op = Cells(i + 1, 3).Value


'Since the next row is a new ticker go ahead and set the Total of the current row under column Total Stock Volume
            
            Cells(total_row, 12).Value = count

'Add one to total_row and row so the next row is where the new ticker data will be stored
            
            total_row = total_row + 1
            Row = Row + 1

'Set the count back to 0 so the new summation will only reflect the new ticker symbol

            count = 0
        End If
    Next i
    Call great
End Sub

Sub great()

    For i = 2 To Cells(Rows.count, "I").End(xlUp).Row
    
'Check to see if the percent change cells are equal to 0 if not start finding the greatest increase/decrease
        
        If Cells(i, 11).Value <> 0 Then

'Check to see if current cell is greater than the previous greatest % increase and update great_i
            
            If Cells(i, 11).Value > great_i Then
                great_i = Cells(i, 11).Value
                Range("O2").Value = Cells(i, 9).Value
                Range("P2").Value = Format(great_i, "Percent")
                
'Check to see if the current cell is less than the greatest previous decrease and update great_d
            
            ElseIf Cells(i, 11).Value < great_d Then
                great_d = Cells(i, 11).Value
                Range("O3").Value = Cells(i, 9).Value
                Range("P3").Value = Format(great_d, "Percent")
            
            End If:
        End If:

'Check to see if the total stock volume = 0 if not start searching for the greatest total stock volume

        If Cells(i, 12).Value <> 0 Then
        
'Check to see if current cell is greater than the previous total stock volume and update great_t

            If Cells(i, 12).Value > great_t Then
                great_t = Cells(i, 12).Value
                Range("O4").Value = Cells(i, 9).Value
                Range("P4").Value = great_t
            End If:
        End If:
    Next i:
End Sub
