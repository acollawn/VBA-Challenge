Attribute VB_Name = "Module1"
Sub vba_analysis():

    ' PART 1
    ' set all of the dimensions
    Dim i As Long
    Dim j As Integer
    Dim total As Double
    Dim start As Long
    Dim lastrow As Long
    Dim yearly_change As Double
    Dim percent_change As Double
    
    ' create output rows for ticker, yearly change, percent change, total stock volume
    Range("J1").Value = "ticker"
    Range("K1").Value = "yearly change"
    Range("L1").Value = "percent change"
    Range("M1").Value = "total stock volume"
    
    ' create output rows for ticker,, value, greatest % increase, greatest % decrease, greatest total stock volume
    Range("P2").Value = "greatest % increase"
    Range("P3").Value = "greatest % decrease"
    Range("P4").Value = "greatest total volume"
    Range("Q1").Value = "ticker"
    Range("R1").Value = "value"
    ' establish my initial values
    j = 0
    total = 0
    yearly_change = 0
    start = 2
    
    ' find number of rows
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
  
    For i = 2 To lastrow

        ' If the ticker symbol changes
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' then record new total for ticker's vol
            total = total + Cells(i, 7).Value

            ' If total volume equals zero
            If total = 0 Then
                ' print the results
                Range("J" & 2 + j).Value = Cells(i, 1).Value
                Range("K" & 2 + j).Value = 0
                Range("L" & 2 + j).Value = "%" & 0
                Range("M" & 2 + j).Value = 0

            Else
                ' Find First non zero starting value (dont forget start=2)
                If Cells(start, 3) = 0 Then
                    For nonZero = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = nonZero
                            Exit For
                        End If
                     Next nonZero
                End If

                ' Calculate the change values
                yearly_change = (Cells(i, 6) - Cells(start, 3))
                percent_change = yearly_change / Cells(start, 3)

                ' start of the next stock ticker
                start = i + 1

                ' print the results in appropriate cells (ranges)
                Range("J" & 2 + j).Value = Cells(i, 1).Value
                Range("K" & 2 + j).Value = yearly_change
                Range("K" & 2 + j).NumberFormat = "0.00"
                Range("L" & 2 + j).Value = percent_change
                Range("L" & 2 + j).NumberFormat = "0.00%"
                Range("M" & 2 + j).Value = total

                ' fill cells with corresponding colors based on their values (green, red, blank)
                Select Case yearly_change
                ' set cell color to green for positive numbers
                    Case Is > 0
                        Range("K" & 2 + j).Interior.ColorIndex = 4
                    ' set cell color to red for negative numbers
                    Case Is < 0
                        Range("K" & 2 + j).Interior.ColorIndex = 3
                    ' set cell color to no fill color for 0s
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
            End If



            ' reset variables for next ticker
            total = 0
            yearly_change = 0
            j = j + 1


        ' If ticker doesn't change add values
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

    ' find the largest increase, largest decrease, and largest volume and place them in a separate part in the worksheet
    Range("R2") = "%" & WorksheetFunction.Max(Range("L2:L" & lastrow)) * 100
    Range("R3") = "%" & WorksheetFunction.Min(Range("L2:L" & lastrow)) * 100
    Range("R4") = WorksheetFunction.Max(Range("M2:M" & lastrow))

    ' calculate greatest increase, decrease, and volume number
    largest_increase = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
    largest_decrease = WorksheetFunction.Match(WorksheetFunction.Min(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
    vol_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("M2:M" & lastrow)), Range("M2:M" & lastrow), 0)

    ' display final ticker symbol for  total, largest increase. largest decrease, and largest volume
    Range("Q2") = Cells(largest_increase + 1, 10)
    Range("Q3") = Cells(largest_decrease + 1, 10)
    Range("Q4") = Cells(vol_number + 1, 10)

End Sub

