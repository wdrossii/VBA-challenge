Attribute VB_Name = "Module1"
'DR:  stock_analysis sub macro used to produce 2nd submission. arrayattempt macro left in for reference.
'DR:  stock_analysis sub macro copied from GITLab site, unknown author, uploaded by Steven Green.
'DR:  since solution code only ran for Q1 sheet, added functionality to run code on all quarters.

Sub stock_analysis():
    ' Set dimensions
    Dim total As Double
    Dim i As Long 'i for iterations, also possibly because this value arrives in column 'i'
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim ws As Worksheet
    
Worksheets("Q1").Select
    
    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    ' Set initial values
    j = 0
    total = 0
    change = 0
    start = 2

    ' get the row number of the last row with data.  DR:  I have used a longer form of this, actually activating an arbitrary cell then xlUp. I like this better.
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'DR:  cycle through the rows to analyze

    For i = 2 To rowCount

        ' If ticker changes then print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Stores results in variables
            total = total + Cells(i, 7).Value

            ' Handle zero total volume
            If total = 0 Then
                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0

            Else
                ' Find First non zero starting value
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' Calculate Change
                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = change / Cells(start, 3)

                ' start of the next stock ticker
                start = i + 1

                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).Value = percentChange
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("L" & 2 + j).Value = total

                ' colors positives green and negatives red
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select

            End If

            ' reset variables for new stock ticker
            total = 0
            change = 0
            j = j + 1
            days = 0

        ' If ticker is still the same add results
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

    ' take the max and min and place them in a separate part in the worksheet
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

    ' returns one less because header row not a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

    ' final ticker symbol for  total, greatest % of increase and decrease, and average
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)

'DR:  format cell q4 and column "L" due to large number - format "," with no decimals
    Range("Q4").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" 'code copied from stackoverflow.com and amended.
    Range("L:L").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" 'code copied from stackoverflow.com and amended.
    
'DR:  format result headings (I1 - L1, and O2 - O4) to be bold and italic.
    With Range("I1:L1")
    .Font.Italic = True 'code copied from learn.microsof.com and amended.
    .Font.Bold = True 'code copied from learn.microsof.com and amended.
    End With
    
    With Range("O2:O4")
    .Font.Italic = True 'code copied from learn.microsof.com and amended.
    .Font.Bold = True 'code copied from learn.microsof.com and amended.
    End With
    
Cells.EntireColumn.AutoFit 'DR: code added to refit the columns for different lengths of results.
    
Worksheets("Q2").Select

    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    ' Set initial values
    j = 0
    total = 0
    change = 0
    start = 2

    ' get the row number of the last row with data.  DR:  I have used a longer form of this, actually activating an arbitrary cell then xlUp. I like this better.
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'DR:  cycle through the rows to analyze

    For i = 2 To rowCount

        ' If ticker changes then print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Stores results in variables
            total = total + Cells(i, 7).Value

            ' Handle zero total volume
            If total = 0 Then
                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0

            Else
                ' Find First non zero starting value
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' Calculate Change
                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = change / Cells(start, 3)

                ' start of the next stock ticker
                start = i + 1

                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).Value = percentChange
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("L" & 2 + j).Value = total

                ' colors positives green and negatives red
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select

            End If

            ' reset variables for new stock ticker
            total = 0
            change = 0
            j = j + 1
            days = 0

        ' If ticker is still the same add results
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

    ' take the max and min and place them in a separate part in the worksheet
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

    ' returns one less because header row not a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

    ' final ticker symbol for  total, greatest % of increase and decrease, and average
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)

'DR:  format cell q4 and column "L" due to large number - format "," with no decimals
    Range("Q4").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" 'code copied from stackoverflow.com and amended.
    Range("L:L").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" 'code copied from stackoverflow.com and amended.

'DR:  format result headings (I1 - L1, and O2 - O4) to be bold and italic.
    With Range("I1:L1")
    .Font.Italic = True 'code copied from learn.microsof.com and amended.
    .Font.Bold = True 'code copied from learn.microsof.com and amended.
    End With

    With Range("O2:O4")
    .Font.Italic = True 'code copied from learn.microsof.com and amended.
    .Font.Bold = True 'code copied from learn.microsof.com and amended.
    End With

Cells.EntireColumn.AutoFit 'DR: code added to refit the columns for different lengths of results.

Worksheets("Q3").Select

    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    ' Set initial values
    j = 0
    total = 0
    change = 0
    start = 2

    ' get the row number of the last row with data.  DR:  I have used a longer form of this, actually activating an arbitrary cell then xlUp. I like this better.
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'DR:  cycle through the rows to analyze

    For i = 2 To rowCount

        ' If ticker changes then print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Stores results in variables
            total = total + Cells(i, 7).Value

            ' Handle zero total volume
            If total = 0 Then
                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0

            Else
                ' Find First non zero starting value
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' Calculate Change
                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = change / Cells(start, 3)

                ' start of the next stock ticker
                start = i + 1

                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).Value = percentChange
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("L" & 2 + j).Value = total

                ' colors positives green and negatives red
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select

            End If

            ' reset variables for new stock ticker
            total = 0
            change = 0
            j = j + 1
            days = 0

        ' If ticker is still the same add results
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

    ' take the max and min and place them in a separate part in the worksheet
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

    ' returns one less because header row not a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

    ' final ticker symbol for  total, greatest % of increase and decrease, and average
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)

'DR:  format cell q4 and column "L" due to large number - format "," with no decimals
    Range("Q4").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" 'code copied from stackoverflow.com and amended.
    Range("L:L").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" 'code copied from stackoverflow.com and amended.

'DR:  format result headings (I1 - L1, and O2 - O4) to be bold and italic.
    With Range("I1:L1")
    .Font.Italic = True 'code copied from learn.microsof.com and amended.
    .Font.Bold = True 'code copied from learn.microsof.com and amended.
    End With

    With Range("O2:O4")
    .Font.Italic = True 'code copied from learn.microsof.com and amended.
    .Font.Bold = True 'code copied from learn.microsof.com and amended.
    End With
    
Cells.EntireColumn.AutoFit 'DR: code added to refit the columns for different lengths of results.

Worksheets("Q4").Select

    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    ' Set initial values
    j = 0
    total = 0
    change = 0
    start = 2

    ' get the row number of the last row with data.  DR:  I have used a longer form of this, actually activating an arbitrary cell then xlUp. I like this better.
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'DR:  cycle through the rows to analyze

    For i = 2 To rowCount

        ' If ticker changes then print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Stores results in variables
            total = total + Cells(i, 7).Value

            ' Handle zero total volume
            If total = 0 Then
                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0

            Else
                ' Find First non zero starting value
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' Calculate Change
                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = change / Cells(start, 3)

                ' start of the next stock ticker
                start = i + 1

                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).Value = percentChange
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("L" & 2 + j).Value = total

                ' colors positives green and negatives red
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select

            End If

            ' reset variables for new stock ticker
            total = 0
            change = 0
            j = j + 1
            days = 0

        ' If ticker is still the same add results
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

    ' take the max and min and place them in a separate part in the worksheet
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

    ' returns one less because header row not a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

    ' final ticker symbol for  total, greatest % of increase and decrease, and average
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)

'DR:  format cell q4 and column "L" due to large number - format "," with no decimals
    Range("Q4").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" 'code copied from stackoverflow.com and amended.
    Range("L:L").NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" 'code copied from stackoverflow.com and amended.

'DR:  format result headings (I1 - L1, and O2 - O4) to be bold and italic.
    With Range("I1:L1")
    .Font.Italic = True 'code copied from learn.microsof.com and amended.
    .Font.Bold = True 'code copied from learn.microsof.com and amended.
    End With

    With Range("O2:O4")
    .Font.Italic = True 'code copied from learn.microsof.com and amended.
    .Font.Bold = True 'code copied from learn.microsof.com and amended.
    End With
    
Cells.EntireColumn.AutoFit 'DR: code added to refit the columns for different lengths of results.
    

End Sub

Sub arrayattempt()
Dim stkarray As Variant
Dim endrow As Single
Dim startcell As Single
Dim r As Integer
Dim c As Integer
Dim tablelength As Single
Dim rowCount As Integer

Application.ScreenUpdating = False
   
'Populate ticker symbol list.

Sheets("A").Select
Range("A999999").Select
    Selection.End(xlUp).Select
    endrow = ActiveCell.Row

stkarray = Sheets("A").Range("A2:A" & endrow & "").Value

Sheets("Q1").Select

Range("I2:I" & endrow & "").Value = stkarray

Sheets("B").Select

startcell = endrow + 1

Range("A999999").Select
    Selection.End(xlUp).Select
    endrow = ActiveCell.Row
tablelength = (startcell - 2) + endrow

stkarray = Sheets("B").Range("A2:A" & endrow & "").Value

Sheets("Q1").Select
Range("I" & startcell & ":I" & tablelength).Select
Selection.Value = stkarray

Sheets("C").Select

startcell = tablelength + 1

Range("A999999").Select
    Selection.End(xlUp).Select
    endrow = ActiveCell.Row
tablelength = (startcell - 3) + endrow

stkarray = Sheets("C").Range("A2:A" & endrow & "").Value

Sheets("Q1").Select
Range("I" & startcell - 1 & ":I" & tablelength).Select
Selection.Value = stkarray

Sheets("D").Select

startcell = tablelength + 1

Range("A999999").Select
    Selection.End(xlUp).Select
    endrow = ActiveCell.Row
tablelength = (startcell - 3) + endrow

stkarray = Sheets("D").Range("A2:A" & endrow & "").Value

Sheets("Q1").Select
Range("I" & startcell & ":I" & tablelength).Select
Selection.Value = stkarray

Sheets("E").Select

startcell = tablelength + 1

Range("A999999").Select
    Selection.End(xlUp).Select
    endrow = ActiveCell.Row
tablelength = (startcell - 3) + endrow

stkarray = Sheets("E").Range("A2:A" & endrow & "").Value

Sheets("Q1").Select
Range("I" & startcell & ":I" & tablelength).Select
Selection.Value = stkarray

Sheets("F").Select

startcell = tablelength + 1

Range("A999999").Select
    Selection.End(xlUp).Select
    endrow = ActiveCell.Row
tablelength = (startcell - 3) + endrow

stkarray = Sheets("F").Range("A2:A" & endrow & "").Value

Sheets("Q1").Select
Range("I" & startcell & ":I" & tablelength).Select
Selection.Value = stkarray

Application.ScreenUpdating = True

'remove duplicates from stock ticker list

Range("I:I").Select

Selection.RemoveDuplicates Columns:=1, Header:=xlYes

'Copy the resulting list to all "Q" tabs

Range("I:I").Select

stkarray = Selection.Value

Sheets("Q2").Range("I:I").Value = stkarray
Sheets("Q3").Range("I:I").Value = stkarray
Sheets("Q4").Range("I:I").Value = stkarray

'Calculate volumes for each ticker symbol and for each quater.

Sheets("Q1").Range("I2").Select
Selection.End(xlDown).Select
    endrow = ActiveCell.Row
    
Sheets("Q1").Range("L2:L" & endrow & "").Value = "=SUMIF(A:A,I2,G:G)"

Sheets("Q2").Activate
Range("I2").Select
Selection.End(xlDown).Select
    endrow = ActiveCell.Row
    
Sheets("Q2").Range("L2:L" & endrow & "").Value = "=SUMIF(A:A,I2,G:G)"

Sheets("Q3").Activate
Range("I2").Select
Selection.End(xlDown).Select
    endrow = ActiveCell.Row
    
Sheets("Q3").Range("L2:L" & endrow & "").Value = "=SUMIF(A:A,I2,G:G)"

Sheets("Q4").Activate
Range("I2").Select
Selection.End(xlDown).Select
    endrow = ActiveCell.Row
    
Sheets("Q4").Range("L2:L" & endrow & "").Value = "=SUMIF(A:A,I2,G:G)"

End Sub
