Attribute VB_Name = "Module1"
Sub calculator():

ticker = ""
nextrow = 2
openamount = 0
closeamount = 0
totalamount = 0
maxcolumn = (Cells(Rows.Count, 1).End(xlUp).Row)


greatestper = 0
lestper = 0
greatesttoal = 0



For i = 2 To maxcolumn

'cell values
currentcell = Cells(i, 1).Value
nextcell = Cells(i + 1, 1).Value
previouscell = Cells(i - 1, 1).Value



If (currentcell <> previouscell) Then
'checks if this is the first cell to get the open amount

    totalamount = totalamount + Cells(i, 7).Value
    openamount = Cells(i, 3).Value
    ticker = currentcell

ElseIf (currentcell = nextcell) Then
'checks if the ticker id is the same
    totalamount = totalamount + Cells(i, 7).Value
    ticker = currentcell

Else
'checks if the ticker id is different and prints results at the end'
    totalamount = totalamount + Cells(i, 7).Value
    closeamount = Cells(i, 6).Value
    
    yearlychange = closeamount - openamount
    If openamount = 0 Then
        perchange = 0
    
    Else
        perchange = (yearlychange / openamount)
        
    End If
    
    Cells(nextrow, 9).Value = ticker
    Cells(nextrow, 10).Value = yearlychange
    If (Cells(nextrow, 10).Value > 0) Then
        Cells(nextrow, 10).Interior.ColorIndex = 4
    Else
        Cells(nextrow, 10).Interior.ColorIndex = 3
    End If
    Cells(nextrow, 11).Value = perchange
    Cells(nextrow, 12).Value = totalamount
    nextrow = nextrow + 1
    totalamount = 0
    

End If


Next i


''''''''''''''''''''''''''''''''''''''''''''''' BOUNUS1 ''''''''''''''''''''''''''''''''''''''''''''''


startercell = Cells(2, 11).Value
startcell2 = Cells(3, 11).Value

maxcolumn2 = (Cells(Rows.Count, 11).End(xlUp).Row)

greatercell = 0
lessercell = 0
greatestticker = ""
lestticker = ""


If (startercell > startcell2) Then
' uses the first two cells to create the highest and lowest values '

    greatercell = startercell
    lessercell = startcell2

Else

    greatercell = startcell2
    lessercell = startercell

End If



For j = 4 To maxcolumn2
' uses for loop to check if the value is more or less than the highest or lowest values '
    currentcell2 = Cells(j, 11).Value

    
    If (currentcell2 > greatercell) Then
    greatercell = currentcell2
    greatestticker = Cells(j, 9).Value
    
    ElseIf (currentcell2 < lessercell) Then
    
    lessercell = currentcell2
    lestticker = Cells(j, 9).Value
    
    End If
    

Next j

' prints the results'

Cells(1, 14).Value = greatestticker
Cells(1, 15).Value = greatercell
Cells(2, 14).Value = lestticker
Cells(2, 15).Value = lessercell


''''''''''''''''''''''''''''''''''''''''''''''' BOUNUS2 ''''''''''''''''''''''''''''''''''''''''''''''



greatestticker2 = ""
greatestamount = 0

For k = 2 To maxcolumn2
' uses for loop to find largest amount'
    currentcell3 = Cells(k, 12).Value

    If (currentcell3 > greatestamount) Then
    
        greatestamount = currentcell3
        greatestticker2 = Cells(k, 9).Value
    End If

Next k

'prints out results '
Cells(3, 14).Value = greatestticker2
Cells(3, 15).Value = greatestamount



End Sub
