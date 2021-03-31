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
'checks if the ticker id is different
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








End Sub
