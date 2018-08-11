Sub stockdata()
'for all worksheets

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

Dim lastrow As Long
'lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
lastrow = 20
'create the variables

Dim openprice As Double
 Dim closeprice As Double
 Dim year As Double
 Dim ticker As String
 Dim percentchange As Double
 Dim volume As Double
 volume = 0
 
 'identify row and colums
 
 Dim row As Double
row = 2

Dim column As Integer
column = 1
'take the data

Dim x As Long
openprice = Cells(2, column + 2).Value


For x = 2 To lastrow
    If Cells(x + 1, column + 2).Value <> Cells(x, column).Value Then
'ticker
        ticker = Cells(x, column).Value
        Cells(row, colum + 8).Value = ticker

'closeprice

        closeprice = Cells(x, column + 5).Value

'yearchange
        years = closeprice - openprice
        Cells(row, column + 9).Value = percentagechange
'%change

    If (openprice = 0 And closeprice = 0) Then
       percentagechange = 0
    ElseIf (openprice = 0 And closeprice <> 0) Then
        percentagechange = 1
    Else: percentagechange = year / openprice
        Cells(row, column + 10).Value = percentagechange
        Cells(row, column + 10).NumberFormat = "0.00%"
'volume

    End If
    volume = volume + Cells(x, column + 6).Value
    Cells(row, column + 11).Value = volume
    row = row + 1
    openprice = Cells(x + 1, colum + 2)
    volume = 0
    
    'same ticker
    Else
        volume = volume + Cells(x, column + 6).Value
     End If
        
Next x
Next ws


End Sub