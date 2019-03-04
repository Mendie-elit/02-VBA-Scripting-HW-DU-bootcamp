Sub List_Unique_Values() 'Will create a list of unique values from Ticker Column

'Define Variables
Dim rSelection As Range
Dim vArray() As Long
Dim i As Long
Dim iColCount As Long
Dim pSelection As Range
Dim lastrow As Long

' Assign Values to variables
Set rSelection = Range("A:A")
Set pSelection = Range("I:I")
Set TickerRange = Range("A2:A40178")
Set TotalVolume = Range("J2:J40178")

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Total Volume"

'Copying selection
rSelection.Copy

'Past values and formats of column into new column
With Range("I1")
    .PasteSpecial xlPasteValues
    .PasteSpecial xlPasteFormats
    '.PasteSpecial xlPasteValuesAndNumberFormats
  End With

'Load array with column count
  'For use when multiple columns are selected
  iColCount = pSelection.Columns.Count
  ReDim vArray(1 To iColCount)
  For i = 1 To iColCount
    vArray(i) = i
  Next i
  
'Remove duplicates
  
ActiveSheet.Range("I:I").RemoveDuplicates Columns:=vArray(i - 1), Header:=xlGuess
  
ActiveSheet.Columns("I").AutoFit

'Exit CutCopyMod
Application.CutCopyMode = False

lastrow = Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrow

Cells(i, 10).FormulaR1C1 = "=SUMIF(C[-9],RC[-1],C[-3])"
    Range("J3").Select


Next i




End Sub