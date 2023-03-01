Attribute VB_Name = "Greatest"
Sub Statistic():

Dim Linha As Double
Dim Increase As Double
Dim Decrease As Double
Dim Volume As Double


Linha = 2 'First Line in ourdashboard'

With Sheet1 'Select the Dashboard Sheet'


.Range("H4").Value = WorksheetFunction.Max(.Range("D:D"))
.Range("H5").Value = WorksheetFunction.Min(.Range("D:D"))
.Range("H6").Value = WorksheetFunction.Max(.Range("E:E"))

.Range("I4").Value = WorksheetFunction.XLookup(.Range("H4").Value, .Range("D:D"), .Range("B:B"), 0, 0, -1)
.Range("I5").Value = WorksheetFunction.XLookup(.Range("H5").Value, .Range("D:D"), .Range("B:B"), 0, 0, -1)
.Range("I6").Value = WorksheetFunction.XLookup(.Range("H6").Value, .Range("E:E"), .Range("B:B"), 0, 0, -1)


End With


MsgBox ("Script Completed")

End Sub
