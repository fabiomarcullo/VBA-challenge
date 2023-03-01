Attribute VB_Name = "Sum_Stock"
Sub Sum_Stok():

Dim Linha As Double
Dim wsName As String


wsName = Range("H2").Value

Linha = 2 'First Line in ourdashboard'

With Sheet1 'Select the Dashboard Sheet'

Do

Linha = Linha + 1 'Add on Line in our Loop'
.Cells(Linha, 5).Value = WorksheetFunction.SumIfs(Worksheets(wsName).Range("G:G"), Worksheets(wsName).Range("A:A"), .Cells(Linha, 2).Text) 'Sum for each row as reference Stock Number'


Loop Until .Cells(Linha + 1, 2) = Empty 'Loop until column is empty'

End With


MsgBox ("Script Completed")

End Sub

