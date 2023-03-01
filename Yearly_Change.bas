Attribute VB_Name = "Yearly_Change"
Sub Yarly_Change():

Dim Linha As Double
Dim A As Double
Dim B As Double
Dim wsName As String

wsName = Range("H2").Value
Linha = 2 'First Line in ourdashboard'



With Sheet1 'Select the Dashboard Sheet'

Do



Linha = Linha + 1 'Add on Line in our Loop'
.Cells(Linha, 5).Value = WorksheetFunction.SumIfs(Worksheets(wsName).Range("G:G"), Worksheets(wsName).Range("A:A"), .Cells(Linha, 2).Text) 'Sum for each row as reference Stock Number''
A = WorksheetFunction.Index(Worksheets(wsName).Range("A:G"), WorksheetFunction.Match(.Cells(Linha, 2).Text, Worksheets(wsName).Range("A:A"), 0), 3)  '1st row for each row as reference Stock Number'


B = WorksheetFunction.VLookup(.Cells(Linha, 2).Text, Worksheets(wsName).Range("A:F"), 6, 1) 'Last row for each row as reference Stock Number'

.Cells(Linha, 3).Value = B - A
.Cells(Linha, 4).Value = (B / A) - 1


If .Cells(Linha, 3).Value >= 0 Then
            
                .Cells(Linha, 3).Interior.ColorIndex = 43
            
            Else
                
                .Cells(Linha, 3).Interior.ColorIndex = 3
                
            End If
            
If .Cells(Linha, 4).Value >= 0 Then
            
                .Cells(Linha, 4).Interior.ColorIndex = 43
            
            Else
                
                .Cells(Linha, 4).Interior.ColorIndex = 3
                
            End If

Loop Until .Cells(Linha + 1, 2) = Empty 'Loop until column is empty'

End With

MsgBox ("Script Completed")
End Sub





