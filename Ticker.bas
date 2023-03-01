Attribute VB_Name = "Ticker"
Sub Ticker():

Dim Linha As Long
Dim W As Worksheet
Dim Ws As Worksheet


Set W = Sheets("2018")
Set Ws = Sheets("Dashboard Stock")


W.Select
Linha = 2
Ws.Select
Range("B3").Select

With W

Do Until .Cells(Linha, 1) = ""

Ws.Select

ActiveCell.Value = .Cells(Linha, 1)
ActiveCell.Offset(1, 0).Select
Linha = Linha + 1
Loop



End With

ActiveSheet.Range("B:B").RemoveDuplicates Columns:=1, Header:=xlYes
Range("B3").Select

MsgBox ("Script Completed")

End Sub
