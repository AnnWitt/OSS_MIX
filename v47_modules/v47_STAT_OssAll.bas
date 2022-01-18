Attribute VB_Name = "STAT_OssAll"
Sub nowy_wiersz_OSS()

Dim VC As Worksheet, OSAL As Worksheet, io As Integer, PBI As Worksheet

Set OSAL = Worksheets("OSS_ALL")
Set VC = Worksheets("VC2")
Set PBI = Worksheets("Raport PBI")

 ' ostatni niepusty


If OSAL.Cells(WorksheetFunction.CountA(OSAL.Columns(1)), "A") = VC.Cells(WorksheetFunction.CountA(VC.Columns(1)), "D") Then ' to nadpis wiersza
io = WorksheetFunction.CountA(OSAL.Columns(1))
Else 'dodac nowy wiersz
io = WorksheetFunction.CountA(OSAL.Columns(1)) + 1

'dane dot INC
OSAL.Cells(io, "A") = VC.Cells(WorksheetFunction.CountA(VC.Columns(1)), "D")
OSAL.Cells(io, "B") = VC.Cells(WorksheetFunction.CountA(VC.Columns(1)), "A")
OSAL.Cells(io, "C") = VC.Cells(WorksheetFunction.CountA(VC.Columns(1)), "B")
OSAL.Cells(io, "D") = Day(OSAL.Cells(io, "A"))
End If


'Dane dot PBI
OSAL.Cells(io, "E") = VC.Cells(WorksheetFunction.CountA(VC.Columns(1)), "F")
OSAL.Cells(io, "F") = VC.Cells(WorksheetFunction.CountA(VC.Columns(1)), "G")
OSAL.Cells(io - 1, "G") = VC.Cells((WorksheetFunction.CountA(VC.Columns(1))) - 1, "H")
OSAL.Cells(io - 1, "H") = VC.Cells((WorksheetFunction.CountA(VC.Columns(1))) - 1, "I")

OSAL.Cells(io, "K") = WorksheetFunction.CountIf(PBI.Columns("R"), "Nie")
OSAL.Cells(io, "L") = WorksheetFunction.CountIf(PBI.Columns("R"), "Tak")
OSAL.Cells(io, "M") = WorksheetFunction.CountIf(PBI.Columns("F"), "Pending")


    Range(OSAL.Cells(io, "A"), OSAL.Cells(io, "O")).Font.Size = 9
    Range(OSAL.Cells(io, "A"), OSAL.Cells(io, "O")).Font.Name = Calibri
    Range(OSAL.Cells(io, "B"), OSAL.Cells(io, "O")).NumberFormat = "general"
    OSAL.Cells(io, "A").NumberFormat = "dd.mm.yyyy"
    Union(Range(OSAL.Cells(io, "B"), OSAL.Cells(io, "D")), Range(OSAL.Cells(io, "I"), OSAL.Cells(io, "J"))).Interior.Color = RGB(217, 217, 217)



Range(OSAL.Cells(io - 3, "N"), OSAL.Cells(io, "O")).Interior.Pattern = xlNone

Range(OSAL.Cells(io - 1, "N"), OSAL.Cells(io - 1, "N")).Interior.Color = RGB(230, 55, 106)
Range(OSAL.Cells(io, "O"), OSAL.Cells(io, "O")).Interior.Color = RGB(230, 55, 106)

OSAL.Cells(io, "M") = WorksheetFunction.CountIf(PBI.Columns("F"), "Pending")


End Sub


Sub odsw()
Dim OSAL As Worksheet
Set OSAL = Worksheets("OSS_ALL")
Dim last As Integer
Dim sz As Worksheet
Set sz = Worksheets("Konfiguracja")

last = WorksheetFunction.CountA(OSAL.Columns(1))

OSAL.Cells(last, "J") = OSAL.Cells(last, "K") + OSAL.Cells(last, "L") + OSAL.Cells(last, "M")
OSAL.Cells(last, "I") = OSAL.Cells(last, "J") + OSAL.Cells(last, "O")

'OSAL.PivotTables(sz.Cells(39, "X")).PivotCache.Refresh
'OSAL.PivotTables(sz.Cells(40, "X")).PivotCache.Refresh
OSAL.PivotTables("suma_orange_t4").PivotCache.Refresh
OSAL.PivotTables("suma_atos_t3").PivotCache.Refresh


Call E2

End Sub
