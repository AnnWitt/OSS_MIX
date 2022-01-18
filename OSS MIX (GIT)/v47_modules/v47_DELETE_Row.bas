Attribute VB_Name = "DELETE_Row"

Sub Delete_row_Klikniecie() 'OSS_MIX - DMP


Dim A As String, B As String, arkusz As String
Dim lp As Integer

Dim BIL As Worksheet, temp As Worksheet, sz As Worksheet, VC12 As Worksheet, WYK As Worksheet, ALL As Worksheet, oal As Worksheet, ow As Worksheet, SRC As Worksheet

Set ow = Worksheets("Oliver Wyman - INC")
Set BIL = Worksheets("Zestawienie Grup")
Set VC12 = Worksheets("VC1VC2")
Set ALL = Worksheets("VC2")
Set sz = Worksheets("Konfiguracja")
Set WYK = Worksheets("Wykresy_INC")
Set oal = Worksheets("OSS_ALL")
Set SRC = Worksheets("STAT_SRC")

R = WorksheetFunction.CountA(oal.Columns(1))
Range(oal.Cells(R, 1), oal.Cells(R, 15)).Clear
Union(Range(oal.Cells(R - 1, 7), oal.Cells(R - 1, 8)), Range(oal.Cells(R - 1, 14), oal.Cells(R - 1, 14))).ClearContents


R = WorksheetFunction.CountA(ALL.Columns(1))

'--DMP START
For lp = 3 To 38 ' v20210929
'--DMP STOP
    arkusz = sz.Cells(lp, "N")
    Set temp = Worksheets(arkusz)
    rmod = sz.Cells(lp, "O")
    Range(temp.Cells(R + rmod - 1, 1), temp.Cells(R + rmod - 1, 14)).Clear
    Range(temp.Cells(R + rmod - 2, 8), temp.Cells(R + rmod - 2, 10)).ClearContents

Next

'wyczyszczenie tez z arkusza grup
'--DMP START
Range(BIL.Cells(R - 1, 1), BIL.Cells(R - 1, 106)).Clear ' v20210929
'--DMP STOP


If VC12.Cells(WorksheetFunction.CountA(VC12.Columns(1)), "B") = ow.Cells(WorksheetFunction.CountA(ow.Columns(1)) + 1, "C") Then
ow.Rows(WorksheetFunction.CountA(ow.Columns(1)) + 1).Clear
Else
End If


'odswiezenie tabel przestawnych



SRC.PivotTables("Dane_wykres1").PivotCache.Refresh
SRC.PivotTables("Dane_wykres2").PivotCache.Refresh
A = Date
B = Date + 1
  SRC.PivotTables("Dane_wykres2").PivotFields(sz.Cells(35, "X")).ClearAllFilters
  SRC.PivotTables("Dane_wykres2").PivotFields(sz.Cells(35, "X")).PivotFilters _
       .Add Type:=xlBefore, Value1:=A

       SRC.PivotTables("Dane_wykres1").PivotFields(sz.Cells(35, "X")).ClearAllFilters
  SRC.PivotTables("Dane_wykres1").PivotFields(sz.Cells(35, "X")).PivotFilters.Add Type:=xlBefore, Value1:=B

 ActiveWorkbook.SlicerCaches("Fragmentator_Rok").ClearManualFilter
    ActiveWorkbook.SlicerCaches(sz.Cells(36, "X")).ClearManualFilter
    ActiveWorkbook.SlicerCaches(sz.Cells(37, "X")).ClearManualFilter
    
'    SRC.PivotTables("Dane_wykres2").PivotFields("Dzieñ").ClearAllFilters
'  SRC.PivotTables("Dane_wykres2").PivotFields("Dzieñ").PivotFilters _
'       .Add Type:=xlBefore, Value1:=A
'
'       SRC.PivotTables("Dane_wykres1").PivotFields("Dzieñ").ClearAllFilters
'  SRC.PivotTables("Dane_wykres1").PivotFields("Dzieñ").PivotFilters.Add Type:=xlBefore, Value1:=B
'
' ActiveWorkbook.SlicerCaches("Fragmentator_Rok").ClearManualFilter
'    ActiveWorkbook.SlicerCaches("Fragmentator_Miesi¹c").ClearManualFilter
'    ActiveWorkbook.SlicerCaches("Fragmentator_Dzieñ").ClearManualFilter
       
       
VC12.Select
MsgBox sz.Cells(38, "X")

End Sub

