Attribute VB_Name = "PIVOTS"
Sub pivot() 'OSS_MIX
Dim A As String
Dim B As String

Dim SRC As Worksheet

Set SRC = Worksheets("STAT_SRC")

Dim sz As Worksheet
Set sz = Worksheets("Konfiguracja")


SRC.PivotTables("Dane_wykres1").PivotCache.Refresh
SRC.PivotTables("Dane_wykres2").PivotCache.Refresh
A = Date
B = Date + 1
  SRC.PivotTables("Dane_wykres2").PivotFields("Dzien").ClearAllFilters
  SRC.PivotTables("Dane_wykres2").PivotFields("Dzien").PivotFilters _
       .Add Type:=xlBefore, Value1:=A
       
       SRC.PivotTables("Dane_wykres1").PivotFields("Dzien").ClearAllFilters
  SRC.PivotTables("Dane_wykres1").PivotFields("Dzien").PivotFilters.Add Type:=xlBefore, Value1:=B
       
 ActiveWorkbook.SlicerCaches("Fragmentator_Rok").ClearManualFilter 'moze rok warto nieczyscic
    ActiveWorkbook.SlicerCaches("Fragmentator_Miesiac").ClearManualFilter
    ActiveWorkbook.SlicerCaches("Fragmentator_Dzien").ClearManualFilter


End Sub

