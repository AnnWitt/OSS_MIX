Attribute VB_Name = "FINAL_eksportRaportu"
Sub preexp()

If Worksheets("emails").Cells(1, 1) = "" Then
Call adresaci
Else
End If

End Sub

Sub exp()

Application.Calculation = xlCalculationAutomatic

Dim Legenda As Worksheet, m As Worksheet
Dim folder As String, data As String, godz As String, plik As String, mi As Integer, old As String, old_src As String

Set m = Worksheets("Metryka zmian")
'Set inc = Worksheets("Raport INC") v50 niepotrzebne
Set L = Worksheets("Legenda")

Dim sz As Worksheet
Set sz = Worksheets("Konfiguracja")
'
'If Worksheets("emails").Cells(1, 1) = "" Then
'Call adresaci
'Else
'End If

If Worksheets("GO").Cells(13, "O") = "Nie" Then
        odp = MsgBox(sz.Cells(3, "X") & vbline & sz.Cells(4, "X"), vbYesNo + vbQuestion)
        If odp = vbYes Then
        Worksheets("GO").Cells(13, "O") = "Tak"
        Else
        
        End If
Else
End If




mi = WorksheetFunction.CountA(m.Columns(1))
Application.DisplayAlerts = False


old = ActiveWorkbook.FullName
old_src = ActiveWorkbook.Path
'MsgBox ActiveWorkbook.FullName 'zwraca pelna sciezke do pliku wraz z nazw¹

On Error Resume Next ' na wypadek gdyby katalog juz by³



'folder = ActiveWorkbook.Path & "\" & "Raport dzienny " & Date & " OSS_MIX"


If Worksheets("GO").Cells(10, "K") = "Tak" Then
Worksheets("Daily").ExportAsFixedFormat Type:=xlTypePDF, Filename:=folder & "\" & Worksheets("Daily").Cells(1, 1) & " OSS_INC.pdf"
Else

End If


If Len(Month(Now())) < 2 Then
    data = Year(Now()) & "0" & Month(Now())
Else
    data = Year(Now()) & Month(Now())
End If

If Len(Day(Now())) < 2 Then
    data = data & "0" & Day(Now())
Else
    data = data & Day(Now())
End If


If Len(Hour(Now())) < 2 Then
    godz = "0" & Hour(Now())
Else
    godz = Hour(Now())
End If

If Len(Minute(Now())) < 2 Then
    godz = godz & "0" & Minute(Now())
Else
    godz = godz & Minute(Now())
End If

folder = ActiveWorkbook.Path & "\" & data & " Raport dzienny " & " OSS_MIX"

MkDir folder

plik = "RaportDzienny " & data & "_" & godz & ".xlsx"

ActiveWorkbook.SaveAs Filename:=old_src & "\" & data & "_" & godz & " " & m.Cells(mi, "C") & ".xlsm"
ActiveWorkbook.SaveAs Filename:=folder & "\" & plik, FileFormat:=xlOpenXMLWorkbook


x = ActiveWorkbook.FullName

Workbook.Open (x)

Dim lp As Integer, arkusz As String, temp As Worksheet

''---DMP START
For lp = 2 To 38 'v20210929
'---DMP STOP

    arkusz = sz.Cells(lp, "N")
    Set temp = Worksheets(arkusz)
    temp.Delete
Next

Worksheets("CSV").Delete
Worksheets("Wykresy_INC").Delete
Worksheets("Wykresy_OSS").Delete
Worksheets("STAT_SRC").Delete
Worksheets("Oliver Wyman - INC").Delete
Worksheets("Konfiguracja").Delete
Worksheets("TO DO").Delete
Worksheets("GO").Delete
Worksheets("JIRA OSS").Delete
Worksheets("EU_AA").Delete
Worksheets("PBI_Remedy").Delete
Worksheets("INC_Remedy").Delete
Worksheets("STAT_SRC").Delete
Worksheets("Errors").Delete
Worksheets("emails").Delete
Worksheets("OSS_ALL").Delete
Worksheets("Metryka zmian").Delete
Worksheets("Daily").Delete
Worksheets("Zestawienie Grup").Delete



ActiveWorkbook.SaveAs Filename:=folder & "\" & plik, FileFormat:=xlOpenXMLWorkbook


Kill old

'Application.Quit

MsgBox sz.Cells(12, "X")


End Sub

