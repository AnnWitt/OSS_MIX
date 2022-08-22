Attribute VB_Name = "START_Wywolania"

Public Sub start() 'póki co sam dzienny
Dim e As Worksheet
Dim go As Worksheet
Dim jira As Worksheet
Dim exportB As Shape


Set jira = Worksheets("JIRA OSS")

'info: dodawanie nowego produktu: wiersze DMP, nowy arkusz, modyfikacje w arkuszu Konfiguracja, STAT_SRC,oliver wyman, zestawienie grup
'mo¿liwe, ¿e równie¿ zmiany w zapytaniach (modul selecty + arkusz GO)

Set go = Worksheets("GO")

'If walidacja = False Then
'Call replace
'End If

walidacja = True
Application.ScreenUpdating = False
Call czyszczenie
If go.Cells(2, "M") = "rerun" Then
Call replace
Else
End If

Call filtry


Call E1 'kontrola b³êdów zrzutów

Worksheets("emails").Cells.ClearContents


Dim sz As Worksheet
Set sz = Worksheets("Konfiguracja")

sz.Columns(25).Clear

If Worksheets("Errors").Cells(2, 1) <> "" Then
Worksheets("Errors").Activate
MsgBox sz.Cells(7, "X")
Else

jira.Columns(18).Clear 'v50
Worksheets("EU_AA").Columns(18).Clear 'v50
Call PBI


Worksheets("Errors").Shapes("exportB").Visible = True
Worksheets("Errors").Shapes("assigneeCorrect").Visible = False
Worksheets("Errors").Shapes("rerun").Visible = False

If walidacja = False Then
Worksheets("Errors").Shapes("exportB").Visible = False
Worksheets("Errors").Shapes("assigneeCorrect").Visible = True
Worksheets("Errors").Shapes("rerun").Visible = False
go.Cells(2, "M") = "rerun"
Else


Call ZUA
Call inc
Call VilTul
Call Oliver_Wyman


'czy ze zrzutów czy z tabelki
If go.Cells(2, "K") = sz.Cells(5, "X") Then
Call Przycisk6_Klikniecie
Else
    If Worksheets("STAT_SRC").Cells(3, "B") <> "" Then
    Call dane_zrodlo_Klikniecie
    Else
    MsgBox (sz.Cells(8, "X") & vbCrLf & sz.Cells(9, "X"))
    ie = WorksheetFunction.CountA(Worksheets("Errors").Columns(1)) + 1
    Worksheets("Errors").Cells(ie, "A") = sz.Cells(10, "X")
    Worksheets("Errors").Cells(ie, "B") = "STAT_SRC"
    Worksheets("Errors").Cells(ie, "C") = "-"
    Worksheets("Errors").Cells(ie, "D") = sz.Cells(11, "X")
    Worksheets("Errors").Cells(1, "H") = "X"
    End If
End If



Call nowy_wiersz_OSS

Worksheets("OSS_ALL").Activate


End If
Application.ScreenUpdating = True
End If
End Sub





Sub Przycisk6_Klikniecie() ' plan naprawczy poczatek "rozmiesc dane" //uzupelnianie z z opcji tabela danych


Dim A As String
Dim B As String

Dim csv As Worksheet
Set csv = Worksheets("CSV")

Call count_groups

Call nowy_wiersz_auto

Call add_data

Call bilans

'Call pivot
       
    


End Sub

