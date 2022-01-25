Attribute VB_Name = "START_CzyszczenieObrobkaDanych"
Sub czyszczenie()
Dim PBI_SRC As Worksheet, PBI As Worksheet, Jira As Worksheet, EA_SRC As Worksheet, ea As Worksheet, INC_SRC As Worksheet, inc As Worksheet, csv As Worksheet, e As Worksheet
Dim go As Worksheet

Dim ip_jira As Integer, ip As Integer, users As Integer, i As Integer
Dim rerun_jira As Boolean, rerun_eu As Boolean
rerun_jira = False
rerun_eu_e1 = False
rerun_eu_e2 = False
rerun_eu = False

Dim sz As Worksheet
Set sz = Worksheets("Konfiguracja")
Set go = Worksheets("GO")
Set PBI_SRC = Worksheets("PBI_Remedy")
Set INC_SRC = Worksheets("INC_Remedy")
Set PBI = Worksheets("Raport PBI")
Set inc = Worksheets("Raport INC")
Set Jira = Worksheets("JIRA OSS")
Set EA_SRC = Worksheets("EU_AA")
Set csv = Worksheets("CSV")
Set ea = Worksheets("Zadania ADM i DEV")
Set e = Worksheets("Errors")

Application.Calculation = xlManual


sz.Columns(25).Clear

If go.Cells(2, "M") = "first" Then
go.Cells(2, "M") = "rerun"
Else
End If

ip = 2 'licznik po wyniku

'rerun v53 - polaczyc ?
For i = 2 To WorksheetFunction.CountA(e.Columns(1))
        If go.Cells(2, "M") = "eu_error_rerun" Then
        go.Cells(2, "M") = "rerun"
        e.Cells(i, "D") = "nowy zrzut"
        Else
        End If
        If e.Cells(i, "D") = "Brak zrzutu lub niepoprawny zrzut z Jira" Then
        rerun_jira = True
        go.Cells(2, "M") = "rerun"
        End If
        If e.Cells(i, "D") = sz.Cells(48, "X") Then 'niepoprawny
        rerun_eu = True
        go.Cells(2, "M") = "eu_error_rerun"
        Worksheets("EU_AA").Cells.Clear 'sprawdz
        End If
        If e.Cells(i, "D") = sz.Cells(19, "X") Then 'brak
        rerun_eu = True
        go.Cells(2, "M") = "rerun"
        End If
    'Or e.Cells(i, "D") = sz.Cells(19, "X")
  
Next


'czyszczenie lgou b³êdów
If WorksheetFunction.CountA(e.Columns(1)) > 1 Then
Range(e.Cells(2, "A"), e.Cells(WorksheetFunction.CountA(e.Columns(1)), "D")).Clear
e.Cells(1, "H").Clear
Else
End If


'----czyszczenie œmieci z Remedy PBI
If PBI_SRC.Cells(1, 1) = "Problem ID*+" Then
PBI_SRC.Rows(WorksheetFunction.CountA(PBI_SRC.Columns(1))).Clear
PBI_SRC.Rows(WorksheetFunction.CountA(PBI_SRC.Columns(1))).Clear
PBI_SRC.Columns("F:I").Replace What:="-", Replacement:="-", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
PBI_SRC.Columns("F:I").NumberFormat = "yyyy/mm/dd hh:mm:ss"
PBI_SRC.Cells(1, 1) = "Problem ID"
Else
End If
'----koniec czyszczenia + zabezpieczenie aby nie usuwa³ w nieskonczonosc

'----czyszczenie œmieci z Remedy INC
If INC_SRC.Cells(1, 1) = "Incident ID*+" Then
INC_SRC.Rows(WorksheetFunction.CountA(INC_SRC.Columns(1))).Clear
INC_SRC.Rows(WorksheetFunction.CountA(INC_SRC.Columns(1))).Clear
INC_SRC.Columns("G:H").Replace What:="-", Replacement:="-", LookAt:=xlPart, _
SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
ReplaceFormat:=False
INC_SRC.Columns("G:H").NumberFormat = "yyyy/mm/dd hh:mm:ss"
INC_SRC.Cells(1, 1) = "Incident ID"
Else
End If
'----koniec czyszczenia + zabezpieczenie aby nie usuwa³ w nieskonczonosc


'Niezgodno?? priorytetów mi?dzy Remedy a JIRA (JIRA - Awaria/Remedy - Critical)
'Niezgodno?? priorytetów mi?dzy Remedy a JIRA (JIRA - Low/Remedy - Medium)

users = WorksheetFunction.CountA(sz.Columns(26))


'e.Cells(1, "H") = "X"

'obróbka zrzutu z jiry





If (Worksheets("GO").Cells(2, "K") <> "Tabela podsumowania") Or (go.Cells(2, "M") = "rerun") Or (go.Cells(2, "M") = "eu_error_rerun") Then 'v53
    If (Jira.Cells(4, 2) = "ID" Or Jira.Cells(4, 2) = "Key") Then 'koniecznosc czyszczenia
    'Jira.Rows(WorksheetFunction.CountA(Jira.Columns(1)) + 1).Clear
    Jira.Rows(WorksheetFunction.CountA(Jira.Columns(1)) + 1).Clear
    Jira.Rows(WorksheetFunction.CountA(Jira.Columns(1)) + 1).UnMerge
    Jira.Rows(1).UnMerge
    Jira.Rows(2).UnMerge
    Jira.Rows(3).UnMerge
    Jira.Rows(1).Clear
    Jira.Rows(2).Clear
    Jira.Rows(3).Clear
    Jira.Cells(4, "T") = 2 'v50
    'Jira.Columns(1).Delete
        For ip_jira = 5 To WorksheetFunction.CountA(Jira.Columns(1)) + 3
        
        Jira.Cells(ip_jira, "T") = ip_jira 'v50
        Jira.Cells(ip_jira, "H") = Right(Jira.Cells(ip_jira, "H"), Len(Jira.Cells(ip_jira, "H")) - 4)
        '-- podmiana assignerów
        If IsError(Application.VLookup(Jira.Cells(ip_jira, 7), Range(sz.Cells(1, 26), sz.Cells(users, 27)), 2, 0)) = True Then
        Jira.Cells(ip_jira, "S") = "E" '//wpis do errors E1 'v50
        Else
        Jira.Cells(ip_jira, 7) = Application.VLookup(Jira.Cells(ip_jira, 7), Range(sz.Cells(1, 26), sz.Cells(users, 27)), 2, 0)
        End If
        Next
    Jira.Columns("A:T").Sort Key1:=Jira.Columns("T"), Order1:=xlAscending 'v50
    Jira.Columns(20).Clear 'v50
    On Error Resume Next
         Jira.Activate
        ActiveSheet.Shapes.Range(Array("Picture 1")).Select
        Selection.Cut
        Resume Next
    End If

'dla zadan EU
If (EA_SRC.Cells(4, 1) = "Typ Zadania" Or EA_SRC.Cells(4, 1) = "Issue Type") Then  'rusza obróbka
EA_SRC.Rows(WorksheetFunction.CountA(EA_SRC.Columns(1)) + 1).Clear
EA_SRC.Rows(WorksheetFunction.CountA(EA_SRC.Columns(1)) + 1).UnMerge
EA_SRC.Rows(1).Clear
EA_SRC.Rows(1).UnMerge
EA_SRC.Rows(2).Clear
EA_SRC.Rows(2).UnMerge
EA_SRC.Rows(3).Clear
EA_SRC.Rows(3).UnMerge

EA_SRC.Cells(4, "Q") = 4 'v50

    For ip_jira = 5 To WorksheetFunction.CountA(EA_SRC.Columns(1)) + 3
    EA_SRC.Cells(ip_jira, "Q") = ip_jira 'v50
    
    
        '-- podmiana assignerów
    If IsError(Application.VLookup(EA_SRC.Cells(ip_jira, 6), Range(sz.Cells(1, 26), sz.Cells(users, 27)), 2, 0)) = True Then
    EA_SRC.Cells(ip_jira, "R") = "E" '//wpis do errors E1 v50
    Else
    EA_SRC.Cells(ip_jira, 6) = Application.VLookup(EA_SRC.Cells(ip_jira, 6), Range(sz.Cells(1, 26), sz.Cells(users, 27)), 2, 0)
    End If
    '--
    
    Next

EA_SRC.Columns("A:R").Sort Key1:=EA_SRC.Columns("Q"), Order1:=xlAscending 'v50
EA_SRC.Columns("Q").Clear 'v50


On Error Resume Next
    EA_SRC.Activate
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.Cut
    Resume Next
End If


'----- wpis errors


Else



End If


'koniec obróbki Jiry

'----czyszczenie danych z Raport PBI (do W) 'v50
If WorksheetFunction.CountA(PBI.Columns(2)) > 1 Then
Range(PBI.Cells(2, "A"), PBI.Cells(WorksheetFunction.CountA(PBI.Columns(2)), "W")).Clear 'v50
Else
End If

'---czyszczenie dla EU
If WorksheetFunction.CountA(ea.Columns(1)) > 1 Then
Range(ea.Cells(2, "A"), ea.Cells(WorksheetFunction.CountA(ea.Columns(1)), "P")).Clear 'v50
Else
End If

'---czyszczenie dla INC
If WorksheetFunction.CountA(inc.Columns(1)) > 1 Then
Range(inc.Cells(2, "A"), inc.Cells(WorksheetFunction.CountA(inc.Columns(1)) + 1, "S")).Clear 'v50
inc.Columns("T:U").Clear '50
Else
End If

'---czyszczenie dla INC

If Worksheets("GO").Cells(13, "O") = "Tak" Then
    If WorksheetFunction.CountA(csv.Columns(1)) > 1 Then
    csv.Columns("A:I").ClearContents
    csv.Cells(1, "A") = "Vendor_open_all"
    csv.Cells(1, "C") = "Vendor_SLA"
    csv.Cells(1, "E") = "Vendor_daily_done"
    csv.Cells(1, "G") = "Vendor_daily_new"
    csv.Cells(1, "I") = "Vendor_daily_sla_done"
    Else
    End If
Else
End If
Application.Calculation = xlAutomatic

End Sub

Sub filtry()

Dim sz As Worksheet
Set sz = Worksheets("Konfiguracja")

Dim temp As Worksheet, lp As Integer, arkusz As String

On Error Resume Next
'---DMP START
For lp = 2 To 38 'v20210929
'---DMP STOP
    arkusz = sz.Cells(lp, "N")
    Set temp = Worksheets(arkusz)
    temp.ShowAllData
Next


Worksheets("Konfiguracja").ShowAllData
Worksheets("TO DO").ShowAllData
Worksheets("Metryka zmian").ShowAllData
Worksheets("PBI_Remedy").ShowAllData
Worksheets("INC_Remedy").ShowAllData
Worksheets("CSV").ShowAllData
Worksheets("JIRA OSS").ShowAllData
Worksheets("EU_AA").ShowAllData
Worksheets("STAT_SRC").ShowAllData
Worksheets("GO").ShowAllData
Worksheets("Errors").ShowAllData
Worksheets("OSS_ALL").ShowAllData
Worksheets("Legenda").ShowAllData
Worksheets("Raport PBI").ShowAllData
Worksheets("Raport INC").ShowAllData
Worksheets("Zadania ADM i DEV").ShowAllData
Worksheets("Zestawienie grup").ShowAllData
Worksheets("Oliver Wyman - INC").ShowAllData
Worksheets("Daily").ShowAllData


End Sub

