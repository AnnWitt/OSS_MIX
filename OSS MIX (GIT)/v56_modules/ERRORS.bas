Attribute VB_Name = "Errors"
Public Sub E1()

Dim PBI_SRC As Worksheet, PBI As Worksheet, e As Worksheet, Jira As Worksheet, EU_AA As Worksheet, INC_SRC As Worksheet, csv As Worksheet, sz As Worksheet

Dim ip_src As Integer, ip As Integer, ie As Integer, ij As Integer

Dim go As Worksheet

Set go = Worksheets("GO")
Set e = Worksheets("Errors")
Set PBI_SRC = Worksheets("PBI_Remedy")
Set INC_SRC = Worksheets("INC_Remedy")
Set PBI = Worksheets("Raport PBI")
Set Jira = Worksheets("JIRA OSS")
Set EA_SRC = Worksheets("EU_AA")
Set csv = Worksheets("CSV")
Set sz = Worksheets("Konfiguracja")

'czyszczenie
If WorksheetFunction.CountA(e.Columns(1)) > 1 Then
Range(e.Cells(2, "A"), e.Cells(WorksheetFunction.CountA(e.Columns(1)), "E")).Clear
Else
End If
'----


'ip_src = 2 'licznik po Ÿródle
'ip = 2 'licznik po wyniku
ie = 2

' Brak Ÿróde³ PBI
If WorksheetFunction.CountA(PBI_SRC.Columns(1)) = 0 Then
e.Cells(ie, "A") = sz.Cells(13, "X")
e.Cells(ie, "B") = "PBI_Remedy"
e.Cells(ie, "C") = "-"
e.Cells(ie, "D") = "Brak zrzutu z Remedy dla PBI"
ie = ie + 1
Else
End If

' Brak Ÿróde³ INC
If WorksheetFunction.CountA(INC_SRC.Columns(1)) = 0 Then
e.Cells(ie, "A") = sz.Cells(14, "X")
e.Cells(ie, "B") = "INC_Remedy"
e.Cells(ie, "C") = "-"
e.Cells(ie, "D") = "Brak zrzutu z Remedy dla INC"
ie = ie + 1
Else
End If


' Niepoprawny zrzut PBI
If PBI_SRC.Cells(1, 1) <> "Problem ID" Then
e.Cells(ie, "A") = sz.Cells(13, "X")
e.Cells(ie, "B") = "PBI_Remedy"
e.Cells(ie, "C") = "-"
e.Cells(ie, "D") = "Niepoprawny zrzut z Remedy dla PBI"
ie = ie + 1
Else
End If


' Brak ustawionej flagi stabilizacji
If Worksheets("GO").Cells(8, "N") = "" Then
e.Cells(ie, "A") = sz.Cells(13, "X")
e.Cells(ie, "B") = "Konfiguracja"
e.Cells(ie, "C") = "-"
e.Cells(ie, "D") = "Brak ustawionej flagi stabilizacyjnej"
ie = ie + 1
Else
End If


' Niepoprawny zrzut INC
If INC_SRC.Cells(1, 1) <> "Incident ID" Then
e.Cells(ie, "A") = sz.Cells(14, "X")
e.Cells(ie, "B") = "INC_Remedy"
e.Cells(ie, "C") = "-"
e.Cells(ie, "D") = "Niepoprawny zrzut z Remedy dla INC"
ie = ie + 1
Else
End If


If Jira.Cells(1, 1) <> sz.Cells(15, "X") Or Jira.Cells(1, "H") <> "SLA VC2" Then
e.Cells(ie, "A") = "Raport PBI/INC"
e.Cells(ie, "B") = "JIRA OSS"
e.Cells(ie, "C") = "-"
e.Cells(ie, "D") = "Brak zrzutu lub niepoprawny zrzut z Jira"
ie = ie + 1
End If

'brak przypisanych w tabeli userów - jira
For ij = 2 To WorksheetFunction.CountA(Jira.Columns(1))
    If Jira.Cells(ij, "R") = "E" Then 'v50
    e.Cells(ie, "A") = "Konfiguracja"
    e.Cells(ie, "B") = "Przypisani"
    e.Cells(ie, "C") = Jira.Cells(ij, "B")
    e.Cells(ie, "D") = "Osoby " & Jira.Cells(ij, "G") & " " & sz.Cells(41, "X")
    ie = ie + 1
    e.Cells(1, "H") = "X"
    End If
Next



'brak przypisanych w tabeli userów -ea
For ij = 2 To WorksheetFunction.CountA(EA_SRC.Columns(1))
    If (EA_SRC.Cells(ij, "R") = "E" And EA_SRC.Cells(ij, "F") <> "Unassigned") Then 'v52
    e.Cells(ie, "A") = "Konfiguracja"
    e.Cells(ie, "B") = "Przypisani"
    e.Cells(ie, "C") = EA_SRC.Cells(ij, "C")
    e.Cells(ie, "D") = "Osoby " & EA_SRC.Cells(ij, "F") & " " & sz.Cells(41, "X")
    ie = ie + 1
    End If
Next

'kompletnosc danych w emailach w konfiguracji

If (WorksheetFunction.CountA(sz.Columns(26)) = WorksheetFunction.CountA(sz.Columns(27))) And (WorksheetFunction.CountA(sz.Columns(26)) = WorksheetFunction.CountA(sz.Columns(28))) Then
'ok
Else
'blad
For i = 2 To Application.WorksheetFunction.max(WorksheetFunction.CountA(sz.Columns(26)), WorksheetFunction.CountA(sz.Columns(27)), WorksheetFunction.CountA(sz.Columns(28)))
    If (sz.Cells(i, 26) = "" Or sz.Cells(i, 27) = "" Or sz.Cells(i, 28) = "") Then
    e.Cells(ie, "A") = "Konfiguracja"
    e.Cells(ie, "B") = "Przypisani"
    e.Cells(ie, "C") = "-"
    e.Cells(ie, "D") = "Niepoprawnie wprowadzona osoba w wierszu " & i & " w arkuszu Konfiguracja"
    sz.Cells(i, 25).Interior.ColorIndex = 3
    ie = ie + 1
    Else
    End If
Next
End If

For ij = 2 To WorksheetFunction.CountA(Jira.Columns(1))
    If Jira.Cells(ij, "H") = "priorytet poza VC2" Then
    e.Cells(ie, "A") = sz.Cells(24, "X")
    e.Cells(ie, "B") = "Jira OSS"
    e.Cells(ie, "C") = Jira.Cells(ij, "B") & " - " & Jira.Cells(ij, "A")
    e.Cells(ie, "D") = "Niepoprawny priorytet w jira (poza VC2)"
    'E.Cells(ie, "E") = "E2"
    ie = ie + 1
    Else
    End If
Next

For ij = 2 To WorksheetFunction.CountA(Jira.Columns(1)) 'Unassigned
  If Jira.Cells(ij, "G") = "Unassigned" Then
    e.Cells(ie, "A") = sz.Cells(24, "X")
    e.Cells(ie, "B") = "Jira OSS"
    e.Cells(ie, "C") = Jira.Cells(ij, "B")
    e.Cells(ie, "D") = sz.Cells(47, "X") 'Zg?oszenie nieprzydzielone w Jira. Przydziel zg?oszenie i wygeneruj nowy zrzut z jira.
    ie = ie + 1
    Else
    End If
Next

For ij = 2 To WorksheetFunction.CountA(EA_SRC.Columns(1)) 'Unassigned
  If EA_SRC.Cells(ij, "F") = "Unassigned" Then
    e.Cells(ie, "A") = sz.Cells(18, "X")
    e.Cells(ie, "B") = "EU_AA"
    e.Cells(ie, "C") = EA_SRC.Cells(ij, "C")
    e.Cells(ie, "D") = sz.Cells(47, "X") 'Zg?oszenie nieprzydzielone w Jira. Przydziel zg?oszenie i wygeneruj nowy zrzut z jira.
    ie = ie + 1
    Else
    End If
    
    
    If ((EA_SRC.Cells(1, "C") <> "ID") And (EA_SRC.Cells(1, "C") <> "Key") And (EA_SRC.Cells(1, "C") <> "")) = True Then
    'v53 inny warunek ----------------------------------------
    e.Cells(ie, "A") = sz.Cells(18, "X")
    e.Cells(ie, "B") = "EU_AA"
    e.Cells(ie, "C") = "-"
    e.Cells(ie, "D") = sz.Cells(48, "X")
    go.Cells(2, "M") = "eu_error_rerun"
    Worksheets("EU_AA").Cells.Clear
    EA_SRC.Select
    On Error Resume Next
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.Cut
    Else
    End If
Next


'-----v52
For ip = 2 To WorksheetFunction.CountA(PBI.Columns(1))
    If IsError(Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(WorksheetFunction.CountA(Jira.Columns(1)), 18)), 8, 0)) = False Then 'v50
        If Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(WorksheetFunction.CountA(Jira.Columns(1)), 18)), 8, 0) <> PBI.Cells(ip, "L") Then 'v50
        e.Cells(ie, "A") = sz.Cells(24, "X")
        e.Cells(ie, "B") = "Raport PBI/JIRA OSS"
        e.Cells(ie, "C") = PBI.Cells(ip, "A") & " - " & PBI.Cells(ip, "B")
        e.Cells(ie, "D") = sz.Cells(25, "X") & " (JIRA - " & Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(WorksheetFunction.CountA(Jira.Columns(1)), 18)), 8, 0) & "/Remedy - " & PBI.Cells(ip, "L") & ")" 'v50
        ie = ie + 1
        Else
        End If

        If WorksheetFunction.CountIf(Range(Jira.Cells(2, 1), Jira.Cells(WorksheetFunction.CountA(Jira.Columns(1)), 1)), PBI.Cells(ip, "B")) > 1 Then
        e.Cells(ie, "A") = sz.Cells(24, "X")
        e.Cells(ie, "B") = "Raport PBI/JIRA OSS"
        e.Cells(ie, "C") = PBI.Cells(ip, "B")
        e.Cells(ie, "D") = sz.Cells(26, "X")
        ie = ie + 1
        Else
        End If

    Else
        e.Cells(ie, "A") = sz.Cells(24, "X")
        e.Cells(ie, "B") = "Raport PBI/JIRA OSS"
        e.Cells(ie, "C") = PBI.Cells(ip, "A") & " - " & PBI.Cells(ip, "B")
        e.Cells(ie, "D") = sz.Cells(27, "X")
       ' E.Cells(ie, "E") = "E2"
        ie = ie + 1
    End If
Next

'Next


End Sub

Public Sub E2()

Dim PBI_SRC As Worksheet, PBI As Worksheet, e As Worksheet, Jira As Worksheet, EU_AA As Worksheet, inc As Worksheet, csv As Worksheet
Dim ip_src As Integer, ip As Integer, ie As Integer, ij As Integer, ii As Integer, cnt As Integer, dd As Date, odp As Integer

Dim sz As Worksheet, go As Worksheet

Set sz = Worksheets("Konfiguracja")
Set inc = Worksheets("Raport INC")
Set e = Worksheets("Errors")
Set PBI_SRC = Worksheets("PBI_Remedy")
Set PBI = Worksheets("Raport PBI")
Set Jira = Worksheets("JIRA OSS")
Set EA_SRC = Worksheets("EU_AA")
Set csv = Worksheets("CSV")
Set go = Worksheets("go")

ie = WorksheetFunction.CountA(e.Columns(1)) + 1

'


' Brak Ÿróde³ INC_done/new

If Worksheets("GO").Cells(2, "K") <> "Tabela podsumowania" Then
    If (csv.Cells(2, "E") = "") Or (csv.Cells(2, "G") = "") Then
    e.Cells(ie, "A") = sz.Cells(16, "X")
    e.Cells(ie, "B") = "CSV"
    e.Cells(ie, "C") = "-"
    e.Cells(ie, "D") = sz.Cells(17, "X")
    ie = ie + 1
    Else
    End If
Else
End If



    If ((EA_SRC.Cells(1, "C") <> "ID") And (EA_SRC.Cells(1, "C") <> "Key")) Then
    e.Cells(ie, "A") = sz.Cells(18, "X")
    e.Cells(ie, "B") = "EU_AA"
    e.Cells(ie, "C") = "-"
    e.Cells(ie, "D") = sz.Cells(19, "X")
    go.Cells(2, "M") = "rerun"
    ie = ie + 1
    Else
    End If


For ip_src = 2 To WorksheetFunction.CountA(PBI_SRC.Columns(1))
    If PBI_SRC.Cells(ip_src, "K") = "" Then
    e.Cells(ie, "A") = "Remedy ITSM"
    e.Cells(ie, "B") = "-"
    e.Cells(ie, "C") = PBI_SRC.Cells(ip_src, "A")
    e.Cells(ie, "D") = sz.Cells(20, "X") & PBI_SRC.Cells(ip_src, "A") & " nie ma Assignera w Remedy ITSM, prawdopodobnie wymaga aktualizacji w JIRA"
    ie = ie + 1
    End If
Next

'draft
For ip_src = 2 To WorksheetFunction.CountA(PBI_SRC.Columns(1))
    If PBI_SRC.Cells(ip_src, "C") <> "Assigned" And PBI_SRC.Cells(ip_src, "C") <> "Pending" Then
    e.Cells(ie, "A") = "Remedy ITSM"
    e.Cells(ie, "B") = "Raport PBI"
    e.Cells(ie, "C") = PBI_SRC.Cells(ip_src, "A")
    e.Cells(ie, "D") = PBI_SRC.Cells(ip_src, "C") & " nie jest poprawnym statusem w Remedy ITSM. Skorygowano w raporcie ale wymagana jest korekta w Remedy ITSM."
    ie = ie + 1
    End If
Next



'-----------workarounds
For ip_src = 2 To WorksheetFunction.CountA(PBI_SRC.Columns(1))
    If PBI_SRC.Cells(ip_src, "E") = "" Then
    e.Cells(ie, "A") = "Remedy ITSM"
    e.Cells(ie, "B") = "Raport PBI"
    e.Cells(ie, "C") = PBI_SRC.Cells(ip_src, "A")
    e.Cells(ie, "D") = sz.Cells(42, "X")
    ie = ie + 1
    End If
Next
'-------


For ip = 2 To WorksheetFunction.CountA(PBI.Columns(1))
    If PBI.Cells(ip, "M") < PBI.Cells(ip, "N") Then 'data start < data obejsca
    Else
    e.Cells(ie, "A") = "Raport PBI"
    e.Cells(ie, "B") = "Raport PBI"
    e.Cells(ie, "C") = PBI.Cells(ip, "A") & " - " & PBI.Cells(ip, "B")
    e.Cells(ie, "D") = sz.Cells(21, "X")
    ie = ie + 1
    End If

    If PBI.Cells(ip, "M") < PBI.Cells(ip, "P") Then 'data start < data rozw
    Else
    e.Cells(ie, "A") = "Raport PBI"
    e.Cells(ie, "B") = "Raport PBI"
    e.Cells(ie, "C") = PBI.Cells(ip, "A") & " - " & PBI.Cells(ip, "B")
    e.Cells(ie, "D") = sz.Cells(22, "X")
    ie = ie + 1
    End If

    If PBI.Cells(ip, "N") < PBI.Cells(ip, "P") Then 'data obejscia mniejsza niz data rozwiazania
    Else
    e.Cells(ie, "A") = "Raport PBI"
    e.Cells(ie, "B") = "Raport PBI"
    e.Cells(ie, "C") = PBI.Cells(ip, "A") & " - " & PBI.Cells(ip, "B")
    e.Cells(ie, "D") = sz.Cells(23, "X")
    ie = ie + 1
    End If

Next





'--- dla inc

For ii = 2 To WorksheetFunction.CountA(inc.Columns(1))
    cnt = WorksheetFunction.CountIf(Range(Jira.Cells(2, 1), Jira.Cells(WorksheetFunction.CountA(Jira.Columns(1)), 1)), inc.Cells(ii, "A"))
            If cnt > 1 Then 'podwójne INC
            e.Cells(ie, "A") = sz.Cells(24, "X")
            e.Cells(ie, "B") = "Raport INC/JIRA OSS"
            e.Cells(ie, "C") = inc.Cells(ii, "A")
            e.Cells(ie, "D") = sz.Cells(28, "X") & """" & sz.Cells(15, "X") & """" & "~" & inc.Cells(ii, "A")
            ie = ie + 1
            Else
            End If

    cnt = cnt + (WorksheetFunction.CountIf(Range(Jira.Cells(2, 1), Jira.Cells(WorksheetFunction.CountA(Jira.Columns(1)), 1)), inc.Cells(ii, "E")))
    If cnt > 1 Then 'INC i PBI
    e.Cells(ie, "A") = sz.Cells(24, "X")
    e.Cells(ie, "B") = "Raport INC/JIRA OSS"
    e.Cells(ie, "C") = inc.Cells(ii, "A") & "/" & inc.Cells(ii, "E")
    e.Cells(ie, "D") = sz.Cells(29, "X") & """" & sz.Cells(15, "X") & """" & "~" & inc.Cells(ii, "A") & " or " & """" & sz.Cells(15, "X") & """" & "~" & inc.Cells(ii, "E")
    ' E.Cells(ie, "E") = "E2"
    ie = ie + 1
    Else
    End If
    If inc.Cells(ii, "R") = "" Then 'pusty opis z powodu zbyt d³ugieo stringa
            e.Cells(ie, "D") = sz.Cells(49, "X")
            e.Cells(ie, "A") = sz.Cells(50, "X")
            e.Cells(ie, "B") = sz.Cells(51, "X")
            e.Cells(ie, "C") = inc.Cells(ii, "A")
            inc.Cells(ii, "R").Interior.Color = RGB(130, 29, 29)
        
            ie = ie + 1
            Else
            End If
  
    


If IsError(Application.VLookup(inc.Cells(ii, "A"), Range(Jira.Cells(2, 1), Jira.Cells(WorksheetFunction.CountA(Jira.Columns(1)), 18)), 1, 0)) = False Then ' znalazl inca 'v50


    'data start - JAKIS PROBLEM TU BY£
    dd = Application.VLookup(inc.Cells(ii, "A"), Range(Jira.Cells(2, 1), Jira.Cells(WorksheetFunction.CountA(Jira.Columns(1)), 18)), 9, 0) ' zobacz czy  inie z PBI 'v50
    If Year(inc.Cells(ii, "I")) = Year(dd) And Month(inc.Cells(ii, "I")) = Month(dd) And Hour(inc.Cells(ii, "I")) = Hour(dd) And Minute(inc.Cells(ii, "I")) = Minute(dd) And Day(inc.Cells(ii, "I")) = Day(dd) Then
    'zgodne
    Else
    'niezgodne
    e.Cells(ie, "A") = sz.Cells(24, "X")
    e.Cells(ie, "B") = "Raport INC/JIRA"
    e.Cells(ie, "C") = inc.Cells(ii, "H") & " - " & inc.Cells(ii, "A")
    e.Cells(ie, "D") = sz.Cells(30, "X") & inc.Cells(ii, "I") & " vs JIRA " & dd & " (niepoprawna)"
    ie = ie + 1
    End If

    'data "obejscia"
    dd = Application.VLookup(inc.Cells(ii, "A"), Range(Jira.Cells(2, 1), Jira.Cells(WorksheetFunction.CountA(Jira.Columns(1)), 18)), 11, 0) ' zobacz czy  inie z PBI 'v50
    If Year(inc.Cells(ii, "J")) = Year(dd) And Month(inc.Cells(ii, "J")) = Month(dd) And Hour(inc.Cells(ii, "J")) = Hour(dd) And Minute(inc.Cells(ii, "J")) = Minute(dd) And Day(inc.Cells(ii, "J")) = Day(dd) Then
    'zgodne
    Else
    'niezgodne
    e.Cells(ie, "A") = sz.Cells(24, "X")
    e.Cells(ie, "B") = "Raport INC/JIRA"
    e.Cells(ie, "C") = inc.Cells(ii, "H") & " - " & inc.Cells(ii, "A")
    e.Cells(ie, "D") = sz.Cells(31, "X") & inc.Cells(ii, "J") & " vs JIRA " & dd & " (niepoprawna)"
    ie = ie + 1
    End If

Else

End If

Next








If WorksheetFunction.CountA(e.Columns(1)) = 1 Then
MsgBox "Raporty wykonane poprawnie", vbInformation
Else
MsgBox (sz.Cells(32, "X") & vbCrLf & sz.Cells(33, "X")), vbCritical
Worksheets("GO").Cells(2, "K") = "Tabela podsumowania"
End If

Worksheets("Errors").Activate

odp = MsgBox(sz.Cells(34, "X"), vbYesNo + vbQuestion)
If odp = vbYes Then
Call preexp
Else
End If



End Sub

