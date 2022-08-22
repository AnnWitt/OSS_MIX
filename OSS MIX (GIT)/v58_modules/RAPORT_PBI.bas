Attribute VB_Name = "RAPORT_PBI"

Option Explicit
Global walidacja As Boolean



Sub PBI()
Dim PBI_SRC As Worksheet, PBI As Worksheet, jira As Worksheet, sz As Worksheet, e As Worksheet, ie As Integer

Dim ip_jira As Integer, ip As Integer



Set PBI_SRC = Worksheets("PBI_Remedy")
Set PBI = Worksheets("Raport PBI")
Set jira = Worksheets("JIRA OSS")
Set sz = Worksheets("Konfiguracja")
Set e = Worksheets("Errors")




ip_jira = WorksheetFunction.CountA(jira.Columns(1)) 'zamkni�cie zakresu

For ip = 2 To WorksheetFunction.CountA(PBI_SRC.Columns(1))
'dane z Remedy - w tym przypadku moze by� ten sam licznik, przy jira nie

PBI.Cells(ip, "B") = PBI_SRC.Cells(ip, "A")
PBI.Cells(ip, "E") = PBI_SRC.Cells(ip, "B")
PBI.Cells(ip, "F") = PBI_SRC.Cells(ip, "C")

If PBI.Cells(ip, "F") <> "Assigned" And PBI.Cells(ip, "F") <> "Pending" Then
PBI.Cells(ip, "F") = "Assigned"
Else
End If


PBI.Cells(ip, "G") = PBI_SRC.Cells(ip, "D")
PBI.Cells(ip, "L") = PBI_SRC.Cells(ip, "J")
PBI.Cells(ip, "M") = PBI_SRC.Cells(ip, "G") 'z jiry
PBI.Cells(ip, "N") = PBI_SRC.Cells(ip, "I") 'z jiry
PBI.Cells(ip, "P") = PBI_SRC.Cells(ip, "H") 'z jiry
Union(PBI.Columns("M:N"), PBI.Columns("P:P"), PBI.Columns("T:T"), PBI.Columns("W:W")).NumberFormat = "yyyy/mm/dd hh:mm:ss" 'v50
    If PBI_SRC.Cells(ip, "E") <> "" Then
    PBI.Cells(ip, "O") = "Tak"
    Else
    PBI.Cells(ip, "O") = "Nie"
    End If
'Dane z jira
    If IsError(Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 3, 0)) = True Then 'v50
    PBI.Cells(ip, "A") = "-"
    PBI.Cells(ip, "C") = "-"
    PBI.Cells(ip, "D") = "-"
    PBI.Cells(ip, "H") = "-"
    PBI.Cells(ip, "I") = "-"
    PBI.Cells(ip, "J") = "-"
    PBI.Cells(ip, "K") = "-"
    PBI.Cells(ip, "S") = "-"
    PBI.Cells(ip, "T") = "-"
    PBI.Cells(ip, "U") = "-"
    PBI.Cells(ip, "V") = "-"
    PBI.Cells(ip, "Q") = "-"
    PBI.Cells(ip, "R") = "-"
     PBI.Cells(ip, "W") = "-" 'v50
    Else 'v50
    PBI.Cells(ip, "A") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 2, 0)
    PBI.Cells(ip, "C") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 3, 0)
    PBI.Cells(ip, "D") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 4, 0)
    PBI.Cells(ip, "H") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 5, 0)
    PBI.Cells(ip, "I") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 6, 0)
    
    PBI.Cells(ip, "J") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 7, 0) 'left 3 lub 2 czyli ew wyciac spacje
    PBI.Cells(ip, "J") = Trim(Left(Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 7, 0), 3))
     
   
    PBI.Cells(ip, "K") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 7, 0)
    
    If (PBI.Cells(ip, "K") = "nPPKuser") Or (PBI.Cells(ip, "K") = "OCLuser") Then
    PBI.Cells(ip, "J") = "DEV"
    Else
    End If
    
    If (PBI.Cells(ip, "K") = "Nieprzydzielone") Then
    PBI.Cells(ip, "J") = "#ND"
    Else
    End If
    
    
      'blokada ! --  bedzie musiala by� dla zadan i inc
    If (PBI.Cells(ip, "J") <> "PM" And PBI.Cells(ip, "J") <> "ANL" And PBI.Cells(ip, "J") <> "TST" And PBI.Cells(ip, "J") <> "#ND" And PBI.Cells(ip, "J") <> "ADM" And PBI.Cells(ip, "J") <> "DEV" And PBI.Cells(ip, "J") <> " PM") Then
    Worksheets("Errors").Activate
    ie = WorksheetFunction.CountA(e.Columns(1)) + 1
    e.Cells(ie, "A") = "Jira/Szablon"
    e.Cells(ie, "B") = "Konfiguracja"
    e.Cells(ie, "C") = PBI.Cells(ip, "K")
    e.Cells(ie, "D") = sz.Cells(52, "X")
    Dim lisz As Integer
    lisz = WorksheetFunction.CountA(sz.Columns(26))
    If IsError(Application.VLookup(e.Cells(ie, "C"), Range(sz.Cells(1, 26), sz.Cells(lisz, 26)), 1, 0)) = True Then
    sz.Cells(WorksheetFunction.CountA(sz.Columns(26)) + 1, 26) = e.Cells(ie, "C")
    Else
    End If
    
    
    'sprawdz czy ju� jest
    
    
    walidacja = False
    Else
    End If
    
    
    
    '---v50
    PBI.Cells(ip, "M") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 9, 0)
    PBI.Cells(ip, "N") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 10, 0)
    PBI.Cells(ip, "P") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 11, 0)
    
    PBI.Cells(ip, "S") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 12, 0)
    PBI.Cells(ip, "T") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 13, 0)
    PBI.Cells(ip, "U") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 14, 0)
    PBI.Cells(ip, "V") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 15, 0)
    PBI.Cells(ip, "W") = Application.VLookup(PBI.Cells(ip, "B"), Range(jira.Cells(2, 1), jira.Cells(ip_jira, 18)), 17, 0)
     If Len(PBI.Cells(ip, "W")) < 6 Then
    PBI.Cells(ip, "W") = " "
    Else
    End If
    End If
    
    
    '---v50
        If PBI_SRC.Cells(ip, "F") < PBI_SRC.Cells(ip, "I") Then
        PBI.Cells(ip, "Q") = "Tak"
        PBI.Cells(ip, "Q").Interior.Color = RGB(101, 217, 101)
        Else
        PBI.Cells(ip, "Q") = "Nie"
        PBI.Cells(ip, "Q").Interior.Color = RGB(222, 85, 74)
        End If
    
        If PBI.Cells(ip, "P") > Now() Then
        PBI.Cells(ip, "R") = "Tak"
        PBI.Cells(ip, "R").Interior.Color = RGB(101, 217, 101)
        Else
        PBI.Cells(ip, "R") = "Nie"
        PBI.Cells(ip, "R").Interior.Color = RGB(222, 85, 74)
        End If
        
        If PBI.Cells(ip, "O") = "Tak" Then
        PBI.Cells(ip, "O").Interior.Color = RGB(101, 217, 101)
        Else
        PBI.Cells(ip, "O").Interior.Color = RGB(222, 85, 74)
        End If
        
        If PBI.Cells(ip, "D") = "Tak" Then
        Union(Range(PBI.Cells(ip, "P"), PBI.Cells(ip, "P")), Range(PBI.Cells(ip, "S"), PBI.Cells(ip, "W")), Range(PBI.Cells(ip, "A"), PBI.Cells(ip, "N"))).Font.ColorIndex = 3 'v50
        
        Else
        End If
        
        'kolorowanie stabilizacji
        If PBI.Cells(ip, "E") = Worksheets("GO").Cells(8, "N") Then
        Union(Range(PBI.Cells(ip, "P"), PBI.Cells(ip, "P")), Range(PBI.Cells(ip, "S"), PBI.Cells(ip, "W")), Range(PBI.Cells(ip, "A"), PBI.Cells(ip, "N"))).Interior.Color = RGB(225, 240, 255) 'v50
        Else
            If (Worksheets("GO").Cells(10, "N") <> "" And PBI.Cells(ip, "E") = Worksheets("GO").Cells(10, "N")) Then
            Union(Range(PBI.Cells(ip, "P"), PBI.Cells(ip, "P")), Range(PBI.Cells(ip, "S"), PBI.Cells(ip, "W")), Range(PBI.Cells(ip, "A"), PBI.Cells(ip, "N"))).Interior.Color = RGB(225, 250, 240) 'v50
            Else
            PBI.Cells(ip, "E") = "Utrzymanie"
            End If
        End If
        
        
    'End If
'ip = ip + 1

Next

Union(PBI.Columns("D:D"), PBI.Columns("F:F"), PBI.Columns("J:J"), PBI.Columns("H:H"), PBI.Columns("T:U"), PBI.Columns("L:R"), PBI.Columns("W:W")).HorizontalAlignment = xlCenter 'v50
PBI.Columns("A:W").VerticalAlignment = xlCenter 'v50

Range(PBI.Rows(2), PBI.Rows(WorksheetFunction.CountA(PBI.Columns(1)))).RowHeight = 20
PBI.Columns("V:V").NumberFormat = "@"

Range(PBI.Rows(2), PBI.Rows(WorksheetFunction.CountA(PBI.Columns(1)))).Font.Name = "Calibri"
Range(PBI.Rows(2), PBI.Rows(WorksheetFunction.CountA(PBI.Columns(1)))).Font.Size = 11

PBI.Activate
ActiveWindow.ScrollColumn = 1
ActiveWindow.ScrollRow = 1
PBI.Cells(1, 1).Activate

If walidacja = False Then
Worksheets("Errors").Activate
Else
End If

End Sub

