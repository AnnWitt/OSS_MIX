Attribute VB_Name = "RAPORT_PBI"

Sub PBI()
Dim PBI_SRC As Worksheet, PBI As Worksheet, Jira As Worksheet, sz As Worksheet
Dim ip_jira As Integer, ip As Integer



Set PBI_SRC = Worksheets("PBI_Remedy")
Set PBI = Worksheets("Raport PBI")
Set Jira = Worksheets("JIRA OSS")
Set sz = Worksheets("Konfiguracja")

ip_jira = WorksheetFunction.CountA(Jira.Columns(1)) 'zamkniêcie zakresu

For ip = 2 To WorksheetFunction.CountA(PBI_SRC.Columns(1))
'dane z Remedy - w tym przypadku moze byæ ten sam licznik, przy jira nie
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
Union(PBI.Columns("M:N"), PBI.Columns("P:P"), PBI.Columns("T:T")).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    If PBI_SRC.Cells(ip, "E") <> "" Then
    PBI.Cells(ip, "O") = "Tak"
    Else
    PBI.Cells(ip, "O") = "Nie"
    End If
'Dane z jira
    If IsError(Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 3, 0)) = True Then
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
    Else
    PBI.Cells(ip, "A") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 2, 0)
    PBI.Cells(ip, "C") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 3, 0)
    PBI.Cells(ip, "D") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 4, 0)
    PBI.Cells(ip, "H") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 5, 0)
    PBI.Cells(ip, "I") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 6, 0)
    
    PBI.Cells(ip, "J") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 7, 0) 'left 3 lub 2 czyli ew wyciac spacje
    PBI.Cells(ip, "J") = Trim(Left(Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 7, 0), 3))
    
    PBI.Cells(ip, "K") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 7, 0)
    
    If (PBI.Cells(ip, "K") = "nPPKuser") Or (PBI.Cells(ip, "K") = "OCLuser") Then
    PBI.Cells(ip, "J") = "DEV"
    Else
    End If
    
    If (PBI.Cells(ip, "K") = "Nieprzydzielone") Then
    PBI.Cells(ip, "J") = "#ND"
    Else
    End If
    
    
    
    PBI.Cells(ip, "M") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 9, 0)
    PBI.Cells(ip, "N") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 10, 0)
    PBI.Cells(ip, "P") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 11, 0)
    
    PBI.Cells(ip, "S") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 12, 0)
    PBI.Cells(ip, "T") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 13, 0)
    PBI.Cells(ip, "U") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 14, 0)
    PBI.Cells(ip, "V") = Application.VLookup(PBI.Cells(ip, "B"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 16)), 15, 0)
    
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
        Union(Range(PBI.Cells(ip, "P"), PBI.Cells(ip, "P")), Range(PBI.Cells(ip, "S"), PBI.Cells(ip, "V")), Range(PBI.Cells(ip, "A"), PBI.Cells(ip, "N"))).Font.ColorIndex = 3
        
        Else
        End If
        
        'kolorowanie stabilizacji
        If PBI.Cells(ip, "E") = Worksheets("GO").Cells(8, "N") Then
        Union(Range(PBI.Cells(ip, "P"), PBI.Cells(ip, "P")), Range(PBI.Cells(ip, "S"), PBI.Cells(ip, "V")), Range(PBI.Cells(ip, "A"), PBI.Cells(ip, "N"))).Interior.Color = RGB(225, 240, 255)
        Else
            If (Worksheets("GO").Cells(10, "N") <> "" And PBI.Cells(ip, "E") = Worksheets("GO").Cells(10, "N")) Then
            Union(Range(PBI.Cells(ip, "P"), PBI.Cells(ip, "P")), Range(PBI.Cells(ip, "S"), PBI.Cells(ip, "V")), Range(PBI.Cells(ip, "A"), PBI.Cells(ip, "N"))).Interior.Color = RGB(225, 250, 240)
            Else
            PBI.Cells(ip, "E") = "Utrzymanie"
            End If
        End If
        
        
    End If
'ip = ip + 1

Next

Union(PBI.Columns("D:D"), PBI.Columns("F:F"), PBI.Columns("J:J"), PBI.Columns("H:H"), PBI.Columns("T:U"), PBI.Columns("L:R")).HorizontalAlignment = xlCenter
PBI.Columns("A:V").VerticalAlignment = xlCenter

Range(PBI.Rows(2), PBI.Rows(WorksheetFunction.CountA(PBI.Columns(1)))).RowHeight = 20


Range(PBI.Rows(2), PBI.Rows(WorksheetFunction.CountA(PBI.Columns(1)))).Font.Name = "Calibri"
Range(PBI.Rows(2), PBI.Rows(WorksheetFunction.CountA(PBI.Columns(1)))).Font.Size = 11

PBI.Activate
ActiveWindow.ScrollColumn = 1
ActiveWindow.ScrollRow = 1
PBI.Cells(1, 1).Activate

End Sub

