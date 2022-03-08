Attribute VB_Name = "RAPORT_INC"

Sub inc() '---MOZLIWE DMP JESLI NP GOLD 24/7

Dim INC_SRC As Worksheet, Jira As Worksheet, inc As Worksheet, csv As Worksheet
Dim ip_jira As Integer, ii_src As Integer, ii As Integer
Dim r1 As Integer, r2 As Integer, o1 As Integer, o2 As Integer, g1 As Integer, g2 As Integer

Set INC_SRC = Worksheets("INC_Remedy")
Set inc = Worksheets("Raport INC")
Set Jira = Worksheets("JIRA OSS")
Set csv = Worksheets("CSV")

Range(inc.Cells(1, "T"), inc.Cells(2, "U")) = "x" 'v50
inc.Columns("U").NumberFormat = "0" 'v50
'ii = 3

ip_jira = WorksheetFunction.CountA(Jira.Columns(1))

On Error Resume Next

For ii_src = 2 To WorksheetFunction.CountA(INC_SRC.Columns(1))
'dane z Remedy - w tym przypadku moze byæ ten sam licznik, przy jira nie

'ze zrzutu z Remedy
inc.Cells(ii_src, "A") = INC_SRC.Cells(ii_src, "A")
inc.Cells(ii_src, "B") = INC_SRC.Cells(ii_src, "B")
inc.Cells(ii_src, "C") = INC_SRC.Cells(ii_src, "C")
inc.Cells(ii_src, "D") = INC_SRC.Cells(ii_src, "D")
inc.Cells(ii_src, "E") = INC_SRC.Cells(ii_src, "E")
inc.Cells(ii_src, "F") = INC_SRC.Cells(ii_src, "F")
inc.Cells(ii_src, "I") = INC_SRC.Cells(ii_src, "G")
inc.Cells(ii_src, "J") = INC_SRC.Cells(ii_src, "H")



inc.Cells(ii_src, "P") = INC_SRC.Cells(ii_src, "I")
inc.Cells(ii_src, "Q") = INC_SRC.Cells(ii_src, "J")
inc.Cells(ii_src, "R") = INC_SRC.Cells(ii_src, "L")



If IsError(Application.VLookup(inc.Cells(ii_src, "A"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 2, 0)) = True Then 'v50
    If IsError(Application.VLookup(inc.Cells(ii_src, "E"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 2, 0)) = True Then  ' (v50) je¿eli inca nie ma w jira to niech pokaze ew PBI
    inc.Cells(ii_src, "H") = "-"
    inc.Cells(ii_src, "G") = "-"
    inc.Cells(ii_src, "L") = "-"
    inc.Cells(ii_src, "M") = "-"
    inc.Cells(ii_src, "N") = "-"
    inc.Cells(ii_src, "O") = "-"
    inc.Cells(ii_src, "S") = "-"
    Else 'v50
    inc.Cells(ii_src, "H") = Application.VLookup(inc.Cells(ii_src, "E"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 2, 0)
    inc.Cells(ii_src, "G") = Application.VLookup(inc.Cells(ii_src, "E"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 7, 0)
    inc.Cells(ii_src, "L") = Application.VLookup(inc.Cells(ii_src, "E"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 12, 0)
    inc.Cells(ii_src, "M") = Application.VLookup(inc.Cells(ii_src, "E"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 14, 0)
    inc.Cells(ii_src, "N") = Application.VLookup(inc.Cells(ii_src, "E"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 15, 0)
    inc.Cells(ii_src, "O") = Application.VLookup(inc.Cells(ii_src, "E"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 5, 0)
    inc.Cells(ii_src, "S") = Application.VLookup(inc.Cells(ii_src, "E"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 17, 0)
     If Len(inc.Cells(ii_src, "S")) < 6 Then
    inc.Cells(ii_src, "S") = ""
    Else
    End If
    End If
Else 'v50
inc.Cells(ii_src, "H") = Application.VLookup(inc.Cells(ii_src, "A"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 2, 0)
inc.Cells(ii_src, "G") = Application.VLookup(inc.Cells(ii_src, "A"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 7, 0)
inc.Cells(ii_src, "L") = Application.VLookup(inc.Cells(ii_src, "A"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 12, 0)
inc.Cells(ii_src, "M") = Application.VLookup(inc.Cells(ii_src, "A"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 14, 0)
inc.Cells(ii_src, "N") = Application.VLookup(inc.Cells(ii_src, "A"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 15, 0)
inc.Cells(ii_src, "O") = Application.VLookup(inc.Cells(ii_src, "A"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 5, 0)
inc.Cells(ii_src, "S") = Application.VLookup(inc.Cells(ii_src, "A"), Range(Jira.Cells(2, 1), Jira.Cells(ip_jira, 18)), 17, 0)
End If

If inc.Cells(ii_src, "N") = "" Then
inc.Cells(ii_src, "N") = "-"
Else
End If

If inc.Cells(ii_src, "O") = "" Then
inc.Cells(ii_src, "O") = "-"
Else
End If



' uwaga dla remedy dub !
inc.Cells(ii_src, "K").NumberFormat = "[hhh]:mm:ss;@"


    If inc.Cells(ii_src, "B") = "VC_OSS_FIXED_REMEDY-DUB" Then
        If (Hour(inc.Cells(ii_src, "I")) >= 16) = True Or (Hour(inc.Cells(ii_src, "I")) <= 7) = True Then
        'inc_wynik.Cells(rmd, "Z") = "poza oknem - godziny"
        Union(Range(inc.Cells(ii_src, "A"), inc.Cells(ii_src, "J")), Range(inc.Cells(ii_src, "L"), inc.Cells(ii_src, "S"))).Interior.Color = RGB(204, 204, 255) 'v50
        Else
            If (Weekday(inc.Cells(ii_src, "I"), vbMonday) > 5) Then
            'inc_wynik.Cells(rmd, "Z") = "poza oknem - dzien tygodnia"
            Union(Range(inc.Cells(ii_src, "A"), inc.Cells(ii_src, "J")), Range(inc.Cells(ii_src, "L"), inc.Cells(ii_src, "S"))).Interior.Color = RGB(204, 204, 255) 'v50
            Else
           ' inc_wynik.Cells(rmd, "Z") = "w oknie"
            End If
        End If

    '-------------
        If inc.Cells(ii_src, "J") > Now() Then 'terminowe
        inc.Cells(ii_src, "K").Interior.Color = RGB(101, 217, 101)
        inc.Cells(ii_src, "T") = "Green" 'v50
        inc.Cells(ii_src, "U") = (inc.Cells(ii_src, "J") - Now()) 'v50
        'inc.Cells(ii_src, "U") = Int(inc.Cells(ii_src, "J") - Now()) 'v50
        inc.Cells(ii_src, "K") = Int(inc.Cells(ii_src, "J") - Now()) & " dni kal."
        
            If Int(inc.Cells(ii_src, "J") - Now()) <= 1 Then
                inc.Cells(ii_src, "K") = inc.Cells(ii_src, "J") - Now() ' uwaga, jezeli mniej ni¿ doba to Ÿle policzy
                'inc.Cells(ii_src, "K") = Int(inc.Cells(ii_src, "J") - Now()) ' uwaga, jezeli mniej ni¿ doba to Ÿle policzy
                inc.Cells(ii_src, "K").Interior.Color = RGB(255, 204, 0)
                inc.Cells(ii_src, "T") = "Orange" 'v50
                inc.Cells(ii_src, "U") = inc.Cells(ii_src, "J") - Now() 'v50
                'inc.Cells(ii_src, "U") = Int(inc.Cells(ii_src, "J") - Now())
                inc.Cells(ii_src, "K").NumberFormat = "[hhh]:mm:ss;@"
            Else
                If Int(inc.Cells(ii_src, "J") - Now()) <= 3 Then
                inc.Cells(ii_src, "K") = Int(inc.Cells(ii_src, "J") - Now())
                inc.Cells(ii_src, "K").Interior.Color = RGB(255, 204, 0)
                inc.Cells(ii_src, "T") = "Orange" ' v50
                inc.Cells(ii_src, "K").NumberFormat = "0"
                inc.Cells(ii_src, "U") = (inc.Cells(ii_src, "J") - Now()) 'v50
                'inc.Cells(ii_src, "U") = Int(inc.Cells(ii_src, "J") - Now())'v50
                inc.Cells(ii_src, "K") = Int(inc.Cells(ii_src, "J") - Now()) & " dni kal."
                Else
                inc.Cells(ii_src, "K").NumberFormat = "0"
                End If
            End If
    
        Else
        
        inc.Cells(ii_src, "K").Interior.Color = RGB(222, 85, 74)
        inc.Cells(ii_src, "T") = "Red" 'v50
        inc.Cells(ii_src, "U") = (Now() - inc.Cells(ii_src, "J")) 'v50
        'inc.Cells(ii_src, "U") = Int(Now() - inc.Cells(ii_src, "J"))
            If Day(inc.Cells(ii_src, "J")) = Day(Now()) And Month(inc.Cells(ii_src, "J")) = Month(Now()) And Year(inc.Cells(ii_src, "J")) = Year(Now()) Then 'MIESIAC !!!!!!! I ROK
            inc.Cells(ii_src, "K") = "0 dni kal."
            Else
            inc.Cells(ii_src, "K") = Int(Now() - inc.Cells(ii_src, "J")) & " dni kal."
            End If
        inc.Cells(ii_src, "K").NumberFormat = "0"
        End If
    
'---------------pozostale
    Else
    
        If inc.Cells(ii_src, "J") > Now() Then 'terminowe
        inc.Cells(ii_src, "K").Interior.Color = RGB(101, 217, 101)
        inc.Cells(ii_src, "T") = "Green" 'v50
        inc.Cells(ii_src, "U") = (WorksheetFunction.NetworkDays(Now(), inc.Cells(ii_src, "J"))) 'v50
        inc.Cells(ii_src, "K") = (WorksheetFunction.NetworkDays(Now(), inc.Cells(ii_src, "J"))) & " dni rob"
        
            If (WorksheetFunction.NetworkDays(Now(), inc.Cells(ii_src, "J"))) <= 2 Then
                inc.Cells(ii_src, "K") = inc.Cells(ii_src, "J") - Now()
                inc.Cells(ii_src, "K").Interior.Color = RGB(255, 204, 0)
                inc.Cells(ii_src, "T") = "Orange" 'v50
                inc.Cells(ii_src, "U") = (WorksheetFunction.NetworkDays(Now(), inc.Cells(ii_src, "J"))) 'v50
                inc.Cells(ii_src, "K").NumberFormat = "[hhh]:mm:ss;@"
            Else
                If (WorksheetFunction.NetworkDays(Now(), inc.Cells(ii_src, "J"))) <= 3 Then
                inc.Cells(ii_src, "K") = inc.Cells(ii_src, "J") - Now()
                inc.Cells(ii_src, "K").Interior.Color = RGB(255, 204, 0)
                inc.Cells(ii_src, "T") = "Orange" 'v50
                inc.Cells(ii_src, "U") = (WorksheetFunction.NetworkDays(Now(), inc.Cells(ii_src, "J"))) 'v50
                inc.Cells(ii_src, "K").NumberFormat = "0"
                inc.Cells(ii_src, "K") = (WorksheetFunction.NetworkDays(Now(), inc.Cells(ii_src, "J"))) & " dni rob."
                Else
                inc.Cells(ii_src, "K").NumberFormat = "0"
                End If
            End If
    
        Else
        
        inc.Cells(ii_src, "K").Interior.Color = RGB(222, 85, 74)
        inc.Cells(ii_src, "T") = "Red" 'v50
        inc.Cells(ii_src, "U") = (WorksheetFunction.NetworkDays(inc.Cells(ii_src, "J"), Now())) 'v50
            If Day(inc.Cells(ii_src, "J")) = Day(Now()) And Month(inc.Cells(ii_src, "J")) = Month(Now()) And Year(inc.Cells(ii_src, "J")) = Year(Now()) Then
            inc.Cells(ii_src, "K") = "0 dni rob."
            Else
            inc.Cells(ii_src, "K") = (WorksheetFunction.NetworkDays(inc.Cells(ii_src, "J"), Now())) & " dni rob."
            End If
        inc.Cells(ii_src, "K").NumberFormat = "0"
        End If
    End If
    
Next

Union(inc.Columns("I:J"), inc.Columns("S:S")).NumberFormat = "yyyy/mm/dd hh:mm:ss" 'v50
Range(inc.Rows(2), inc.Rows(ii_src)).RowHeight = 15
Range(inc.Rows(2), inc.Rows(ii_src)).Font.Name = "Calibri"
Range(inc.Rows(2), inc.Rows(ii_src)).Font.Size = 11

Union(inc.Columns("H:K"), inc.Columns("O:O"), inc.Columns("C:E"), inc.Columns("S:S")).HorizontalAlignment = xlCenter 'v50

Range(inc.Cells(2, "A"), inc.Cells(WorksheetFunction.CountA(INC_SRC.Columns(1)) + 1, "U")).Sort Key1:=inc.Columns("T"), Order1:=xlDescending, Header:=xlNo 'v50

r1 = 2
g2 = WorksheetFunction.CountA(INC_SRC.Columns(1)) '+ 1
o2 = g2 - Application.WorksheetFunction.CountIf(Range(inc.Cells(3, "T"), inc.Cells(g2, "T")), "Green") '- 1 'v50
g1 = o2 + 1
r2 = o2 - Application.WorksheetFunction.CountIf(Range(inc.Cells(3, "T"), inc.Cells(g2, "T")), "Orange") '+ 1'v50
o1 = r2 + 1



Range(inc.Cells(r1, "A"), inc.Cells(r2, "U")).Sort Key1:=inc.Columns("U"), Order1:=xlDescending, Header:=xlNo 'Red 'v50
'Range(inc.Cells(o1, "A"), inc.Cells(g2, "U")).Sort Key1:=inc.Columns("U"), Order1:=xlAscending, Header:=xlNo 'Orange + green 'v50
Range(inc.Cells(o1, "A"), inc.Cells(g2, "U")).Sort Key1:=inc.Columns("J"), Order1:=xlAscending, Header:=xlNo 'Orange + green'v50

For iv = 2 To WorksheetFunction.CountA(inc.Columns(2))
csv.Cells(iv, "A") = inc.Cells(iv, 2)
    If inc.Cells(iv, "T") = "Red" Then 'v50
        csv.Cells(iv, "C") = inc.Cells(iv, 2)
    Else
    End If
Next
inc.Columns("R:R").NumberFormat = "@"
inc.Columns("T:U").Clear 'v50

inc.Activate
ActiveWindow.ScrollColumn = 1
ActiveWindow.ScrollRow = 1
inc.Cells(1, 1).Activate

End Sub

