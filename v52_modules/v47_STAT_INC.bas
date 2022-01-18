Attribute VB_Name = "STAT_INC"

Sub dane_zrodlo_Klikniecie() 'uzupelnienie z tabeli OSS_MIX
Dim A As String
Dim B As String

Dim csv As Worksheet
Set csv = Worksheets("CSV")

Call Daily

Call nowy_wiersz_auto

Call add_data

Call bilans

'Call pivot

'---czyszczenie dla INC


If (WorksheetFunction.CountA(csv.Columns(1)) Or WorksheetFunction.CountA(csv.Columns(5))) > 1 Then ' tu dodaæ warunek
    If Worksheets("GO").Cells(13, "O") = "Tak" And (WorksheetFunction.CountA(Range(csv.Cells(1, "E"), csv.Cells(2, "I"))) < 4) Then
    'dodatkowe
    Range(csv.Cells(2, "E"), csv.Cells(Application.WorksheetFunction.max(WorksheetFunction.CountA(csv.Columns("E")), WorksheetFunction.CountA(csv.Columns("G")), WorksheetFunction.CountA(csv.Columns("I"))), "I")).ClearContents
    End If
'z dziennego
    If WorksheetFunction.CountA(csv.Columns(1)) > 1 Then
    Range(csv.Cells(2, "A"), csv.Cells(Application.WorksheetFunction.max(WorksheetFunction.CountA(csv.Columns("A")), WorksheetFunction.CountA(csv.Columns("C")), WorksheetFunction.CountA(csv.Columns("E")), WorksheetFunction.CountA(csv.Columns("G")), WorksheetFunction.CountA(csv.Columns("I"))), "D")).ClearContents
    Else
    End If
Else
End If

End Sub


Sub nowy_wiersz_auto() 'OSS_MIX --DMP

Dim lp As Integer
Dim arkusz As String

Dim sz As Worksheet, temp As Worksheet, ALL As Worksheet
Dim mc_short As String, mc_long As String, tg As String

Set ALL = Worksheets("VC2")
Set sz = Worksheets("Konfiguracja")

R = WorksheetFunction.CountA(ALL.Columns(1)) 'ostatni wypelniony wiersz w VC2

    If ALL.Cells(R, 4) <> Worksheets("GO").Cells(8, "J") = True Then ' nowy wiersz
    'R=R
    Else
    R = R - 1 'czyli wiersz juz jest
    End If
    
    If Len(Month(Worksheets("GO").Cells(8, "J"))) = 2 Then
    mc_short = Year(Worksheets("GO").Cells(8, "J")) & "-" & Month(Worksheets("GO").Cells(8, "J"))
    Else
    mc_short = Year(Worksheets("GO").Cells(8, "J")) & "-0" & Month(Worksheets("GO").Cells(8, "J"))
    End If
    
    mc_long = Application.WorksheetFunction.VLookup(Month(Worksheets("GO").Cells(8, "J")), sz.Range("D3:E14"), 2, 0)
    tg = Application.WorksheetFunction.WeekNum(Worksheets("GO").Cells(8, "J"))
'---DMP START
For lp = 3 To 38 'v20210929
'---DMP STOP
    
    arkusz = sz.Cells(lp, "N")
    Set temp = Worksheets(arkusz)

    rmod = sz.Cells(lp, "O")
    temp.Cells(R + rmod, 4) = Worksheets("GO").Cells(8, "J")
    temp.Cells(R + rmod, 1) = Year(temp.Cells(R + rmod, 4))
    temp.Cells(R + rmod, 5) = tg
    temp.Cells(R + rmod, 2) = mc_long
    temp.Cells(R + rmod, 3) = mc_short

    temp.Cells(R + rmod, 1).NumberFormat = "General"
    temp.Cells(R + rmod, 3).NumberFormat = "0"
    temp.Cells(R + rmod, 4).NumberFormat = "dd.mm.yyyy"
    temp.Cells(R + rmod, 5).NumberFormat = "0"
    Range(temp.Cells(R + rmod, 1), temp.Cells(R + rmod, 5)).Interior.Color = RGB(sz.Cells(lp, "P"), sz.Cells(lp, "Q"), sz.Cells(lp, "R"))
    Range(temp.Cells(R + rmod, 1), temp.Cells(R + rmod, 5)).Font.Size = 9
    Range(temp.Cells(R + rmod, 1), temp.Cells(R + rmod, 5)).Font.Name = Calibri
    Range(temp.Cells(R + rmod, 1), temp.Cells(R + rmod, 5)).HorizontalAlignment = xlCenter
    Range(temp.Cells(R + rmod, 1), temp.Cells(R + rmod, 5)).VerticalAlignment = xlCenter
    Union(temp.Cells(R + rmod, 5), temp.Cells(R + rmod, 10), temp.Cells(R + rmod, 13)).Borders(xlEdgeRight).Weight = xlMedium
    
Next


End Sub

Sub count_groups() 'OSS_MIX -- DMP

Dim SRC As Worksheet, csv As Worksheet
Dim lp As Integer

'--- mo¿e byc zbedne - sprawdz 'v20210929
'Dim IDF As Worksheet
'Set IDF = Worksheets("IDF")
'---

Dim max As Integer

Set SRC = Worksheets("STAT_SRC")
Set csv = Worksheets("CSV")

'--DMP START
Range(SRC.Cells(3, 2), SRC.Cells(37, 6)).ClearContents 'DMP 'v20210929
'--DMP STOP

max = Application.WorksheetFunction.max(WorksheetFunction.CountA(csv.Columns("A")), WorksheetFunction.CountA(csv.Columns("C")), WorksheetFunction.CountA(csv.Columns("E")), WorksheetFunction.CountA(csv.Columns("G")), WorksheetFunction.CountA(csv.Columns("I")))

'--DMP modyfikacja zapytania je¿eli inna grupa
'v20210929
SRC.Cells(3, 2) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 1)), (csv.Cells(max, 1))), "VC_OSS_FIXED_*") + _
Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 1)), (csv.Cells(max, 1))), "MIESZKO VENDOR") _
+ Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 1)), (csv.Cells(max, 1))), "APLIKACJE_ATRIUM")

SRC.Cells(3, 3) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 3)), (csv.Cells(max, 3))), "VC_OSS_FIXED_*") + _
Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 3)), (csv.Cells(max, 3))), "MIESZKO VENDOR") + _
Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 3)), (csv.Cells(max, 3))), "APLIKACJE_ATRIUM")

SRC.Cells(3, 4) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 5)), (csv.Cells(max, 5))), "VC_OSS_FIXED_*") _
+ Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 5)), (csv.Cells(max, 5))), "MIESZKO VENDOR") _
+ Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 5)), (csv.Cells(max, 5))), "APLIKACJE_ATRIUM")

SRC.Cells(3, 5) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 7)), (csv.Cells(max, 7))), "VC_OSS_FIXED_*") _
+ Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 7)), (csv.Cells(max, 7))), "MIESZKO VENDOR") _
+ Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 7)), (csv.Cells(max, 7))), "APLIKACJE_ATRIUM")

SRC.Cells(3, 6) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 9)), (csv.Cells(max, 9))), "VC_OSS_FIXED_*") _
+ Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 9)), (csv.Cells(max, 9))), "MIESZKO VENDOR") _
+ Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 9)), (csv.Cells(max, 9))), "APLIKACJE_ATRIUM")
 
For lp = 4 To 35 ' DMP 'v20210929

SRC.Cells(lp, 2) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 1)), (csv.Cells(max, 1))), SRC.Cells(lp, "A"))
SRC.Cells(lp, 3) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 3)), (csv.Cells(max, 3))), SRC.Cells(lp, "A"))
SRC.Cells(lp, 4) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 5)), (csv.Cells(max, 5))), SRC.Cells(lp, "A"))
SRC.Cells(lp, 5) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 7)), (csv.Cells(max, 7))), SRC.Cells(lp, "A"))
SRC.Cells(lp, 6) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 9)), (csv.Cells(max, 9))), SRC.Cells(lp, "A"))

'-------- tu wstawiamy ew dodatkowe produkty

Next
'--DMP

'tu lp ma 36

SRC.Cells(lp, 2) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 1)), (csv.Cells(max, 1))), "VC_TP_OSS*")
SRC.Cells(lp, 3) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 3)), (csv.Cells(max, 3))), "VC_TP_OSS*")
SRC.Cells(lp, 4) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 5)), (csv.Cells(max, 5))), "VC_TP_OSS*")
SRC.Cells(lp, 5) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 7)), (csv.Cells(max, 7))), "VC_TP_OSS*")
SRC.Cells(lp, 6) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 9)), (csv.Cells(max, 9))), "VC_TP_OSS*")

lp = lp + 1
'-----------------------------------------------------------------------
SRC.Cells(lp, 2) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 1)), (csv.Cells(max, 1))), "VC_OSS_FIXED_DZIA£ANIA_WSPIERAJ¥CE")
SRC.Cells(lp, 3) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 3)), (csv.Cells(max, 3))), "VC_OSS_FIXED_DZIA£ANIA_WSPIERAJ¥CE")
SRC.Cells(lp, 4) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 5)), (csv.Cells(max, 5))), "VC_OSS_FIXED_DZIA£ANIA_WSPIERAJ¥CE")
SRC.Cells(lp, 5) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 7)), (csv.Cells(max, 7))), "VC_OSS_FIXED_DZIA£ANIA_WSPIERAJ¥CE")
SRC.Cells(lp, 6) = Application.WorksheetFunction.CountIf(Range((csv.Cells(2, 9)), (csv.Cells(max, 9))), "VC_OSS_FIXED_DZIA£ANIA_WSPIERAJ¥CE")

Call Daily

End Sub



Sub add_data() 'OSS_MIX

Dim SRC As Worksheet, ALL As Worksheet, sz As Worksheet
Dim lp As Integer, lps As Integer

Set ALL = Worksheets("VC2")
Set sz = Worksheets("Konfiguracja")
Set SRC = Worksheets("STAT_SRC")

R = WorksheetFunction.CountA(ALL.Columns(1)) - 1

On Error Resume Next

''ALL osobno
'--DMP START 'v20210929
rmod = sz.Cells(38, "O")
ALL.Cells(R + rmod, 6) = SRC.Cells(3, 2) - SRC.Cells(36, 2)
ALL.Cells(R + rmod, 7) = SRC.Cells(3, 3) - SRC.Cells(36, 3)
ALL.Cells(R + rmod - 1, 8) = SRC.Cells(3, 4) - SRC.Cells(36, 4)
ALL.Cells(R + rmod - 1, 9) = SRC.Cells(3, 5) - SRC.Cells(36, 5)
ALL.Cells(R + rmod - 1, 10) = SRC.Cells(3, 6) - SRC.Cells(36, 6)
'--DMP STOP
ALL.Cells(R + rmod, 7).NumberFormat = "0"
ALL.Cells(R + rmod, 13).NumberFormat = "0"
ALL.Cells(R + rmod - 1, 12).NumberFormat = "0"
ALL.Cells(R + rmod, 11) = "-"
ALL.Cells(R + rmod, 11) = ALL.Cells(R + rmod, 7) / ALL.Cells(R + rmod, 6)
ALL.Cells(R + rmod - 1, 12) = ALL.Cells(R + rmod - 1, 8) - ALL.Cells(R + rmod - 1, 9)

    If ALL.Cells(R + rmod - 1, 12) = 0 Then
    ALL.Cells(R + rmod - 1, 12).Font.ColorIndex = 1
    Else
        If ALL.Cells(R + rmod - 1, 12) > 0 Then
        ALL.Cells(R + rmod - 1, 12).Font.ColorIndex = 4
        Else
        ALL.Cells(R + rmod - 1, 12).Font.ColorIndex = 3
        End If
    End If

ALL.Cells(R + rmod, 13) = ALL.Cells(R + rmod, 7) - ALL.Cells(R + rmod - 1, 7)
    If ALL.Cells(R + rmod, 13) = 0 Then
    ALL.Cells(R + rmod, 13).Font.ColorIndex = 1
    Else
        If ALL.Cells(R + rmod, 13) < 0 Then
        ALL.Cells(R + rmod, 13).Font.ColorIndex = 4
        Else
        ALL.Cells(R + rmod, 13).Font.ColorIndex = 3
        End If
    End If


    Union(Range(ALL.Cells(R + rmod - 1, 6), ALL.Cells(R + rmod, 10)), ALL.Cells(R + rmod, "N"), ALL.Cells(R + rmod, "M")).NumberFormat = "0"
    Range(ALL.Cells(R + rmod - 1, 6), ALL.Cells(R + rmod, 14)).Font.Size = 9
    Range(ALL.Cells(R + rmod - 1, 6), ALL.Cells(R + rmod, 14)).Font.Name = Calibri
    ALL.Cells(R + rmod, 11).NumberFormat = "0%"
    Range(ALL.Cells(R + rmod - 1, 1), ALL.Cells(R + rmod, 14)).HorizontalAlignment = xlCenter
    Range(ALL.Cells(R + rmod - 1, 1), ALL.Cells(R + rmod, 14)).VerticalAlignment = xlCenter
 
lps = 3
'--DMP START
For lp = 3 To 37 ' 'v20210929
'--DMP STOP
arkusz = sz.Cells(lp, "N")
Set temp = Worksheets(arkusz)
rmod = sz.Cells(lp, "O")


temp.Cells(R + rmod, 6) = SRC.Cells(lps, 2)
temp.Cells(R + rmod, 7) = SRC.Cells(lps, 3)
temp.Cells(R + rmod, 7).NumberFormat = "0" 'test
temp.Cells(R + rmod, 13).NumberFormat = "0" 'test
temp.Cells(R + rmod - 1, 12).NumberFormat = "0" 'test
temp.Cells(R + rmod - 1, 8) = SRC.Cells(lps, 4) '4 obsluzone
temp.Cells(R + rmod - 1, 9) = SRC.Cells(lps, 5) '5 zgloszone
temp.Cells(R + rmod - 1, 10) = SRC.Cells(lps, 6)
temp.Cells(R + rmod, 11) = "-"
temp.Cells(R + rmod, 11) = temp.Cells(R + rmod, 7) / temp.Cells(R + rmod, 6)
temp.Cells(R + rmod - 1, 12) = temp.Cells(R + rmod - 1, 8) - temp.Cells(R + rmod - 1, 9)
temp.Cells(R + rmod, 13) = temp.Cells(R + rmod, 7) - temp.Cells(R + rmod - 1, 7)


    If temp.Cells(R + rmod - 1, 12) = 0 Then
    temp.Cells(R + rmod - 1, 12).Font.ColorIndex = 1
    Else
        If temp.Cells(R + rmod - 1, 12) > 0 Then
        temp.Cells(R + rmod - 1, 12).Font.ColorIndex = 4
        Else
        temp.Cells(R + rmod - 1, 12).Font.ColorIndex = 3
        End If
    End If

    If temp.Cells(R + rmod, 13) = 0 Then
    temp.Cells(R + rmod, 13).Font.ColorIndex = 1
    Else
        If temp.Cells(R + rmod, 13) < 0 Then
        temp.Cells(R + rmod, 13).Font.ColorIndex = 4
        Else
        temp.Cells(R + rmod, 13).Font.ColorIndex = 3
        End If
    End If
    
    If arkusz = "VC1VC2" Then
    temp.Cells(R + rmod, "N") = 680
    temp.Cells(R + rmod, "N").Interior.Color = RGB(250, 191, 143)
    temp.Cells(R + rmod, "N").Borders(xlEdgeRight).Weight = xlMedium

    End If
    


    Union(Range(temp.Cells(R + rmod - 1, 6), temp.Cells(R + rmod, 10)), temp.Cells(R + rmod, "N"), temp.Cells(R + rmod, "M")).NumberFormat = "0"
    Range(temp.Cells(R + rmod - 1, 6), temp.Cells(R + rmod, 14)).Font.Size = 9
    Range(temp.Cells(R + rmod - 1, 6), temp.Cells(R + rmod, 14)).Font.Name = Calibri
    temp.Cells(R + rmod, 11).NumberFormat = "0%"
    Range(temp.Cells(R + rmod - 1, 1), temp.Cells(R + rmod, 14)).HorizontalAlignment = xlCenter
    Range(temp.Cells(R + rmod - 1, 1), temp.Cells(R + rmod, 14)).VerticalAlignment = xlCenter

lps = lps + 1

Next

End Sub



