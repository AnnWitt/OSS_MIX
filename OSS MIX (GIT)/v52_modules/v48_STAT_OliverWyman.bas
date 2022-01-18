Attribute VB_Name = "STAT_OliverWyman"
Sub Oliver_Wyman() 'OSS MIX -> DMP
Dim ow As Worksheet, sz As Worksheet, temp As Worksheet, ALL As Worksheet
Dim o As Integer, v As Integer, new_row As Boolean, same_month As Boolean, x1 As Integer, x2 As Integer, x3 As Integer, x4 As Integer, lp As Integer, arkusz As String

Set ALL = Worksheets("VC2")
Set ow = Worksheets("Oliver Wyman - INC")
Set sz = Worksheets("Konfiguracja")

o = WorksheetFunction.CountA(ow.Columns(1)) + 2 'pierwszy wolny wiersz

v = WorksheetFunction.CountA(ALL.Columns(1)) 'ostatni wype³niony wiersz
same_month = False
new_row = False

If ALL.Cells(v, "A") <> ALL.Cells(v - 1, "A") Then
        If (ALL.Cells(v - 1, "A") <> ow.Cells(o - 1, "B")) Or (ALL.Cells(v - 1, "B") <> ow.Cells(o - 1, "C")) Then
        new_row = True
        Else
        End If
Else

    If ALL.Cells(v, "B") = ow.Cells(o - 1, "C") Then 'ten sam miesiac
        same_month = True
        Else
        new_row = True
        End If

End If

If (new_row = True Or same_month = True) Then
    If same_month = True Then
    o = o - 1
    Else
    End If
    ow.Cells(o, "B") = ALL.Cells(v, "A")
    ow.Cells(o, "C") = ALL.Cells(v, "B")
    ow.Cells(o, "A") = ow.Cells(o - 1, "A") + 1
    ow.Cells(o, "D") = ow.Cells(o - 1, "D")
    Range(ow.Cells(o, "A"), ow.Cells(o, "C")).Interior.Color = RGB(49, 134, 155)
    Range(ow.Cells(o, "A"), ow.Cells(o, "C")).Font.Color = RGB(255, 255, 255)
    '---DMP START 'v20210929
    '---DMP START 'v20210929
    Range(ow.Cells(o, "A"), ow.Cells(o, "EF")).Font.Size = 9
    Range(ow.Cells(o, "A"), ow.Cells(o, "EF")).Font.Name = Calibri
    Range(ow.Cells(o, "A"), ow.Cells(o, "EF")).NumberFormat = "general"
    Range(ow.Cells(o, "A"), ow.Cells(o, "EF")).HorizontalAlignment = xlCenter
    Range(ow.Cells(o, "A"), ow.Cells(o, "EF")).VerticalAlignment = xlCenter
    '---DMP STOP


    ow.Cells(o, "D").Interior.Color = RGB(166, 166, 166)


    x1 = 9 'i potem +4
    x2 = 10
    x3 = 11
    x4 = 12
    '---DMP START
    For lp = 4 To 35 ' do   DMP v20210929
    '---DMP STOP
    arkusz = sz.Cells(lp, "N")
    Set temp = Worksheets(arkusz)
    'niby dla idf mogloby byc v - 1626 ale to byc trzeba wyjsc poza pêtle, nic mu nie bêdzie jak posumuje tez po pustym zakresie
    ow.Cells(o, x1) = WorksheetFunction.SumIfs(Range(temp.Cells(2, "I"), temp.Cells(v, "I")), Range(temp.Cells(2, "A"), temp.Cells(v, "A")), _
    ow.Cells(o, "B"), Range(temp.Cells(2, "B"), temp.Cells(v, "B")), ow.Cells(o, "C"))

    'v20210929
    ow.Cells(o, x2) = ow.Cells(o, x1) / WorksheetFunction.VLookup(ow.Cells(1, x1), Range(sz.Cells(3, "A"), sz.Cells(35, "B")), 2, 0) 'DMP |+ dopisz w konfiguracji
    If ow.Cells(o, x1) = 0 Then
    ow.Cells(o, x3) = "-"
    Else
    ow.Cells(o, x3) = ow.Cells(o, x2) / ow.Cells(o, "D")
    End If
    If WorksheetFunction.IsNumber(ow.Cells(o - 1, x2)) = False Then
    ow.Cells(o, x4) = "-"
    Else
        If ow.Cells(o - 1, x2) = 0 Then
        ow.Cells(o, x4) = "-"
        Else
            If ow.Cells(o, x2) > 0 Then
            ow.Cells(o, x4) = ow.Cells(o, x2) / ow.Cells(o - 1, x2) - 1
            Else
            ow.Cells(o, x4) = 0
            End If
        End If
    End If

    ow.Cells(o, x1).Interior.Color = RGB(192, 80, 77)

    Range(ow.Cells(3, x4), ow.Cells(o, x4)).Borders(xlEdgeRight).Weight = xlMedium
    ow.Cells(o, x3).NumberFormat = "0.00%"
    ow.Cells(o, x4).NumberFormat = "0.00%"

    x1 = x1 + 4 'i potem +4
    x2 = x2 + 4
    x3 = x3 + 4
    x4 = x4 + 4
    Next

    'osobny dla VC2

     ow.Cells(o, 5) = WorksheetFunction.SumIfs(Range(ALL.Cells(2, "I"), ALL.Cells(v, "I")), Range(ALL.Cells(2, "A"), _
     ALL.Cells(v, "A")), ow.Cells(o, "B"), Range(ALL.Cells(2, "B"), ALL.Cells(v, "B")), ow.Cells(o, "C"))

    ow.Cells(o, 6) = ow.Cells(o, 5) / WorksheetFunction.VLookup(ow.Cells(1, 5), Range(sz.Cells(3, "A"), sz.Cells(35, "B")), 2, 0) 'DMP zakres w sz
    If ow.Cells(o, 5) = 0 Then
    ow.Cells(o, 7) = "-"
    Else
    ow.Cells(o, 7) = ow.Cells(o, 6) / ow.Cells(o, "D")
    End If
    If WorksheetFunction.IsNumber(ow.Cells(o - 1, 6)) = False Then
    ow.Cells(o, 8) = "-"
    Else
        If ow.Cells(o - 1, 6) = 0 Then
        ow.Cells(o, 8) = "-"
        Else
            If ow.Cells(o, 6) > 0 Then
            ow.Cells(o, 8) = ow.Cells(o, 6) / ow.Cells(o - 1, 6) - 1
            Else
            ow.Cells(o, 8) = 0
            End If
        End If
    End If

    ow.Cells(o, 5).Interior.Color = RGB(192, 80, 77)

    Range(ow.Cells(3, 8), ow.Cells(o, 8)).Borders(xlEdgeRight).Weight = xlMedium
    ow.Cells(o, 7).NumberFormat = "0.00%"
    ow.Cells(o, 8).NumberFormat = "0.00%"


Else
End If


End Sub

