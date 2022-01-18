Attribute VB_Name = "STAT_BilansGrup"
Sub bilans() 'OSS_MIX --DMP

Dim BIL As Worksheet, sz As Worksheet, arkusz As String, temp As Worksheet, ALL As Worksheet, VC12 As Worksheet

Dim k1 As Integer, k2 As Integer, k3 As Integer

Set VC12 = Worksheets("VC1VC2")
Set ALL = Worksheets("VC2")
Set BIL = Worksheets("Zestawienie Grup")
Set sz = Worksheets("Konfiguracja")

R = WorksheetFunction.CountA(BIL.Columns(1)) ' ostatni wype³niony

If BIL.Cells(WorksheetFunction.CountA(BIL.Columns(1)), 1) <> ALL.Cells(WorksheetFunction.CountA(ALL.Columns(4)) + 1, "D") Then

BIL.Cells(R + 1, 1).Interior.Color = RGB(146, 204, 220)
BIL.Cells(R + 1, 1).NumberFormat = "dd.mm.yyyy"

Else
R = R - 1 ' dla nadpisu wiersza
End If

BIL.Cells(R + 1, 1) = ALL.Cells((WorksheetFunction.CountA(ALL.Columns(4)) + 1), "D")

k1 = 5
k2 = 6
k3 = 7

'--DMP START
For lp = 4 To 36
'--DMP STOP

arkusz = sz.Cells(lp, "N")
Set temp = Worksheets(arkusz)
rmod = sz.Cells(lp, "O")

    BIL.Cells(R + 1, k1) = temp.Cells(R + rmod, 9)
    BIL.Cells(R + 1, k2) = temp.Cells(R + rmod, 8)
    BIL.Cells(R + 1, k3) = BIL.Cells(R + 1, k2) - BIL.Cells(R + 1, k1)

    If BIL.Cells(R + 1, k3) < 0 Then
    BIL.Cells(R + 1, k3).Font.ColorIndex = 3
    Else
    BIL.Cells(R + 1, k3).Font.ColorIndex = 4
    End If
    
    BIL.Cells(R + 1, k3).Interior.Color = RGB(242, 242, 242)
    BIL.Cells(R + 1, k3).Borders(xlEdgeRight).Weight = xlThin

k1 = k1 + 3
k2 = k2 + 3
k3 = k3 + 3
Next

    
    rmod = sz.Cells(3, "O")
    BIL.Cells(R + 1, 2) = VC12.Cells(R + rmod, 9)
    BIL.Cells(R + 1, 3) = VC12.Cells(R + rmod, 8)
    BIL.Cells(R + 1, 4) = BIL.Cells(R + 1, 3) - BIL.Cells(R + 1, 2)
    
    If BIL.Cells(R + 1, 4) < 0 Then
    BIL.Cells(R + 1, 4).Font.ColorIndex = 3
    Else
    BIL.Cells(R + 1, 4).Font.ColorIndex = 4
    End If
    
    BIL.Cells(R + 1, 4).Interior.Color = RGB(242, 242, 242)
    BIL.Cells(R + 1, 4).Borders(xlEdgeRight).Weight = xlThin

'---DMP START
Union(BIL.Cells(R + 1, "CM"), BIL.Cells(R + 1, "CY")).Borders(xlEdgeRight).Weight = xlThin 'DMP
Range(BIL.Cells(R + 1, 2), BIL.Cells(R + 1, 103)).NumberFormat = "0" 'DMP
'---DMP STOP

End Sub
