Attribute VB_Name = "STAT_OsobyProdukty"
Sub VilTul()
Dim VT As Worksheet
Dim PBI As Worksheet
Dim ea As Worksheet
Dim inc As Worksheet

Dim cpbi As Integer
Dim cinc As Integer
Dim cea As Integer, duplikaty As Integer
Dim cv As Integer
Dim proc As Integer
Dim xos As Integer, xprod As Integer
Dim i As Integer


Set VT = Worksheets("STAT")
Set PBI = Worksheets("Raport PBI")
Set ea = Worksheets("Zadania ADM i DEV")
Set inc = Worksheets("Raport INC")

'czyszczenie
If VT.Cells(3, 1) <> "" Then
xos = Application.WorksheetFunction.CountA(VT.Columns(1)) + 1
Range(VT.Cells(3, 1), VT.Cells(xos, 10)).Clear


VT.Cells.Borders.LineStyle = xlNone

Else
End If

If VT.Cells(3, 12) <> "" Then
xprod = Application.WorksheetFunction.CountA(VT.Columns(12)) + 1
Range(VT.Cells(3, 12), VT.Cells(xprod, 19)).Clear
Else
End If

cv = 3
cpbi = 2
cinc = 3
cea = 2

'podliczenie dla przypisanych
Do While cpbi <= Application.WorksheetFunction.CountA(PBI.Columns(1))
If Application.WorksheetFunction.CountIf(VT.Columns(1), PBI.Cells(cpbi, "K")) < 1 Then
VT.Cells(cv, "A") = PBI.Cells(cpbi, "K")
cpbi = cpbi + 1
cv = cv + 1
Else
cpbi = cpbi + 1
End If
Loop
'x = x
Do While cinc <= Application.WorksheetFunction.CountA(inc.Columns(1))
    If IsError(inc.Cells(cinc, "G")) = False Then
            If Application.WorksheetFunction.CountIf(VT.Columns(1), inc.Cells(cinc, "G")) < 1 Then
                If inc.Cells(cinc, "G") = "-" Then
                cinc = cinc + 1
                Else
                VT.Cells(cv, "A") = inc.Cells(cinc, "G")
                cv = cv + 1
                cinc = cinc + 1
                End If
            Else
            cinc = cinc + 1
            End If
        Else
        cinc = cinc + 1
    End If
Loop
'x = x

Do While cea <= Application.WorksheetFunction.CountA(ea.Columns(1))
    If Application.WorksheetFunction.CountIf(VT.Columns(1), ea.Cells(cea, "H")) < 1 Then
        If ea.Cells(cea, "H") <> "#Informacje o pracach#" Then
        VT.Cells(cv, "A") = ea.Cells(cea, "H")
        cea = cea + 1
        cv = cv + 1
        Else
        cea = cea + 1
        End If
    Else
    cea = cea + 1
    End If
Loop


'duplikaty = Application.WorksheetFunction.CountA(PBI.Columns(1))


Range(VT.Cells(3, "A"), VT.Cells(Application.WorksheetFunction.CountA(VT.Columns(1)) + 1, "A")).Sort _
Key1:=VT.Columns("A"), Order1:=xlAscending, Header:=xlNo


For cv = 3 To Application.WorksheetFunction.CountA(VT.Columns(1)) + 1
VT.Cells(cv, "B") = Application.WorksheetFunction.CountIf((PBI.Columns("K")), VT.Cells(cv, 1))
VT.Cells(cv, "C") = Application.WorksheetFunction.CountIf((inc.Columns("G")), VT.Cells(cv, 1))

VT.Cells(cv, "D") = Application.WorksheetFunction.CountIfs(PBI.Columns("K"), VT.Cells(cv, 1), PBI.Columns("F"), "Pending")
VT.Cells(cv, "E") = Application.WorksheetFunction.CountIfs(inc.Columns("G"), VT.Cells(cv, 1), inc.Columns("C"), "Pending")

VT.Cells(cv, "F") = (Application.WorksheetFunction.CountIfs(PBI.Columns("K"), VT.Cells(cv, 1), PBI.Columns("F"), "Assigned")) + Application.WorksheetFunction.CountIfs(PBI.Columns("K"), VT.Cells(cv, 1), PBI.Columns("F"), "Draft") ' bo foch na funkcje :) (popraw)

VT.Cells(cv, "G") = VT.Cells(cv, "C") - VT.Cells(cv, "E")




VT.Cells(cv, "H") = Application.WorksheetFunction.CountIf((ea.Columns("H")), VT.Cells(cv, 1))
VT.Cells(cv, "I") = Application.WorksheetFunction.Sum(Range(VT.Cells(cv, "F"), VT.Cells(cv, "H")))

Next

VT.Cells(cv, "A") = "Suma"
VT.Cells(cv, "B") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "B"), VT.Cells(cv - 1, "B")))
VT.Cells(cv, "C") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "C"), VT.Cells(cv - 1, "C")))
VT.Cells(cv, "D") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "D"), VT.Cells(cv - 1, "D")))
VT.Cells(cv, "E") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "E"), VT.Cells(cv - 1, "E")))
VT.Cells(cv, "F") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "F"), VT.Cells(cv - 1, "F")))
VT.Cells(cv, "G") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "G"), VT.Cells(cv - 1, "G")))
VT.Cells(cv, "H") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "H"), VT.Cells(cv - 1, "H")))
VT.Cells(cv, "I") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "I"), VT.Cells(cv - 1, "I")))

For proc = 3 To Application.WorksheetFunction.CountA(VT.Columns(1))
VT.Cells(proc, "J") = VT.Cells(proc, "I") / VT.Cells(cv, "I")
Next

VT.Cells(cv, "J") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "J"), VT.Cells(cv - 1, "J")))
'formatowania

Union(Range(VT.Cells(cv, "A"), VT.Cells(cv, "J")), Range(VT.Cells(3, "J"), VT.Cells(cv, "J")), Range(VT.Cells(3, "I"), _
VT.Cells(cv, "I")), Range(VT.Cells(3, "B"), VT.Cells(cv, "B"))).Font.Bold = True

Range(VT.Cells(3, "D"), VT.Cells(cv, "E")).Interior.Color = RGB(234, 234, 234)

Union(Range(VT.Cells(cv, "A"), VT.Cells(cv, "J")), Range(VT.Cells(3, "J"), VT.Cells(cv, "J")), Range(VT.Cells(3, "B"), _
VT.Cells(cv, "C"))).Interior.Color = RGB(192, 0, 0)
Union(Range(VT.Cells(cv, "A"), VT.Cells(cv, "J")), Range(VT.Cells(3, "J"), VT.Cells(cv, "J")), Range(VT.Cells(3, "B"), _
VT.Cells(cv, "C"))).Font.ColorIndex = 2

Range(VT.Cells(3, "J"), VT.Cells(cv, "J")).NumberFormat = "0.00%"
Range(VT.Cells(3, 2), VT.Cells(cv, 10)).HorizontalAlignment = xlCenter
Range(VT.Cells(3, "B"), VT.Cells(cv, "I")).NumberFormat = "0"

' kolorowanie max
For i = 3 To cv
    If VT.Cells(i, "I") = Application.WorksheetFunction.max((Range(VT.Cells(3, "I"), VT.Cells(cv - 1, "I")))) Then
    Union(Range(VT.Cells(i, "A"), VT.Cells(i, "A")), Range(VT.Cells(i, "D"), VT.Cells(i, "I"))).Interior.Color = RGB(242, 197, 192)
    End If
Next

x1 = cv

'DLA udzia³u PBI

cv = 3
cpbi = 2

Do While cpbi <= Application.WorksheetFunction.CountA(PBI.Columns(1))
    If Application.WorksheetFunction.CountIf(VT.Columns(12), PBI.Cells(cpbi, "C")) < 1 Then 'xxxx
    VT.Cells(cv, "L") = PBI.Cells(cpbi, "C")
    cpbi = cpbi + 1
    cv = cv + 1
    Else
    cpbi = cpbi + 1
    End If
Loop


Range(VT.Cells(3, "L"), VT.Cells((Application.WorksheetFunction.CountA(VT.Columns(12))) + 1, "L")).Sort _
Key1:=VT.Columns("L"), Order1:=xlAscending, Header:=xlNo

For cv = 3 To (Application.WorksheetFunction.CountA(VT.Columns(12)) + 1)


VT.Cells(cv, "M") = Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Assigned", _
PBI.Columns("J"), VT.Cells(2, "M")) + Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Draft") ' bo foch na funkcje :)
VT.Cells(cv, "N") = Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Assigned", _
PBI.Columns("J"), VT.Cells(2, "N")) + Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Draft") ' bo foch na funkcje :)
VT.Cells(cv, "O") = Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Assigned", _
PBI.Columns("J"), VT.Cells(2, "O")) + Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Draft") ' bo foch na funkcje :)
VT.Cells(cv, "P") = Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Assigned", _
PBI.Columns("J"), VT.Cells(2, "P")) + Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Draft") ' bo foch na funkcje :)
VT.Cells(cv, "Q") = Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Assigned", _
PBI.Columns("J"), VT.Cells(2, "Q")) + Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Draft") ' bo foch na funkcje :)
VT.Cells(cv, "R") = Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Assigned") + _
Application.WorksheetFunction.CountIfs(PBI.Columns("C"), VT.Cells(cv, 12), PBI.Columns("F"), "Draft") ' bo foch na funkcje :)

Next

VT.Cells(cv, "L") = "Suma"
VT.Cells(cv, "M") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "M"), VT.Cells(cv - 1, "M")))
VT.Cells(cv, "N") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "N"), VT.Cells(cv - 1, "N")))
VT.Cells(cv, "O") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "O"), VT.Cells(cv - 1, "O")))
VT.Cells(cv, "P") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "P"), VT.Cells(cv - 1, "P")))
VT.Cells(cv, "Q") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "Q"), VT.Cells(cv - 1, "Q")))
VT.Cells(cv, "R") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "R"), VT.Cells(cv - 1, "R")))




For proc = 3 To (Application.WorksheetFunction.CountA(VT.Columns(12)) + 1)
VT.Cells(proc, "S") = VT.Cells(proc, "R") / VT.Cells(cv, "R")
Next

VT.Cells(cv, "S") = Application.WorksheetFunction.Sum(Range(VT.Cells(3, "S"), VT.Cells(cv - 1, "S")))

Union(Range(VT.Cells(cv, "L"), VT.Cells(cv, "S")), Range(VT.Cells(3, "S"), VT.Cells(cv, "S")), Range(VT.Cells(3, "R"), VT.Cells(cv, "R"))).Font.Bold = True


Union(Range(VT.Cells(cv, "L"), VT.Cells(cv, "S")), Range(VT.Cells(3, "S"), VT.Cells(cv, "S"))).Interior.Color = RGB(192, 0, 0)
Union(Range(VT.Cells(cv, "L"), VT.Cells(cv, "S")), Range(VT.Cells(3, "S"), VT.Cells(cv, "S"))).Font.ColorIndex = 2

Range(VT.Cells(3, "S"), VT.Cells(cv, "S")).NumberFormat = "0.00%"
Range(VT.Cells(3, 13), VT.Cells(cv, 19)).HorizontalAlignment = xlCenter
Range(VT.Cells(3, 12), VT.Cells(cv, 12)).HorizontalAlignment = xlLeft 'v55
Range(VT.Cells(3, "M"), VT.Cells(cv, "R")).NumberFormat = "0"


' kolorowanie max
For i = 3 To cv
    If VT.Cells(i, "R") = Application.WorksheetFunction.max((Range(VT.Cells(3, "R"), VT.Cells(cv - 1, "R")))) Then
    Range(VT.Cells(i, "L"), VT.Cells(i, "R")).Interior.Color = RGB(242, 197, 192)
    End If
Next

x2 = cv
'obramowania


Union(Range(VT.Cells(1, "A"), VT.Cells(2, "J")), Range(VT.Cells(x1, "A"), VT.Cells(x1, "J"))).Borders(xlEdgeBottom).Weight = xlMedium
Range(VT.Cells(1, "D"), VT.Cells(1, "H")).Borders(xlEdgeBottom).Weight = xlMedium
Union(Range(VT.Cells(1, "A"), VT.Cells(2, "J")), Range(VT.Cells(x1, "A"), VT.Cells(x1, "J"))).Borders(xlEdgeTop).Weight = xlMedium
Union(Range(VT.Cells(1, "A"), VT.Cells(2, "J")), Range(VT.Cells(x1, "A"), VT.Cells(x1, "J")), Range(VT.Cells(1, "A"), VT.Cells(x1, "A")), _
Range(VT.Cells(3, "J"), VT.Cells(x1, "J")), Range(VT.Cells(1, "H"), VT.Cells(x1, "H")), Range(VT.Cells(2, "E"), _
VT.Cells(x1, "E"))).Borders(xlEdgeRight).Weight = xlMedium
Union(Range(VT.Cells(1, "A"), VT.Cells(2, "J")), Range(VT.Cells(x1, "A"), VT.Cells(x1, "J")), Range(VT.Cells(3, "A"), VT.Cells(x1, "A")), _
Range(VT.Cells(1, "D"), VT.Cells(x1, "D")), Range(VT.Cells(1, "J"), VT.Cells(x1, "J"))).Borders(xlEdgeLeft).Weight = xlMedium
Range(VT.Cells(1, "D"), VT.Cells(x1, "D")).Borders(xlEdgeLeft).Weight = xlMedium

Union(Range(VT.Cells(2, "L"), VT.Cells(1, "S")), Range(VT.Cells(x2, "L"), VT.Cells(x2, "S")), Range(VT.Cells(3, "R"), VT.Cells(x2, "R"))). _
Borders(xlEdgeBottom).Weight = xlMedium
Union(Range(VT.Cells(1, "L"), VT.Cells(1, "S")), Range(VT.Cells(x2, "L"), VT.Cells(x2, "S")), Range(VT.Cells(3, "R"), VT.Cells(x2, "R"))). _
Borders(xlEdgeTop).Weight = xlMedium
Union(Range(VT.Cells(1, "L"), VT.Cells(2, "S")), Range(VT.Cells(x2, "L"), VT.Cells(x2, "S")), Range(VT.Cells(3, "L"), VT.Cells(x2, "L")), _
Range(VT.Cells(3, "S"), VT.Cells(x2, "S")), Range(VT.Cells(3, "Q"), VT.Cells(x2, "Q"))).Borders(xlEdgeRight).Weight = xlMedium
Union(Range(VT.Cells(1, "L"), VT.Cells(2, "S")), Range(VT.Cells(x2, "L"), VT.Cells(x2, "S")), Range(VT.Cells(3, "L"), VT.Cells(x2, "L")), _
Range(VT.Cells(3, "S"), VT.Cells(x2, "S"))).Borders(xlEdgeLeft).Weight = xlMedium
Range(VT.Cells(1, "M"), VT.Cells(2, "M")).Borders(xlEdgeLeft).Weight = xlMedium
Range(VT.Cells(1, "R"), VT.Cells(2, "R")).Borders(xlEdgeLeft).Weight = xlMedium
Range(VT.Cells(1, "S"), VT.Cells(2, "S")).Borders(xlEdgeLeft).Weight = xlMedium
Range(VT.Cells(1, "M"), VT.Cells(1, "Q")).Borders(xlEdgeBottom).Weight = xlMedium


VT.Cells(x2 - 1, "R").Borders(xlEdgeBottom).Weight = xlMedium ' bo wycina³o ten jebany dól, chuj wie czemu ale nie chce mi sie z tym dlu¿ej walczyæ

VT.Activate
ActiveWindow.ScrollColumn = 1
ActiveWindow.ScrollRow = 1
VT.Cells(1, 1).Activate

End Sub



