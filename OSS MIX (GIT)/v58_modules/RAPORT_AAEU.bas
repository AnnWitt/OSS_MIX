Attribute VB_Name = "RAPORT_AAEU"
Sub ZUA()
Dim ea As Worksheet, EA_SRC As Worksheet
Dim le As Integer


Set EA_SRC = Worksheets("EU_AA")
Set ea = Worksheets("Zadania ADM i DEV")


For le = 2 To WorksheetFunction.CountA(EA_SRC.Columns(1))

ea.Cells(le, "A") = EA_SRC.Cells(le, "A")
ea.Cells(le, "B") = EA_SRC.Cells(le, "B")
ea.Cells(le, "C") = EA_SRC.Cells(le, "E")
ea.Cells(le, "D") = EA_SRC.Cells(le, "C")
ea.Cells(le, "E") = EA_SRC.Cells(le, "K")
ea.Cells(le, "F") = EA_SRC.Cells(le, "L")
ea.Cells(le, "H") = EA_SRC.Cells(le, "F")
ea.Cells(le, "I") = EA_SRC.Cells(le, "D")
ea.Cells(le, "J") = EA_SRC.Cells(le, "H")
ea.Cells(le, "O") = EA_SRC.Cells(le, "O") 'v50
ea.Cells(le, "P") = EA_SRC.Cells(le, "P") 'v50

    If IsEmpty(EA_SRC.Cells(le, "I")) = False Then
    ea.Cells(le, "K") = EA_SRC.Cells(le, "I")
    Else
    ea.Cells(le, "K") = EA_SRC.Cells(le, "Q") 'v50
    End If
    
    If EA_SRC.Cells(le, "N") <> "" Then
    ea.Cells(le, "L") = EA_SRC.Cells(le, "N")
    Else
    ea.Cells(le, "L") = ""
    End If

ea.Cells(le, "M") = EA_SRC.Cells(le, "M")
ea.Cells(le, "N") = EA_SRC.Cells(le, "J")

If (Left(ea.Cells(le, "H"), 3) <> "Inf" And Left(ea.Cells(le, "H"), 3) <> "#ND") Then
ea.Cells(le, "G") = Left(ea.Cells(le, "H"), 3)
Else
ea.Cells(le, "G") = "-"
End If

Union(ea.Columns("J:K"), ea.Columns("O:O")).NumberFormat = "yyyy/mm/dd hh:mm:ss" 'v50


If ea.Cells(le, "J") >= Now() Then
    'prace weekendowe (najblizszy weekend)
    If (WorksheetFunction.Weekday(ea.Cells(le, "J"), vbMonday) >= 5 And ea.Cells(le, "J") < Now() + 3) Then
    Range(ea.Cells(le, "A"), ea.Cells(le, "P")).Interior.Color = RGB(128, 248, 225) 'bylo 225 v50
    Else
    ''prace na 24h
        If (WorksheetFunction.Weekday(ea.Cells(le, "J"), vbMonday) < 5 And (ea.Cells(le, "J") - Now()) >= 0 And ea.Cells(le, "J") - Now() <= 1) Then
        Range(ea.Cells(le, "A"), ea.Cells(le, "P")).Interior.Color = RGB(128, 248, 225) 'v50
        Else
        End If
    End If
Else

End If


Next

Range(ea.Rows(2), ea.Rows(WorksheetFunction.CountA(ea.Columns(1)))).RowHeight = 20
Range(ea.Rows(2), ea.Rows(WorksheetFunction.CountA(ea.Columns(1)))).Font.Name = "Calibri"
Range(ea.Rows(2), ea.Rows(WorksheetFunction.CountA(ea.Columns(1)))).Font.Size = 11

'licznik mo¿e byc jeden

Union(ea.Columns("D:D"), ea.Columns("G:G"), ea.Columns("J:L"), ea.Columns("O:O")).HorizontalAlignment = xlCenter 'v50
ea.Columns("A:P").VerticalAlignment = xlCenter 'v50

Range(ea.Cells(2, "A"), ea.Cells(le, "P")).Sort _
Key1:=ea.Columns("J"), Order1:=xlAscending, Header:=xlNo 'v50




ea.Activate
ActiveWindow.ScrollColumn = 1
ActiveWindow.ScrollRow = 1
ea.Cells(1, 1).Activate

End Sub

