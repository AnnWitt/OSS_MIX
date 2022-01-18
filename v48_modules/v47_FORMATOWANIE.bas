Attribute VB_Name = "FORMATOWANIE"

Sub poprawaformatow() 'DMP - DO EW KOREKT
Dim lp As Integer
Dim arkusz As String

Dim sz As Worksheet, temp As Worksheet, BIL As Worksheet

Set sz = Worksheets("Konfiguracja")
    
    '--DMP START
For lp = 3 To 38 'v20210929
'--DMP STOP
    'nowy wpis

    arkusz = sz.Cells(lp, "N")
    Set temp = Worksheets(arkusz)

    Range(temp.Cells(3, "L"), temp.Cells(WorksheetFunction.CountA(temp.Columns(1)), "L")).NumberFormat = "0" 'v20210929
Next

Set BIL = Worksheets("Zestawienie Grup")



'--DMP START
For R = 4 To WorksheetFunction.CountA(BIL.Columns(1))
    If BIL.Cells(R, "CV") < 0 Then
    BIL.Cells(R, "CV").Font.ColorIndex = 3
    Else
    BIL.Cells(R, "CV").Font.ColorIndex = 4
    End If
    
    
    '-------------czy potrzebne ????
'    BIL.Cells(R, "CP").Interior.Color = RGB(242, 242, 242)
'    BIL.Cells(R, "CM").Interior.Color = RGB(242, 242, 242)
    Next

End Sub
'--DMP STOP
