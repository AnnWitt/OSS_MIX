Attribute VB_Name = "FORMATOWANIE"

Sub poprawaformatow() 'DMP - DO EW KOREKT
Dim lp As Integer
Dim arkusz As String

Dim sz As Worksheet, temp As Worksheet, BIL As Worksheet

Set sz = Worksheets("Konfiguracja")
    
    '--DMP START
For lp = 3 To 37
'--DMP STOP
    'nowy wpis
    
    arkusz = sz.Cells(lp, "N")
    Set temp = Worksheets(arkusz)

    Range(temp.Cells(3, "L"), temp.Cells(5000, "L")).NumberFormat = "0"
Next

Set BIL = Worksheets("Zestawienie Grup")


For R = 4 To 1822
    If BIL.Cells(R, "CS") < 0 Then
    BIL.Cells(R, "CS").Font.ColorIndex = 3
    Else
    BIL.Cells(R, "CS").Font.ColorIndex = 4
    End If
    
    BIL.Cells(R, "CP").Interior.Color = RGB(242, 242, 242)
    BIL.Cells(R, "CM").Interior.Color = RGB(242, 242, 242)
    Next

End Sub
