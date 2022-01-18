Attribute VB_Name = "STAT_DAily"
Sub pdf() 'OSS_MIX
Dim Daily As Worksheet
Dim plik As String


Set Daily = Worksheets("Daily")

Daily.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\ " & Daily.Cells(1, 1) & " OSS_INC.pdf"
MsgBox "Arkusz daily wyeksportowany do lokalizacji szablonu"

End Sub

Sub Daily() 'OSS_MIX //do przerobki po licz je¿eli
Dim Daily, SRC, csv As Worksheet
Dim odp As Integer

Set SRC = Worksheets("STAT_SRC")
Set Daily = Worksheets("Daily")

Set csv = Worksheets("CSV")


Range(Daily.Cells(3, "C"), Daily.Cells(7, "G")) = 0


Daily.Cells(1, 1) = Worksheets("GO").Cells(8, "J")
'bez dzialan wspierajacych

'--DMP START
Daily.Cells(3, 3) = SRC.Cells(3, 2) - SRC.Cells(35, 2)
Daily.Cells(3, 4) = SRC.Cells(3, 3) - SRC.Cells(35, 3)
Daily.Cells(3, 5) = SRC.Cells(3, 4) - SRC.Cells(35, 4)
Daily.Cells(3, 6) = SRC.Cells(3, 5) - SRC.Cells(35, 5)
Daily.Cells(3, 7) = SRC.Cells(3, 6) - SRC.Cells(35, 6)

    For rd = 4 To 34
    
    Select Case SRC.Cells(rd, "G")
    Case Is = "GOLD"
    Daily.Cells(4, "C") = Daily.Cells(4, "C") + SRC.Cells(rd, "B")
    Daily.Cells(4, "D") = Daily.Cells(4, "D") + SRC.Cells(rd, "C")
    Daily.Cells(4, "E") = Daily.Cells(4, "E") + SRC.Cells(rd, "D")
    Daily.Cells(4, "F") = Daily.Cells(4, "F") + SRC.Cells(rd, "E")
    Daily.Cells(4, "G") = Daily.Cells(4, "G") + SRC.Cells(rd, "F")
    Case Is = "SILVER"
    Daily.Cells(5, "C") = Daily.Cells(5, "C") + SRC.Cells(rd, "B")
    Daily.Cells(5, "D") = Daily.Cells(5, "D") + SRC.Cells(rd, "C")
    Daily.Cells(5, "E") = Daily.Cells(5, "E") + SRC.Cells(rd, "D")
    Daily.Cells(5, "F") = Daily.Cells(5, "F") + SRC.Cells(rd, "E")
    Daily.Cells(5, "G") = Daily.Cells(5, "G") + SRC.Cells(rd, "F")
    Case Is = "BRONZE"
    Daily.Cells(6, "C") = Daily.Cells(6, "C") + SRC.Cells(rd, "B")
    Daily.Cells(6, "D") = Daily.Cells(6, "D") + SRC.Cells(rd, "C")
    Daily.Cells(6, "E") = Daily.Cells(6, "E") + SRC.Cells(rd, "D")
    Daily.Cells(6, "F") = Daily.Cells(6, "F") + SRC.Cells(rd, "E")
    Daily.Cells(6, "G") = Daily.Cells(6, "G") + SRC.Cells(rd, "F")
    Case Is = "TOMBAK"
    Daily.Cells(7, "C") = Daily.Cells(7, "C") + SRC.Cells(rd, "B")
    Daily.Cells(7, "D") = Daily.Cells(7, "D") + SRC.Cells(rd, "C")
    Daily.Cells(7, "E") = Daily.Cells(7, "E") + SRC.Cells(rd, "D")
    Daily.Cells(7, "F") = Daily.Cells(7, "F") + SRC.Cells(rd, "E")
    Daily.Cells(7, "G") = Daily.Cells(7, "G") + SRC.Cells(rd, "F")
    End Select
'--DMP STOP

'Daily.Cells(4, "H") = x

Next

'odp = MsgBox("Czy generowaæ daily do pdf?  ", vbYesNo + vbQuestion)
If Worksheets("GO").Cells(10, "K") = "Tak" Then
Call pdf
Else

End If


End Sub
