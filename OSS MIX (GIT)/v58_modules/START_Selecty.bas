Attribute VB_Name = "START_Selecty"
Sub selecty() 'OSS_MIX DMP - DOPISANIE GRUPY KOBAT KOLEKTOR !!

x = """"

'te dwa pierwsze w sumie niepotrzebne bo z dziennego to ma isc ale co tam, niech bedzie :)  or 'Assigned Group*+'="MIESZKO_VENDOR"
Worksheets("CSV").Cells(2, "N") = "('Assigned Group*+' LIKE " & x & "VC_OSS_FIXED_%" & x & " or " & "'Assigned Group*+' LIKE " & x & "VC_TP_OSS_%" & x & " or " & _
"'Assigned Group*+' = " & x & "MIESZKO_VENDOR" & x & " or " & "'Assigned Group*+' = " & x & "APLIKACJE_ATRIUM" & x & " or " & "'Assigned Group*+' = " & x & "DOSTAWCA_ATRIUM" & x & ")" & " AND 'Status*' < " & x & "Resolved" & x

Worksheets("CSV").Cells(3, "N") = "('Assigned Group*+' LIKE " & x & "VC_OSS_FIXED_%" & x & " or " & "'Assigned Group*+' LIKE " & x & "VC_TP_OSS_%" & x & " or " & _
"'Assigned Group*+' = " & x & "MIESZKO_VENDOR" & x & " or " & "'Assigned Group*+' = " & x & "APLIKACJE_ATRIUM" & x & " or " & "'Assigned Group*+' = " & x & "DOSTAWCA_ATRIUM" & x & ")" & " AND 'Status*' < " & x & "Pending" & x & " AND ('Resolve to'<$TIMESTAMP$ " & ")"


Worksheets("CSV").Cells(4, "N") = "('Assigned Group*+' LIKE " & x & "VC_OSS_FIXED_%" & x & " or " & "'Assigned Group*+' LIKE " & x & "VC_TP_OSS_%" & x & " or " & _
"'Assigned Group*+' = " & x & "MIESZKO_VENDOR" & x & " or " & "'Assigned Group*+' = " & x & "DOSTAWCA_ATRIUM" & x & ")" & " AND 'Status*' >= " & x & "Resolved" & x & " AND 'Last Resolved Date' >" & x & Worksheets("GO").Cells(4, "L") & " 00:00:00" & x & " AND 'Last Resolved Date' < " & x & Worksheets("GO").Cells(5, "L") & " 23:59:59" & x
Worksheets("CSV").Cells(5, "N") = "('Assigned Group*+' LIKE " & x & "VC_OSS_FIXED_%" & x & " or " & "'Assigned Group*+' LIKE " & x & "VC_TP_OSS_%" & x & " or " & _
"'Assigned Group*+' = " & x & "MIESZKO_VENDOR" & x & " or " & "'Assigned Group*+' = " & x & "DOSTAWCA_ATRIUM" & x & ")" & " AND 'Submit Date' >" & x & Worksheets("GO").Cells(4, "L") & " 00:00:00" & x & " AND 'Submit Date' < " & x & Worksheets("GO").Cells(5, "L") & " 23:59:59" & x
Worksheets("CSV").Cells(6, "N") = "('Assigned Group*+' LIKE " & x & "VC_OSS_FIXED_%" & x & " or " & "'Assigned Group*+' LIKE " & x & "VC_TP_OSS_%" & x & " or " & _
"'Assigned Group*+' = " & x & "MIESZKO_VENDOR" & x & " or " & "'Assigned Group*+' = " & x & "DOSTAWCA_ATRIUM" & x & ")" & " AND ('Resolve to'<'Last Resolved Date'  AND 'Status*' >= " & x & "Resolved" & x & ") AND ('Last Resolved Date' >" & x & Worksheets("GO").Cells(4, "L") & " 00:00:00" & x & " AND 'Last Resolved Date' < " & x & Worksheets("GO").Cells(5, "L") & " 23:59:59" & x & ")"

Call sel_oss

End Sub


Sub Przycisk9_Klikniecie() 'OSS_MIX
Range("l4").Value = Range("l4") + 1
'Range("l4").NumberFormat = "yyyy-mm-dd"
Call calc1
Call selecty
End Sub

Sub Przycisk7_Klikniecie() 'OSS_MIX
Range("l4").Value = Range("l4") - 1
Call calc1
Call selecty
End Sub
Sub Przycisk10_Klikniecie() 'OSS_MIX
Range("l5").Value = Range("l5") + 1
Call calc2
Call selecty
End Sub
Sub Przycisk11_Klikniecie() 'OSS_MIX
Range("l5").Value = Range("l5") - 1
Call calc2
Call selecty
End Sub
Sub CSV_dataplus_Klikniecie() 'OSS_MIX
Range("j8").Value = Range("j8").Value + 1 'v52
Call selecty
End Sub

Sub CSV_dataminus_Klikniecie() 'OSS_MIX
Range("j8").Value = Range("j8").Value - 1
Call selecty
End Sub

Function calc2() 'OSS_MIX
Dim m As String, z As String, d As String

A = Range("l5") 'v52
z = 0

If Month(A) > 9 Then
m = Month(A)
Else
m = z & Month(A)

End If

If Day(A) > 9 Then
d = Day(A)
Else
d = "0" & Day(A)
End If

'
'Range("l5").NumberFormat = "@"
'Range("l5") = (Year(A)) & "-" & m & "-" & d
Range("l13").NumberFormat = "@"
Range("l13") = (Year(A)) & "-" & m & "-" & d
Call sel_oss
End Function

Function calc1() 'OSS_MIX
Dim m As String, z As String, d As String

A = Range("l4")
z = 0

If Month(A) > 9 Then
m = Month(A)
Else
m = z & Month(A)

End If

If Day(A) > 9 Then
d = Day(A)
Else
d = "0" & Day(A)
End If


'Range("l5").NumberFormat = "@"
'Range("l5") = (Year(A)) & "-" & m & "-" & d
'Range("l4").NumberFormat = "@"
'Range("l4") = (Year(A)) & "-" & m & "-" & d
Range("k13").NumberFormat = "@"
Range("k13") = (Year(A)) & "-" & m & "-" & d
Call sel_oss
End Function


Sub sel_oss()
Dim go As Worksheet, sz As Worksheet, x As String

Set go = Worksheets("GO")
Set sz = Worksheets("Konfiguracja")


'go.Cells(13, "K") = go.Cells(4, "L")
'go.Cells(13, "L") = go.Cells(5, "L")

'go.Range(Cells(13, "K"), Cells(13, "L")).NumberFormat = "yyyy-mm-dd"
'go.Cells(13, "K") = WorksheetFunction.Text(go.Cells(4, "L"), "yyyy-mm-dd")
'go.Cells(13, "L") = go.Cells(5, "L")


'problem z datami 2022-03-08


x = """"

go.Cells(13, "L") = WorksheetFunction.Text(go.Cells(13, "L"), "YYYY-MM-DD")
go.Cells(13, "K") = WorksheetFunction.Text(go.Cells(13, "K"), "YYYY-MM-DD")

go.Cells(14, "I") = "(" & x & sz.Cells(15, "X") & x & " ~ " & x & "PBI*" & x & ") and (" & x & "Data Start TP" & x & ">" & x & go.Cells(13, "K") & " 00:00" & x & " and " & x & "Data Start TP" & x & "<" & x & go.Cells(13, "L") & " 23:59" & x & ") "

End Sub
