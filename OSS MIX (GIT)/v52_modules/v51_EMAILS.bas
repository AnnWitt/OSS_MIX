Attribute VB_Name = "EMAILS"

Sub adresaci()

Dim sz As Worksheet, emails As Worksheet, stat As Worksheet
Set sz = Worksheets("Konfiguracja")
Set emails = Worksheets("emails")
Set stat = Worksheets("STAT")
Dim msgTo As String, msgCC As String



emails.Cells.ClearContents

Dim i As Integer

emails.Cells(1, 1) = "Do"
emails.Cells(2, 1) = "DW"

Dim licznik As Integer
For licznik = 3 To Application.WorksheetFunction.CountA(stat.Columns(1))
If Application.VLookup(stat.Cells(licznik, "A"), Range(sz.Cells(2, 27), sz.Cells(Application.WorksheetFunction.CountA(sz.Columns(28)), 28)), 2, 0) <> "#ND" Then
msgTo = msgTo & Application.VLookup(stat.Cells(licznik, "A"), Range(sz.Cells(2, 27), sz.Cells(Application.WorksheetFunction.CountA(sz.Columns(28)), 28)), 2, 0) & ";"
Else
End If
Next

For licznik = 2 To Application.WorksheetFunction.CountA(sz.Columns(30))
If Application.VLookup(sz.Cells(licznik, "AD"), Range(sz.Cells(2, 27), sz.Cells(Application.WorksheetFunction.CountA(sz.Columns(28)), 28)), 2, 0) <> "#ND" Then
msgCC = msgCC & Application.VLookup(sz.Cells(licznik, "AD"), Range(sz.Cells(2, 27), sz.Cells(Application.WorksheetFunction.CountA(sz.Columns(28)), 28)), 2, 0) & ";"
Else
End If
Next
emails.Cells(2, "B") = msgCC


emails.Cells(1, "B") = msgTo



       
       odp = MsgBox(sz.Cells(45, "X") & vbline, vbYesNo + vbQuestion)
        If odp = vbYes Then
        MsgBox (sz.Cells(46, "X"))
        Call Send_Emails
        Else
        End If

Worksheets("emails").Activate


End Sub


Sub Send_Emails()

  Dim OutlookApp As Outlook.Application
  Dim OutlookMail As Outlook.MailItem
  
  Dim data As String, godz As String


        If Len(Month(Now())) < 2 Then
            data = Year(Now()) & "0" & Month(Now())
        Else
            data = Year(Now()) & Month(Now())
        End If
        
        If Len(Day(Now())) < 2 Then
            data = data & "0" & Day(Now())
        Else
            data = data & Day(Now())
        End If
        
        
        If Len(Hour(Now())) < 2 Then
            godz = "0" & Hour(Now())
        Else
            godz = Hour(Now())
        End If
        
        If Len(Minute(Now())) < 2 Then
            godz = godz & "0" & Minute(Now())
        Else
            godz = godz & Minute(Now())
        End If


  Set OutlookApp = New Outlook.Application
  Set OutlookMail = OutlookApp.CreateItem(olMailItem)

  With OutlookMail
    .BodyFormat = olFormatHTML
    .Display
    .To = Worksheets("emails").Cells(1, 2)
    .CC = Worksheets("emails").Cells(2, 2)
    .Subject = "Orange OSS - Raport Dzienny " & data & "_" & godz
    '.Attachments = ThisWorkbook runtime error bo tylko do odczytu
    '.Send
  End With
  
End Sub
