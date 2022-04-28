Sub mailloop()

For i = 2 To 4
Dim userName As String
Dim sendemail As String
Dim ccmail As String
Dim newVar As String

repname = Sheets("Sheet1").Range("B" & i)
sendemail = Sheets("Sheet1").Range("E" & i)
ccmail = Sheets("Sheet1").Range("F" & i).Value
totaldays = Sheets("Sheet1").Range("h" & i)

customMailsender userName, sendemail, ccmail, newVar
Next


End Sub




Sub customMailsender(rName As String, primemail As String, ccmail As String, newVar As String)

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

        linknheader = "<table border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse'><tr style='height:45.35pt'><td width=641 style='width:481.0pt;background:#003057;padding:0px 0px 0px 20px';height:45.35pt'><font style='font-size:24.0pt;font-family:Arial;color:white;font-weight:600;letter-spacing:.6px'>Mini Bid Shift Change</font></td></tr></table>" _
        & "<br>Congratulations " & rName & "!<br>" _
       
    customstring = "<br> Here's a string to start off the template. Here's another string that leads into your custom varialbe. " & newVar & " and also a part that leads out of it."
    closeit = "<br>Quick note on this other thing. <br> <br> Closing! <br><br><br> Note: I can also send things in CC to a string of emails. "

    On Error Resume Next
    With OutMail
        .To = primemail
        .CC = ccmail
        .BCC = ""
        .Subject = yourSubject
        .HTMLBody = linknheader & customstring & closeit
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
    '    .Send   'or use
        .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
