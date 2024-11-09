Sub Email_Auto()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim mailBody As String
    Dim ws As Worksheet
    Dim i As Long
    Dim recipient As String
    Dim recipientName As String
    Dim customMessage As String

    Set ws = ThisWorkbook.Sheets("Sheet1")

    On Error Resume Next
    Set OutlookApp = GetObject(Class:="Outlook.Application")

    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject(Class:="Outlook.Application")
    End If
    On Error GoTo 0

    For i = 1 To 4
        recipient = ws.Cells(i, 1).Value
        recipientName = ws.Cells(i, 2).Value
        customMessage = ws.Cells(i, 3).Value

        If recipientName <> "" Then
            mailBody = "<p>Dear " & recipientName & ",</p>" & _
                       "<p>" & customMessage & "</p>" & _
                       "<p>Best regards,<br>pranav</p>"

            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                .To = recipient
                .Subject = "Automated Email from Excel"
                .HTMLBody = mailBody
                .Display
                .Send
            End With
            Set OutlookMail = Nothing
        End If
    Next i

    MsgBox "Emails processed successfully!"
    ActiveWorkbook.Save
End Sub
