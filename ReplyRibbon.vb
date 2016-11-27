Imports Microsoft.Office.Tools.Ribbon

Public Class ReplyRibbon


    Private Sub btnStandardReply_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnStandardReply.Click

        ' Used by Honours Reception staff to generate a standard response to e-mail enquiries

        Dim objOL As Object
        Dim curMsg As Outlook.MailItem, replyMsg As Outlook.MailItem
        Dim SentDate As String, strReply As String
        Dim blnReadMode As Boolean

        objOL = CreateObject("Outlook.Application")
        curMsg = objOL.ActiveInspector.CurrentItem
        blnReadMode = curMsg.Sent
        If blnReadMode Then
            SentDate = Format(curMsg.SentOn, "dd MMMM yyyy")
        Else
            SentDate = Format(Now, "dd MMMM yyyy")
        End If
        strReply = "<html><head><style type='text/css'>BODY {font-family: 'Calibri', Helvetica, Arial, sans-serif;}</style></head><body>" & _
            "Thank you for your correspondence of " & SentDate & ".<p>" & _
            "Please accept this acknowledgement of your letter and, unless you feel<br>" & _
            "it is essential, it is not necessary to forward a signed copy of your<br>" & _
            "correspondence by post.<p>" & _
            "Should there be a need, the Honours Secretariat will contact you again.<p>" & _
            "Australian Honours and Awards Secretariat<br>" & _
            "Government House, Canberra<p>&nbsp;<p>" & _
            "<i>Do you know someone who could be considered for an award in the Australian honours system? <a href='https://www.gg.gov.au/australian-honours-and-awards/nomination-forms'>Click here to nominate someone</a></i><p>" & _
            "</body></html>"

        If blnReadMode Then
            replyMsg = curMsg.Reply
            curMsg.Close(Outlook.OlInspectorClose.olDiscard)
            replyMsg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            replyMsg.HTMLBody = strReply
            replyMsg.Display()
        Else
            curMsg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            curMsg.HTMLBody = strReply
        End If
        objOL = Nothing
    End Sub


    Private Sub btnNomReply_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnNomReply.Click

        ' Used by Honours Reception staff to generate a standard response to new nominations

        Dim objOL As Object
        Dim curMsg As Outlook.MailItem, replyMsg As Outlook.MailItem
        Dim ThisDate As String, strReply As String
        Dim blnReadMode As Boolean

        objOL = CreateObject("Outlook.Application")
        curMsg = objOL.ActiveInspector.CurrentItem
        blnReadMode = curMsg.Sent
        ThisDate = Format(Now(), "dd MMMM yyyy")
        strReply = "<html><head><style type='text/css'>BODY {font-family: 'Calibri', Helvetica, Arial, sans-serif;}</style></head><body>" &
            "<p>Dear <name></p>" &
            "<p>Your nomination of <nominee> for an award in the Order of Australia has been received on " & ThisDate & " in the Australian Honours Secretariat.</p>" &
            "<p>The nomination process usually takes between 18 to 24 months from receipt of nomination to complete.  Should there be any further information required you will be contacted again by the Secretariat.</p>" &
            "Australian Honours and Awards Secretariat<br>" &
            "Government House, Canberra<p>" &
            "<i>Do you know someone who could be considered for an award in the Australian honours system? <a href='https://www.gg.gov.au/australian-honours-and-awards/nomination-forms'>Click here to nominate someone</a></i><p></body></html>"

        If blnReadMode Then
            replyMsg = curMsg.Reply
            curMsg.Close(Outlook.OlInspectorClose.olDiscard)
            replyMsg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            replyMsg.HTMLBody = strReply
            replyMsg.Display()
        Else
            curMsg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            curMsg.HTMLBody = strReply
        End If

        objOL = Nothing
    End Sub


    Private Sub btnNEMReply_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnNEMReply.Click

        ' Used by Honours National Emergency Medal staff to generate a standard response to new nominations

        Dim objOL As Object
        Dim curMsg As Outlook.MailItem, replyMsg As Outlook.MailItem
        Dim ThisDate As String, strReply As String
        Dim blnReadMode As Boolean

        objOL = CreateObject("Outlook.Application")
        curMsg = objOL.ActiveInspector.CurrentItem
        blnReadMode = curMsg.Sent
        ThisDate = Format(Now(), "dd MMMM yyyy")
        strReply = "<html><head><style type='text/css'>BODY {font-family: 'Calibri', Helvetica, Arial, sans-serif;}</style></head><body>" & _
            "Your email has been successfully received.<p>" & _
            "Thank you for your interest in the National Emergency Medal.<br>" & _
            "We will be in contact shortly if further action is required.<p>" & _
            "Australian Honours and Awards Secretariat<br>" & _
            "Government House, Canberra<p>" & _
            "<i>Do you know someone who could be considered for an award in the Australian honours system? <a href='https://www.gg.gov.au/australian-honours-and-awards/nomination-forms'>Click here to nominate someone</a></i><p></body></html>"

        If blnReadMode Then
            replyMsg = curMsg.Reply
            curMsg.Close(Outlook.OlInspectorClose.olDiscard)
            replyMsg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            replyMsg.HTMLBody = strReply
            replyMsg.Display()
        Else
            curMsg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            curMsg.HTMLBody = strReply
        End If

        objOL = Nothing
    End Sub


    Private Sub btnGenericResponse_Click(sender As System.Object, e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles btnGenericResponse.Click

        ' Used by Honours Reception staff to generate a generic response to e-mail enquiries

        Dim objOL As Object
        Dim curMsg As Outlook.MailItem, replyMsg As Outlook.MailItem
        Dim strReply As String
        Dim blnReadMode As Boolean

        objOL = CreateObject("Outlook.Application")
        curMsg = objOL.ActiveInspector.CurrentItem
        blnReadMode = curMsg.Sent
        strReply = "<html><head><style type='text/css'>BODY {font-family: 'Calibri', Helvetica, Arial, sans-serif;}</style></head><body>" & _
            "Thank you for your email.<p>" & _
            "It has been forwarded to the appropriate area.<p>" & _
            "Australian Honours and Awards Secretariat<br>" & _
            "Government House, Canberra<p>&nbsp;<p>" & _
            "<i>Do you know someone who could be considered for an award in the Australian honours system? <a href='https://www.gg.gov.au/australian-honours-and-awards/nomination-forms'>Click here to nominate someone</a></i><p></body></html>"

        If blnReadMode Then
            replyMsg = curMsg.Reply
            curMsg.Close(Outlook.OlInspectorClose.olDiscard)
            replyMsg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            replyMsg.HTMLBody = strReply
            replyMsg.Display()
        Else
            curMsg.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            curMsg.HTMLBody = strReply
        End If
        objOL = Nothing
    End Sub


End Class
