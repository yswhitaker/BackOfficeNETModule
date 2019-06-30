Imports System.Net.Mail
Imports System.Text

Public Structure EmailInfo
    Dim From As String
    Dim MailTo As String
    Dim Subject As String
    Dim Body As String
    Dim MailServer As String
    Dim Attach As String
End Structure

Namespace GGBackOffice



    Public Class BAKEmailManager


        Public Sub SendEmail(ByVal strEmailInfo As String)
            Dim strFrom As String = vbNullString
            Dim strTo As String = vbNullString
            Dim strSubject As String = vbNullString
            Dim strBody As String = vbNullString
            Dim strMailServer As String = vbNullString
            Dim strAttach As String = vbNullString

            Dim strEmailArgs() As String = Split(strEmailInfo, "++")

            strFrom = strEmailArgs(0)
            strTo = strEmailArgs(1)
            strSubject = strEmailArgs(2)
            strBody = strEmailArgs(3)
            strMailServer = strEmailArgs(4)

            If Not UBound(strEmailArgs) = 4 Then
                strAttach = strEmailArgs(5)
            End If

            Dim mailServerName As String = strMailServer
            Dim Message As New MailMessage(strFrom, strTo, strSubject, strBody)
            Dim mailClient As New SmtpClient

            Message.IsBodyHtml = False
            mailClient.Host = mailServerName
            mailClient.UseDefaultCredentials = True

            'for running on local machine COMMENT WHEN COMPILED---------------------------------------------------
            Dim basicAuthenticationInfo As New System.Net.NetworkCredential("yvette@yswconsulting.com", "millions")

            mailClient.UseDefaultCredentials = False
            mailClient.Credentials = basicAuthenticationInfo

            '-----------------------------------------------------------------------------------------------------

            If Not Trim(strAttach) = vbNullString Then
                Dim Attach As Attachment = New Attachment(strAttach)
                Message.Attachments.Add(Attach)
            End If

            mailClient.Send(Message)


        End Sub

        Public Sub SendEmail(ByVal objE As EmailInfo, Optional ByVal blRequestReceipt As Boolean = False)

            Dim strFrom As String = objE.From
            Dim strTo As String = objE.MailTo
            Dim strSubject As String = objE.Subject
            Dim strBody As String = objE.Body
            Dim strMailServer As String = objE.MailServer
            Dim strAttach As String = objE.Attach

            Dim mailServerName As String = strMailServer
            Dim Message As New MailMessage(strFrom, strTo, strSubject, strBody)
            Dim mailClient As New SmtpClient

            Message.IsBodyHtml = False
            mailClient.Host = mailServerName
            '3/7/10 - have to use explicit credentials to get this to run from Betty's machine
            'mailClient.UseDefaultCredentials = True

            Dim basicAuthenticationInfo As New System.Net.NetworkCredential("BJM", "queenb1219")

            mailClient.UseDefaultCredentials = False
            mailClient.Credentials = basicAuthenticationInfo

            If Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release" OrElse Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug" Then
                'for running on local machine COMMENT WHEN COMPILED---------------------------------------------------
                'Dim basicAuthenticationInfo As New System.Net.NetworkCredential("yvette@yswconsulting.com", "hopeful1")

                'mailClient.UseDefaultCredentials = False
                'mailClient.Credentials = basicAuthenticationInfo

                '-----------------------------------------------------------------------------------------------------
            End If

            If blRequestReceipt Then
                Message.Headers.Add("Disposition-Notification-To", objE.From)
            End If


            If Not Trim(strAttach) = vbNullString Then
                Dim Attach As Attachment = New Attachment(strAttach)
                Message.Attachments.Add(Attach)
            End If

            mailClient.Send(Message)


        End Sub

        Public Sub SendOnlineInvoiceNotification(ByVal lngInvNumber As Long, ByVal strEmail As String)

            'look up bill attn email address and invoice info with clcode and invnumber

            Dim objE As EmailInfo = New EmailInfo
            Dim sb As StringBuilder = New StringBuilder
          
            'From
            If Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release" OrElse Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug" Then
                objE.From = "yvette@yswconsulting.com"
            Else
                objE.From = "billing@gatewaypersonnel.com"
            End If

            'To - need to use objC.BillAddress
            objE.MailTo = strEmail

            'Subject
            objE.Subject = "Invoice No. " & CStr(lngInvNumber) & " From Gateway Group Personnel Is Now Online."

            'Body ---------------------------------------------------------
            sb.Append("Invoice No. " & CStr(lngInvNumber) & " From Gateway Group Personnel Is Now Online.")
            sb.AppendLine()
            sb.AppendLine()
            sb.AppendLine("In order to view your invoice, please log into your account at www.gatewaypersonnel.com and select View Invoices.")
            sb.AppendLine()
            sb.AppendLine()
            sb.Append("We thank you for your business.")
            sb.AppendLine()
            sb.AppendLine()
            sb.Append("Gateway Group Personnel.")

            objE.Body = sb.ToString
            'Mail Server
            If Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release" OrElse Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug" Then
                objE.MailServer = "mail.yswconsulting.com"
            Else
                objE.MailServer = "10.0.0.48"
            End If


            SendEmail(objE)

            objE = Nothing


        End Sub

        Public Sub SendUnpaidInvoiceReminder(ByVal strEmail As String, ByVal lngInvnumber As Long)
            'look up bill attn email address and invoice info with clcode and invnumber
            Dim objE As EmailInfo = New EmailInfo
            Dim sb As StringBuilder = New StringBuilder

            'From
            If Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release" OrElse Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug" Then
                objE.From = "yvette@yswconsulting.com"
            Else
                objE.From = "billing@gatewaypersonnel.com"
            End If

            'To - need to use objC.BillAddress
            objE.MailTo = strEmail

            'Subject
            objE.Subject = "Invoice Reminder"

            'Body ---------------------------------------------------------
            sb.Append("This is just a friendly reminder that we have not yet received payment on Invoice # " & lngInvnumber & ".")
            sb.AppendLine()
            sb.AppendLine()
            sb.Append("If you have a Gateway Group Personnel website account, you can log in at www.gatewaypersonnel.com to view your invoice or download it in PDF format.")
            sb.AppendLine()
            sb.AppendLine()
            sb.Append("We look forward to hearing from you soon and, as always, we appreciate your business!")
            sb.AppendLine()
            sb.AppendLine()
            sb.Append("Gateway Group Personnel.")

            objE.Body = sb.ToString
            'Mail Server
            If Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release" OrElse Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug" Then
                objE.MailServer = "smtp.bizmail.yahoo.com"
            Else
                objE.MailServer = "10.0.0.48"
            End If


            SendEmail(objE)

            objE = Nothing

        End Sub

        Public Sub SendTimeSlipReadyOnlineNotification(ByVal lngBackOfficeID As Long)

            'look up time slip manager email with Back Office ID

        End Sub

        Public Sub SendUnsignedTimeSlipReminder(ByVal strEmail As String)

            Dim objE As EmailInfo = New EmailInfo
            Dim sb As StringBuilder = New StringBuilder

            'From
            If Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release" OrElse Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug" Then
                objE.From = "yvette@yswconsulting.com"
            Else
                objE.From = "billing@gatewaypersonnel.com"
            End If

            'To - need to use objC.BillAddress
            objE.MailTo = strEmail

            'Subject
            objE.Subject = "A Time Slip Requires Your Approval"

            'Body ---------------------------------------------------------
            sb.Append("Please log into www.gatewaypersonnel.com and check the Clients Area for pending time slips. The employee's check cannot be processed until the time slip is approved.")
            sb.AppendLine()
            sb.AppendLine()
            sb.Append("We appreciate your business.")
            sb.AppendLine()
            sb.AppendLine()
            sb.Append("Gateway Group Personnel.")

            objE.Body = sb.ToString
            'Mail Server
            If Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release" OrElse Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug" Then
                objE.MailServer = "smtp.bizmail.yahoo.com"
            Else
                objE.MailServer = "10.0.0.48"
            End If


            SendEmail(objE)

            objE = Nothing

        End Sub

    End Class
End Namespace
