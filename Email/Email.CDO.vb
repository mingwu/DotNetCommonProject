Namespace Email.CDO
    Public Class Email
        Dim iMsg
        Dim iConf
        Dim Flds

        Public Sub New()
            iMsg = CreateObject("cdo.message")
            iConf = CreateObject("cdo.configuration")
            Flds = iConf.Fields
            config()
        End Sub

        Private Sub config()
            Flds("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = System.Configuration.ConfigurationSettings.AppSettings("SmtpServer") '"doule.shepherddigital.com" '"mail.hgst-lucky.com.cn"
            Flds("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 '1 for cdoBasic,0 for all user
            Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = System.Configuration.ConfigurationSettings.AppSettings("SmtpUser") '"doule@doule.shepherddigital.com" '"webmaster@hgst-lucky.com.cn" '"cdosmtpuser"
            Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = System.Configuration.ConfigurationSettings.AppSettings("SmtpPassword") '"doule!@#" '"$mTPUser"
            Flds("http://schemas.microsoft.com/cdo/configuration/languagecode") = "gb2312"
            Flds("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
            Flds.Update()
            iMsg.Configuration = iConf
        End Sub

        Public Sub Sendmail(ByVal receiver, ByVal subject, ByVal body)
            iMsg.To = receiver
            iMsg.From = """doule"" <doule@doule.shepherddigital.com>" 'we need the email account here
            iMsg.Subject = subject
            iMsg.BodyPart.Charset = "gb2312"
            iMsg.HTMLBody = body
            iMsg.Send()
        End Sub
        Public Sub Sendmailurl(ByVal receiver, ByVal subject, ByVal url)
            Dim log As New LogUtility.Logger("email")
            Try
                iMsg.To = receiver
                iMsg.From = "doule@doule.shepherddigital.com" ',"""Hitachi"" <campaign@hbrchina.com>" 'we need the email account here
                iMsg.Subject = subject
                iMsg.BodyPart.Charset = "gb2312"
                iMsg.CreateMHTMLBody(url)
                iMsg.Send()
                log.Info("Sendmailurl: ok")
            Catch ex As Exception
                log.Error("Sendmailurl: " + ex.Message.ToString)
            End Try

        End Sub

        Public Sub Sendmailurl1(ByVal sender, ByVal receiver, ByVal subject, ByVal url)
            iMsg.To = receiver
            iMsg.From = sender 'we need the email account here
            iMsg.Subject = subject
            iMsg.BodyPart.Charset = "gb2312"
            iMsg.CreateMHTMLBody(url)
            iMsg.Send()
        End Sub
    End Class
End Namespace

