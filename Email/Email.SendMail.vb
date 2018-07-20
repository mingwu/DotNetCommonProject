Namespace Email
    Public Class SendMail
#Region "Fields"
        Dim _Log As New LogUtility.Logger("SendMail")
        Private _Sender As System.Net.Mail.MailAddress  ' String = Nothing
        Private _Receivers As New List(Of System.Net.Mail.MailAddress) 'System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection()
        Private _Priority As Net.Mail.MailPriority = Net.Mail.MailPriority.Normal
        Private _Subject As String = ""
        Private _XMailer As String = ""
        Private _Aattachments As System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection()
        Private _MailEncoding As System.Text.Encoding 'MailEncodings = MailEncodings.GB2312
        Private _IsBodyHtml As Boolean ' MailTypes = MailTypes.Html
        Private _MailBody As String
        Private _SmtpServer As String
        Private _UserName As String
        Private _Password As String
#End Region
#Region "Propertes"
        ''' <summary>
        ''' 获取或设置Smtp密码
        ''' </summary>
        Public Property Password() As String
            Get
                Return _Password
            End Get
            Set(ByVal value As String)
                _Password = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置Smtp用户名
        ''' </summary>
        Public Property UserName() As String
            Get
                Return _UserName
            End Get
            Set(ByVal value As String)
                _UserName = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置Smtp服务器地址
        ''' </summary>
        Public Property SmtpServer() As String
            Get
                Return _SmtpServer
            End Get
            Set(ByVal value As String)
                _SmtpServer = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置发件人
        ''' </summary>
        Public Property Sender() As System.Net.Mail.MailAddress
            Get
                Return _Sender
            End Get
            Set(ByVal value As System.Net.Mail.MailAddress)
                _Sender = value
            End Set
        End Property

        ''' <summary>
        ''' 获取收件人地址集合
        ''' </summary>
        Public Property Receivers() As List(Of System.Net.Mail.MailAddress) 'System.Collections.Specialized.StringCollection
            Get
                Return _Receivers
            End Get
            Set(ByVal value As List(Of System.Net.Mail.MailAddress)) 'System.Collections.Specialized.StringCollection)
                _Receivers = value
            End Set
        End Property

        ''' <summary>
        ''' 邮件优先级 [默认为3]
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Priority() As Net.Mail.MailPriority
            Get
                Return _Priority
            End Get
            Set(ByVal value As Net.Mail.MailPriority)
                _Priority = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置邮件主题
        ''' </summary>
        Public Property Subject() As String
            Get
                Return _Subject
            End Get
            Set(ByVal value As String)
                _Subject = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置邮件传送者
        ''' </summary>
        Public Property XMailer() As String
            Get
                Return _XMailer
            End Get
            Set(ByVal value As String)
                _XMailer = value
            End Set
        End Property

        ''' <summary>
        ''' 获取附件列表
        ''' </summary>
        Public ReadOnly Property Attachments() As System.Collections.Specialized.StringCollection
            Get
                Return _Aattachments
            End Get
            'Set(ByVal value As StringCollection())

            'End Set
        End Property

        ''' <summary>
        ''' 获取或设置邮件的编码方式
        ''' </summary>
        Public Property MailEncoding() As System.Text.Encoding ' MailEncodings
            Get
                Return _MailEncoding
            End Get
            Set(ByVal value As System.Text.Encoding) ' MailEncodings)
                _MailEncoding = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置邮件格式
        ''' </summary>
        Public Property IsBodyHtml() As Boolean ' MailTypes
            Get
                Return _IsBodyHtml
            End Get
            Set(ByVal value As Boolean) 'MailTypes)
                _IsBodyHtml = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置邮件正文
        ''' </summary>
        Public Property MailBody() As String
            Get
                Return _MailBody
            End Get
            Set(ByVal value As String)
                _MailBody = value
            End Set
        End Property

        '''' <summary>
        '''' 邮件编码
        '''' </summary>
        'Public Enum MailEncodings
        '    GB2312 = System.Text.Encoding.UTF8
        '    ASCII
        '    Unicode
        '    UTF8
        'End Enum

        '''' <summary>
        '''' 邮件格式
        '''' </summary>
        'Public Enum MailTypes
        '    Html
        '    Text
        'End Enum

        '''' <summary>
        '''' smtp服务器的验证方式
        '''' </summary>
        'Public Enum SmtpValidateTypes
        '    ''' <summary>
        '    ''' 不需要验证
        '    ''' </summary>
        '    None
        '    ''' <summary>
        '    ''' 通用的auth login验证
        '    ''' </summary>
        '    Login
        '    ''' <summary>
        '    ''' 通用的auth plain验证
        '    ''' </summary>
        '    Plain
        '    ''' <summary>
        '    ''' CRAM-MD5验证
        '    ''' </summary>
        '    CRAMMD5
        'End Enum
#End Region

        Public Sub SendMail()
            Dim _SmtpClient As Net.Mail.SmtpClient
            Dim _Message As Net.Mail.MailMessage
            Try
                _SmtpClient = New Net.Mail.SmtpClient(_SmtpServer, 25)
                _SmtpClient.UseDefaultCredentials = True
                _SmtpClient.Credentials = New Net.NetworkCredential(_UserName, _Password)
                _Message = New Net.Mail.MailMessage()
                _Message.From = _Sender
                For i As Integer = 0 To _Receivers.Count - 1
                    _Message.To.Add(_Receivers.Item(i))
                    _Log.Info("Sending an e-mail message to " + _Message.To(i).Address.ToString + " using the SMTP host " + _SmtpClient.Host.ToString)
                Next

                _Message.Subject = _Subject
                _Message.Body = _MailBody
                _Message.Priority = _Priority
                _Message.IsBodyHtml = _IsBodyHtml
                _Message.SubjectEncoding = _MailEncoding
                _Message.BodyEncoding = _MailEncoding ' System.Text.Encoding.UTF8

                _SmtpClient.Send(_Message)
                _Log.Info("Send OK")
            Catch ex As Exception
                'Catch ex As Net.Mail.SmtpFailedRecipientsException
                'For i As Integer = 0 To ex.InnerExceptions.Length
                '    Dim _Status As Net.Mail.SmtpStatusCode = ex.InnerExceptions(i).StatusCode
                _Log.Error(ex.Message.ToString)
                'Next
                'Dim _Status As Net.Mail.SmtpStatusCode = ex.StatusCode
                ' _Log.Error(ex.StatusCode.ToString + "   " + ex.FailedRecipient.ToString + "   " + ex.ToString)
            Finally

            End Try

        End Sub
    End Class
End Namespace

