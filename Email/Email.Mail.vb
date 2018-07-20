Imports System
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Collections
Imports System.Collections.Specialized
Namespace Email.SendEmail
    ''' <summary>
    ''' 邮件内容
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Mail
        Protected _Log As New LogUtility.Logger("SendMail_TCPClient")
        Private _Sender As System.Net.Mail.MailAddress  ' String = Nothing
        Private _Receivers As New List(Of System.Net.Mail.MailAddress) 'System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection()
        Private _Priority As Integer = 3
        Private _Subject As String = ""
        Private _XMailer As String = ""
        Private _Aattachments As System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection()
        Private _MailEncoding As MailEncodings = MailEncodings.GB2312
        Private _MailType As MailTypes = MailTypes.Html
        Private _MailBody() As Byte = Nothing
        Private _Boundary As String = Nothing
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
        Public Property Priority() As Integer
            Get
                Return _Priority
            End Get
            Set(ByVal value As Integer)
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
        Public Property MailEncoding() As MailEncodings
            Get
                Return _MailEncoding
            End Get
            Set(ByVal value As MailEncodings)
                _MailEncoding = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置邮件格式
        ''' </summary>
        Public Property MailType() As MailTypes
            Get
                Return _MailType
            End Get
            Set(ByVal value As MailTypes)
                _MailType = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置邮件正文
        ''' </summary>
        Public Property MailBody() As Byte()
            Get
                Return _MailBody
            End Get
            Set(ByVal value As Byte())
                _MailBody = value
            End Set
        End Property

        ''' <summary>
        ''' 邮件编码
        ''' </summary>
        Public Enum MailEncodings
            GB2312
            ASCII
            Unicode
            UTF8
        End Enum

        ''' <summary>
        ''' 邮件格式
        ''' </summary>
        Public Enum MailTypes
            Html
            Text
        End Enum

        ''' <summary>
        ''' smtp服务器的验证方式
        ''' </summary>
        Public Enum SmtpValidateTypes
            ''' <summary>
            ''' 不需要验证
            ''' </summary>
            None
            ''' <summary>
            ''' 通用的auth login验证
            ''' </summary>
            Login
            ''' <summary>
            ''' 通用的auth plain验证
            ''' </summary>
            Plain
            ''' <summary>
            ''' CRAM-MD5验证
            ''' </summary>
            CRAMMD5
        End Enum

        ''' <summary>
        ''' 获取或设置邮件ID (里面需要有TemplateID EmailListID)
        ''' </summary>
        Public Property Boundary() As String
            Get
                Return _Boundary
            End Get
            Set(ByVal value As String)
                _Boundary = value
            End Set
        End Property
    End Class
End Namespace

