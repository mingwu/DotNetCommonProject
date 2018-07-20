Imports System
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Collections
Imports System.Collections.Specialized
Namespace Email.SendEmail
    ''' <summary>
    ''' 邮件发送类
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Smtp
        Inherits Email.SendEmail.Mail
#Region "Member Fields"
        ''' <summary>
        ''' 连接对象
        ''' </summary>
        Private _TcpClient As System.Net.Sockets.TcpClient
        ''' <summary>
        ''' 网络流
        ''' </summary>
        Private _NetworkStream As System.Net.Sockets.NetworkStream
        ''' <summary>
        ''' 错误的代码字典
        ''' </summary>
        Private _ErrorCodes As System.Collections.Specialized.StringDictionary = New System.Collections.Specialized.StringDictionary()
        ''' <summary>
        ''' 操作执行成功后的响应代码字典
        ''' </summary>
        Private _RightCodes As System.Collections.Specialized.StringDictionary = New System.Collections.Specialized.StringDictionary()
        ''' <summary>
        ''' 执行过程中错误的消息
        ''' </summary>
        Private _ErrorMessage As String = ""
        ''' <summary>
        ''' 执行过程中成功的消息
        ''' </summary>
        Private _SuccessMessage As String = ""
        ''' <summary>
        ''' 记录操作日志
        ''' </summary>
        Private _Logs As String = ""
        ''' <summary>
        ''' 主机登陆的验证方式
        ''' </summary>
        Private _ValidateTypes As System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection()
        ''' <summary>
        ''' 换行常数
        ''' </summary>
        Private Const _CRLF As String = vbCrLf
        Private _ServerName As String = System.Configuration.ConfigurationSettings.AppSettings("LocalDomainName") '"smtp"
        Private _LogPath As String = Nothing
        Private _UserID As String = Nothing
        Private _Password As String = Nothing
        Private _MailEncodingName As String = "GB2312"
        Private _IsSendComplete As Boolean = False
        Private _SmtpValidateType As Email.SendEmail.Mail.SmtpValidateTypes = Email.SendEmail.Mail.SmtpValidateTypes.Login
        Private _StatusCode As String
        Private _Server As String = Nothing
        Private _Port As Integer
        Private _endPoint As IPEndPoint = Nothing
        Private _IP As String = Nothing
#End Region

#Region "Propertys"
        ''' <summary>
        ''' 获取最后一此程序执行中的错误消息
        ''' </summary>
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _ErrorMessage
            End Get
            'Set(ByVal value As String)

            'End Set
        End Property

        ''' <summary>
        ''' 获取最后一此程序执行中的成功消息
        ''' </summary>
        Public ReadOnly Property SuccessMessage() As String
            Get
                Return _SuccessMessage
            End Get
            'Set(ByVal value As String)

            'End Set
        End Property

        '''' <summary>
        '''' 获取或设置日志输出路径
        '''' </summary>
        'Public Property LogPath() As String
        '    Get
        '        Return _LogPath
        '    End Get
        '    Set(ByVal value As String)
        '        _LogPath = value
        '    End Set
        'End Property

        ''' <summary>
        ''' 获取或设置登陆smtp服务器的帐号
        ''' </summary>
        Public Property UserID() As String
            Get
                Return _UserID
            End Get
            Set(ByVal value As String)
                _UserID = value
            End Set
        End Property

        ''' <summary>
        ''' 获取或设置登陆smtp服务器的密码
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
        ''' 获取或设置要使用登陆Smtp服务器的验证方式
        ''' </summary>
        Public Property SmtpValidateType() As SmtpValidateTypes
            Get
                Return _SmtpValidateType
            End Get
            Set(ByVal value As SmtpValidateTypes)
                _SmtpValidateType = value
            End Set
        End Property

        ''' <summary>
        ''' 获取状态码
        ''' </summary>
        Public ReadOnly Property StatusCode() As String
            Get
                If Not _ErrorCodes(_StatusCode) = "" Then
                    Return _StatusCode + " " + _ErrorCodes(_StatusCode)
                ElseIf Not _RightCodes(_StatusCode) = "" Then
                    Return _StatusCode + " " + _RightCodes(_StatusCode)
                Else
                    Return _StatusCode
                End If
            End Get
            'Set(ByVal value As String)
            '    _StatusCode = value
            'End Set
        End Property

        ''' <summary>
        ''' 获取或设置本机名称
        ''' </summary>
        Public Property ServerName() As String
            Get
                Return _ServerName
            End Get
            Set(ByVal value As String)
                _ServerName = value
            End Set
        End Property

        Public Property Port() As Integer
            Get
                Return _Port
            End Get
            Set(ByVal value As Integer)
                _Port = value
            End Set
        End Property

        Public Property Server() As String
            Get
                Return _Server
            End Get
            Set(ByVal value As String)
                _Server = value
            End Set
        End Property

        Public Property IP() As String
            Get
                Return _IP
            End Get
            Set(ByVal value As String)
                _IP = value
            End Set
        End Property

#End Region

#Region "Construct Functions"

        Public Sub New()

        End Sub

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="server">主机名</param>
        ''' <param name="port">端口</param>
        Public Sub New(ByVal Server As String, ByVal Port As Integer)
            _TcpClient = New TcpClient(Server, Port)
            _Server = Server
            _Port = Port
            _NetworkStream = _TcpClient.GetStream()
            _ServerName = Server
            initialFields()
        End Sub

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="ip">主机ip</param>
        Public Sub New(ByVal IP As System.Net.IPEndPoint)
            _TcpClient = New TcpClient(IP)
            _NetworkStream = _TcpClient.GetStream()
            _ServerName = IP.ToString()
            initialFields()
        End Sub
#End Region

#Region "Methods"
        Private Sub ReConnect()
            If IsNothing(_Server) Then
                _TcpClient = New TcpClient(_endPoint)
            Else
                _TcpClient = New TcpClient(_Server, _Port)
            End If
            _NetworkStream = _TcpClient.GetStream()
            _ServerName = _Server
            initialFields()
        End Sub

        Private Sub initialFields() '初始化连接
            _Logs = "================" + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + "===============" + _CRLF
            _ErrorCodes.Clear()
            ''*****************************************************************
            ''错误的状态码
            ''*****************************************************************
            _ErrorCodes.Add("421", "服务未就绪,关闭传输通道")
            _ErrorCodes.Add("432", "需要一个密码转换")
            _ErrorCodes.Add("450", "要求的邮件操作未完成,邮箱不可用(如:邮箱忙)")
            _ErrorCodes.Add("451", "放弃要求的操作,要求的操作未执行")
            _ErrorCodes.Add("452", "系统存储不足,要求的操作未完成")
            _ErrorCodes.Add("454", "临时的认证失败")
            _ErrorCodes.Add("500", "邮箱地址错误")
            _ErrorCodes.Add("501", "参数格式错误")
            _ErrorCodes.Add("502", "命令不可实现")
            _ErrorCodes.Add("503", "命令的次序不正确")
            _ErrorCodes.Add("504", "命令参数不可实现")
            _ErrorCodes.Add("530", "需要认证")
            _ErrorCodes.Add("534", "认证机制过于简单")
            _ErrorCodes.Add("538", "当前请求的认证机制需要加密")
            _ErrorCodes.Add("550", "当前的邮件操作未完成,邮箱不可用(如：邮箱未找到或邮箱不能用)")
            _ErrorCodes.Add("551", "用户非本地,请尝试<forward-path>")
            _ErrorCodes.Add("552", "过量的存储分配,制定的操作未完成")
            _ErrorCodes.Add("553", "邮箱名不可用,如:邮箱地址的格式错误")
            _ErrorCodes.Add("554", "传送失败")
            _ErrorCodes.Add("535", "用户身份验证失败")
            ''****************************************************************
            ''操作执行成功后的状态码
            ''****************************************************************
            _RightCodes.Clear()
            _RightCodes.Add("220", "服务就绪")
            _RightCodes.Add("221", "服务关闭传输通道")
            _RightCodes.Add("235", "验证成功")
            _RightCodes.Add("250", "要求的邮件操作完成")
            _RightCodes.Add("251", "非本地用户,将转发向<forward-path>")
            _RightCodes.Add("334", "服务器响应验证Base64字符串")
            _RightCodes.Add("354", "开始邮件输入,以<_CRLF>.<_CRLF>结束")
            ''读取系统回应
            Dim _Reader As System.IO.StreamReader = New System.IO.StreamReader(_NetworkStream)
            _Logs += _Reader.ReadLine() + _CRLF
        End Sub

        ''' <summary>
        ''' 向SMTP发送命令
        ''' </summary>
        ''' <param name="Cmd"></param>
        '''  <param name="IsMailData"></param>
        Private Function SendCommand(ByVal Cmd As String, ByVal IsMailData As Boolean) As String
            If Not IsNothing(Cmd) And Not String.IsNullOrEmpty(Cmd.Trim()) Then
                Dim _Cmd_b() As Byte = Nothing
                If Not IsMailData Then ''不是邮件数据
                    Cmd += _CRLF
                End If
                _Logs += Cmd
                ''开始写入邮件数据
                If Not IsMailData Then
                    _Cmd_b = System.Text.Encoding.ASCII.GetBytes(Cmd)
                    _NetworkStream.Write(_Cmd_b, 0, _Cmd_b.Length)
                Else
                    _Cmd_b = System.Text.Encoding.GetEncoding(_MailEncodingName).GetBytes(Cmd)
                    '_NetworkStream.BeginWrite(_Cmd_b, 0, _Cmd_b.Length, New AsyncCallback(AddressOf AsyncWriterCallBack), Nothing)
                    _NetworkStream.Write(_Cmd_b, 0, _Cmd_b.Length)

                End If
                '_NetworkStream.BeginRead(_Cmd_b, 0, _Cmd_b.Length, New AsyncCallback(AddressOf AsyncReaderCallBack), Nothing)
                '_NetworkStream.Read(_Cmd_b, 0, _Cmd_b.Length)
                ''读取服务器响应
                Dim _Reader As System.IO.StreamReader = New System.IO.StreamReader(_NetworkStream)
                Dim _Response As String = _Reader.ReadLine()
                _Log.Info(_Mail.Receivers(0).Address.ToString + "[Command]" + Cmd)
                _Log.Debug(_Mail.Receivers(0).Address.ToString + "[检测返回状态]" + _Response)
                _Logs += _Response + _CRLF
                ''检查状态码
                _StatusCode = _Response.Substring(0, 3)
                Dim _IsExist As Boolean = True
                Dim _IsRightCode As Boolean = True
                Dim _IsErrorCode As Boolean = True
                For Each _Err As String In _ErrorCodes.Keys
                    If _StatusCode = _Err Then
                        _IsExist = False
                        _IsRightCode = False
                        _IsErrorCode = True
                        Exit For
                    End If
                Next
                For Each _Right As String In _RightCodes.Keys
                    If _StatusCode = _Right Then
                        _IsExist = False
                        _IsRightCode = True
                        _IsErrorCode = False
                        Exit For
                    End If
                Next
                ''根据状态码来处理下一步的动作
                If _IsExist Then ''不是合法的SMTP主机
                    SetError("不是合法的SMTP主机,或服务器拒绝服务")
                ElseIf _IsErrorCode Then '命令没能成功执行
                    SetError(_StatusCode + ":" + _ErrorCodes(_StatusCode))
                ElseIf _IsRightCode Then ' 命令成功执行
                    SetSuccess(_StatusCode + ":" + _RightCodes(_StatusCode))
                    '_ErrorMessage = ""
                End If
                Return _Response
            Else
                Return ""
            End If
        End Function

        ''' <summary>
        ''' 通过auth login方式登陆smtp服务器
        ''' </summary>
        Private Sub LandingByLogin()
            Dim _Base64UserID As String = ConvertBase64String(_UserID, "ASCII")
            Dim _Base64Pass As String = ConvertBase64String(_Password, "ASCII")
            ''握手
            SendCommand("EHLO " + _ServerName, False) ' 
            ''开始登陆
            SendCommand("AUTH LOGIN", False)
            SendCommand(_Base64UserID, False)
            SendCommand(_Base64Pass, False)
        End Sub

        ''' <summary>
        ''' 通过auth plain方式登陆服务器
        ''' </summary>
        Private Sub LandingByPlain()
            Dim _NULL As String = 0.ToString()
            Dim _LoginStr As String = _NULL + _UserID + _NULL + _Password
            Dim _Base64LoginStr As String = ConvertBase64String(_LoginStr, "ASCII")
            ''握手
            SendCommand("EHLO " + _ServerName, False)
            ''登陆
            SendCommand(_Base64LoginStr, False)
        End Sub

        Private _Mail As Email.SendEmail.Mail
        '<Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)> _
        ''' <summary>
        ''' 发送邮件
        ''' </summary>
        Public Sub SendMail(ByVal Mail As Email.SendEmail.Mail) 'As Boolean
            '            Dim _SendEmailParameter As Mail = Mail(0)
            _Mail = Mail
            _Log.Fatal(_TcpClient.Connected)
            If Not _TcpClient.Connected Then
                ReConnect()
            End If

            Dim _IsSended As Boolean = True
            If IsNothing(Mail.Boundary) Then
                Mail.Boundary = Guid.NewGuid().ToString
            End If
            Try
                ''检测发送邮件的必要条件
                If Mail.Sender.Equals(Nothing) Then
                    SetError("没有设置发信人")
                    _IsSended = False
                    Exit Sub
                End If
                If Mail.Receivers.Count = 0 Then
                    SetError("至少要有一个收件人")
                    _IsSended = False
                    Exit Sub
                End If
                If _SmtpValidateType <> SmtpValidateTypes.None Then
                    If String.IsNullOrEmpty(_UserID) Or String.IsNullOrEmpty(_Password) Then
                        SetError("当前设置需要smtp验证,但是没有给出登陆帐号")
                        _IsSended = False
                        Exit Sub
                    End If
                End If
                ''开始登陆
                Select Case _SmtpValidateType
                    Case SmtpValidateTypes.None
                        SendCommand("EHLO " + _ServerName, False)
                    Case SmtpValidateTypes.Login
                        LandingByLogin()
                    Case SmtpValidateTypes.Plain
                        LandingByPlain()
                    Case Else
                End Select
                ''初始化邮件会话(对应SMTP命令mail) 需要返回邮箱地址　这个地方修改修改  随机从一堆接收反馈邮件地址中得到一个地址写入!!!!!!
                SendCommand("MAIL FROM:<" + Mail.Sender.Address + ">", False)
                ''标识收件人(对应SMTP命令Rcpt)
                For Each _Receive As System.Net.Mail.MailAddress In Mail.Receivers
                    SendCommand("RCPT TO:<" + _Receive.Address + ">", False)
                    If _StatusCode <> "250" Then
                        _IsSended = False
                        Exit Sub
                    End If
                Next
                ''标识开始输入邮件内容(Data)
                SendCommand("DATA", False)
                ''开始编写邮件内容
                Dim _Message As String = "mime-version: 1.0" + _CRLF
                ' _Message += "From: " + Mail.Sender.DisplayName + "<" + Mail.Sender.Address + ">" + _CRLF
                _Message += "From: " + Mail.Sender.ToString + _CRLF
                Dim i As Integer = 0
                For Each _Receive As System.Net.Mail.MailAddress In Mail.Receivers
                    '_Message += "To: " + _Receive.DisplayName + "<" + _Receive.Address + ">" + _CRLF
                    _Message += "To: " + _Receive.ToString + _CRLF
                Next
                _Message += "Date: " + DateTime.Now.ToString("dd MMM yyyy HH:mm:ss zz00", System.Globalization.CultureInfo.CreateSpecificCulture("en-us")) + _CRLF ' + DateTime.Now.ToString("r") 
                _Message += "Subject:" + Mail.Subject + _CRLF
                '_Message += "Reply-To:" + Mail.Sender.DisplayName + "<" + Mail.Sender.Address + ">" + _CRLF
                '_Message += "Reply-To:" + Mail.Sender.Address.ToString + _CRLF
                ' _Message += "X-Priority:" + Mail.Priority.ToString + _CRLF
                ' _Message += "X-MSMail-Priority: Normal" + _CRLF
                '_Message += "X-mailer:" + Mail.XMailer + _CRLF
                ' _Message += "X-MimeOLE: Produced By Microsoft MimeOLE V6.00.2900.2869" + _CRLF
                If Mail.Attachments.Count = 0 Then ''没有附件
                    If Mail.MailType = MailTypes.Text Then ''文本格式
                        _Message += "Content-Type:text/plain;" + _CRLF + " ".PadRight(8, " ") + "charset=""" + Mail.MailEncoding.ToString + """" + _CRLF
                        _Message += "Content-Transfer-Encoding:base64" + _CRLF + _CRLF
                        If Not Mail.MailBody Is Nothing Then
                            _Message += Convert.ToBase64String(Mail.MailBody, 0, Mail.MailBody.Length) + _CRLF + _CRLF
                        End If
                        _Message += _CRLF + "." + _CRLF
                    Else 'Html格式alertnative
                        _Message += "Content-Type:multipart/alertnative;" + _CRLF + " ".PadRight(8, " ") + " boundary" + "=""=====" + Mail.Boundary + "=====""" + _CRLF + _CRLF + _CRLF
                        _Message += "This is a multi-part _Message in MIME format" + _CRLF + _CRLF
                        _Message += "--=====" + Mail.Boundary + "=====" + _CRLF
                        _Message += "Content-Type:text/html;" + _CRLF + " ".PadRight(8, " ") + "charset=""" + Mail.MailEncoding.ToString() + """" + _CRLF
                        _Message += "Content-Transfer-Encoding:base64" + _CRLF + _CRLF
                        If Not Mail.MailBody Is Nothing Then
                            _Message += Convert.ToBase64String(Mail.MailBody, 0, Mail.MailBody.Length) + _CRLF + _CRLF
                        End If
                        _Message += _CRLF + "." + _CRLF
                    End If
                Else '有附件
                    '处理要在邮件中显示的每个附件的数据
                    Dim _AttatchmentDatas As System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection()
                    For Each path As String In Mail.Attachments
                        If Not File.Exists(path) Then
                            SetError("指定的附件没有找到" + path)
                        Else
                            Dim _File As FileInfo = New FileInfo(path)
                            Dim _FileStream As FileStream = New FileStream(path, FileMode.Open, FileAccess.Read)
                            If _FileStream.Length > CLng(Integer.MaxValue) Then
                                SetError("附件的大小超出了最大限制")
                            End If
                            Dim _FileByte(Int(_FileStream.Length)) As Byte
                            _FileStream.Read(_FileByte, 0, _FileByte.Length)
                            _FileStream.Close()
                            Dim _AttatchmentMailStr As String = "Content-Type:application/octet-stream;" + _CRLF + " ".PadRight(8, " ") + "name=" + "\"" + _File.Name + " \ "" + _CRLF
                            _AttatchmentMailStr += "Content-Transfer-Encoding:base64" + _CRLF
                            _AttatchmentMailStr += "Content-Disposition:attachment;" + _CRLF + " ".PadRight(8, " ") + "filename=" + "\"" + _File.Name + " \ "" + _CRLF + _CRLF
                            _AttatchmentMailStr += Convert.ToBase64String(_FileByte, 0, _FileByte.Length) + _CRLF + _CRLF
                            _AttatchmentDatas.Add(_AttatchmentMailStr)
                        End If
                    Next
                    '设置邮件信息
                    If Mail.MailType = MailTypes.Text Then '文本格式
                        _Message += "Content-Type:multipart/mixed;" + _CRLF + " ".PadRight(8, " ") + "boundary=""=====" + Mail.Boundary + "=====""" + _CRLF + _CRLF
                        _Message += "This is a multi-part _Message in MIME format." + _CRLF + _CRLF
                        _Message += "--=====" + Mail.Boundary + "=====" + _CRLF
                        _Message += "Content-Type:text/plain;" + _CRLF + " ".PadRight(8, " ") + "charset=""" + Mail.MailEncoding.ToString().ToLower() + """" + _CRLF
                        _Message += "Content-Transfer-Encoding:base64" + _CRLF + _CRLF
                        If Not Mail.MailBody Is Nothing Then
                            _Message += Convert.ToBase64String(Mail.MailBody, 0, Mail.MailBody.Length) + _CRLF + _CRLF + _CRLF + "." + _CRLF
                        End If
                        For Each s As String In _AttatchmentDatas
                            _Message += "--=====" + Mail.Boundary + "=====" + _CRLF + s + _CRLF + _CRLF
                        Next
                        _Message += "--=====" + Mail.Boundary + "=====--" + _CRLF + _CRLF + _CRLF + "." + _CRLF
                    Else 'Html格式
                        _Message += "Content-Type:multipart/mixed;" + _CRLF + " ".PadRight(8, " ") + " boundary" + "=""=====" + Mail.Boundary + "=====""" + _CRLF + _CRLF + _CRLF
                        _Message += "This is a multi-part _Message in MIME format" + _CRLF + _CRLF
                        _Message += "--=====" + Mail.Boundary + "=====" + _CRLF
                        _Message += "Content-Type:text/html;" + _CRLF + " ".PadRight(8, " ") + "charset=""" + Mail.MailEncoding.ToString().ToLower + """" + _CRLF
                        _Message += "Content-Transfer-Encoding:base64" + _CRLF + _CRLF
                        If Not Mail.MailBody Is Nothing Then
                            _Message += Convert.ToBase64String(Mail.MailBody, 0, Mail.MailBody.Length) + _CRLF + _CRLF
                        End If
                        For i = 0 To _AttatchmentDatas.Count - 1
                            _Message += "--=====" + Mail.Boundary + "=====" + _CRLF + _AttatchmentDatas(i) + _CRLF + _CRLF
                        Next
                        _Message += "--=====" + Mail.Boundary + "=====--" + _CRLF + _CRLF + _CRLF + "." + _CRLF
                    End If
                End If
                ''发送邮件数据
                _MailEncodingName = Mail.MailEncoding.ToString()
                SendCommand(_Message, True)
                If _IsSendComplete Then
                    SendCommand("QUIT", False)
                End If
            Catch ex As Exception
                _Log.Error(ex.ToString)
                _IsSended = False
            Finally
                Disconnect()
                Receivers.Clear()
                ''输出日志文件
                ' _Log.Debug(_Logs)
            End Try
            'Return _IsSended
        End Sub

        ''' <summary>
        ''' 异步写入数据
        ''' </summary>
        ''' <param name="result"></param>
        Private Sub AsyncWriterCallBack(ByVal Result As System.IAsyncResult)
            If Result.IsCompleted Then
                _IsSendComplete = True
            End If
        End Sub

        ''' <summary>
        ''' 异步读取数据
        ''' </summary>
        ''' <param name="result"></param>
        Private Sub AsyncReaderCallBack(ByVal Result As System.IAsyncResult)
            If Result.IsCompleted Then
                _IsSendComplete = True
            End If
        End Sub

        ''' <summary>
        ''' 关闭连接
        ''' </summary> 
        Private Sub Disconnect()
            Try
                _NetworkStream.Close()
                _TcpClient.Close()
            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' 设置出现错误时的动作
        ''' </summary>
        ''' <param name="errorStr"></param>
        Private Sub SetError(ByVal ErrorStr As String)
            _ErrorMessage = ErrorStr
            '_Log.Error(ErrorStr)
            _Logs += "错误:" + _CRLF + _ErrorMessage + _CRLF + "【邮件处理动作中止】" + _CRLF
        End Sub

        ''' <summary>
        ''' 设置出现成功的动作
        ''' </summary>
        ''' <param name="SuccessStr"></param>
        Private Sub SetSuccess(ByVal SuccessStr As String)
            _SuccessMessage = SuccessStr
            '_Log.Info(SuccessStr)
            _Logs += "成功:" + _CRLF + _SuccessMessage + _CRLF + "【邮件处理动作成功】" + _CRLF
        End Sub

        ''' <summary>
        '''将字符串转换为base64
        ''' </summary>
        ''' <param name="str"></param>
        ''' <param name="encodingName"></param>
        ''' <returns></returns>
        Private Function ConvertBase64String(ByVal Str As String, ByVal EncodingName As String) As String
            Dim Str_b() As Byte = Encoding.GetEncoding(EncodingName).GetBytes(Str)
            Return Convert.ToBase64String(Str_b, 0, Str_b.Length)
        End Function
#End Region


    End Class
End Namespace

