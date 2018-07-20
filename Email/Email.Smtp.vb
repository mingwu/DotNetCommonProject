Imports System
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Collections
Imports System.Collections.Specialized
Namespace Email.SendEmail
    ''' <summary>
    ''' �ʼ�������
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Smtp
        Inherits Email.SendEmail.Mail
#Region "Member Fields"
        ''' <summary>
        ''' ���Ӷ���
        ''' </summary>
        Private _TcpClient As System.Net.Sockets.TcpClient
        ''' <summary>
        ''' ������
        ''' </summary>
        Private _NetworkStream As System.Net.Sockets.NetworkStream
        ''' <summary>
        ''' ����Ĵ����ֵ�
        ''' </summary>
        Private _ErrorCodes As System.Collections.Specialized.StringDictionary = New System.Collections.Specialized.StringDictionary()
        ''' <summary>
        ''' ����ִ�гɹ������Ӧ�����ֵ�
        ''' </summary>
        Private _RightCodes As System.Collections.Specialized.StringDictionary = New System.Collections.Specialized.StringDictionary()
        ''' <summary>
        ''' ִ�й����д������Ϣ
        ''' </summary>
        Private _ErrorMessage As String = ""
        ''' <summary>
        ''' ִ�й����гɹ�����Ϣ
        ''' </summary>
        Private _SuccessMessage As String = ""
        ''' <summary>
        ''' ��¼������־
        ''' </summary>
        Private _Logs As String = ""
        ''' <summary>
        ''' ������½����֤��ʽ
        ''' </summary>
        Private _ValidateTypes As System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection()
        ''' <summary>
        ''' ���г���
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
        ''' ��ȡ���һ�˳���ִ���еĴ�����Ϣ
        ''' </summary>
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _ErrorMessage
            End Get
            'Set(ByVal value As String)

            'End Set
        End Property

        ''' <summary>
        ''' ��ȡ���һ�˳���ִ���еĳɹ���Ϣ
        ''' </summary>
        Public ReadOnly Property SuccessMessage() As String
            Get
                Return _SuccessMessage
            End Get
            'Set(ByVal value As String)

            'End Set
        End Property

        '''' <summary>
        '''' ��ȡ��������־���·��
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
        ''' ��ȡ�����õ�½smtp���������ʺ�
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
        ''' ��ȡ�����õ�½smtp������������
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
        ''' ��ȡ������Ҫʹ�õ�½Smtp����������֤��ʽ
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
        ''' ��ȡ״̬��
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
        ''' ��ȡ�����ñ�������
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
        ''' ���캯��
        ''' </summary>
        ''' <param name="server">������</param>
        ''' <param name="port">�˿�</param>
        Public Sub New(ByVal Server As String, ByVal Port As Integer)
            _TcpClient = New TcpClient(Server, Port)
            _Server = Server
            _Port = Port
            _NetworkStream = _TcpClient.GetStream()
            _ServerName = Server
            initialFields()
        End Sub

        ''' <summary>
        ''' ���캯��
        ''' </summary>
        ''' <param name="ip">����ip</param>
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

        Private Sub initialFields() '��ʼ������
            _Logs = "================" + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + "===============" + _CRLF
            _ErrorCodes.Clear()
            ''*****************************************************************
            ''�����״̬��
            ''*****************************************************************
            _ErrorCodes.Add("421", "����δ����,�رմ���ͨ��")
            _ErrorCodes.Add("432", "��Ҫһ������ת��")
            _ErrorCodes.Add("450", "Ҫ����ʼ�����δ���,���䲻����(��:����æ)")
            _ErrorCodes.Add("451", "����Ҫ��Ĳ���,Ҫ��Ĳ���δִ��")
            _ErrorCodes.Add("452", "ϵͳ�洢����,Ҫ��Ĳ���δ���")
            _ErrorCodes.Add("454", "��ʱ����֤ʧ��")
            _ErrorCodes.Add("500", "�����ַ����")
            _ErrorCodes.Add("501", "������ʽ����")
            _ErrorCodes.Add("502", "�����ʵ��")
            _ErrorCodes.Add("503", "����Ĵ�����ȷ")
            _ErrorCodes.Add("504", "�����������ʵ��")
            _ErrorCodes.Add("530", "��Ҫ��֤")
            _ErrorCodes.Add("534", "��֤���ƹ��ڼ�")
            _ErrorCodes.Add("538", "��ǰ�������֤������Ҫ����")
            _ErrorCodes.Add("550", "��ǰ���ʼ�����δ���,���䲻����(�磺����δ�ҵ������䲻����)")
            _ErrorCodes.Add("551", "�û��Ǳ���,�볢��<forward-path>")
            _ErrorCodes.Add("552", "�����Ĵ洢����,�ƶ��Ĳ���δ���")
            _ErrorCodes.Add("553", "������������,��:�����ַ�ĸ�ʽ����")
            _ErrorCodes.Add("554", "����ʧ��")
            _ErrorCodes.Add("535", "�û������֤ʧ��")
            ''****************************************************************
            ''����ִ�гɹ����״̬��
            ''****************************************************************
            _RightCodes.Clear()
            _RightCodes.Add("220", "�������")
            _RightCodes.Add("221", "����رմ���ͨ��")
            _RightCodes.Add("235", "��֤�ɹ�")
            _RightCodes.Add("250", "Ҫ����ʼ��������")
            _RightCodes.Add("251", "�Ǳ����û�,��ת����<forward-path>")
            _RightCodes.Add("334", "��������Ӧ��֤Base64�ַ���")
            _RightCodes.Add("354", "��ʼ�ʼ�����,��<_CRLF>.<_CRLF>����")
            ''��ȡϵͳ��Ӧ
            Dim _Reader As System.IO.StreamReader = New System.IO.StreamReader(_NetworkStream)
            _Logs += _Reader.ReadLine() + _CRLF
        End Sub

        ''' <summary>
        ''' ��SMTP��������
        ''' </summary>
        ''' <param name="Cmd"></param>
        '''  <param name="IsMailData"></param>
        Private Function SendCommand(ByVal Cmd As String, ByVal IsMailData As Boolean) As String
            If Not IsNothing(Cmd) And Not String.IsNullOrEmpty(Cmd.Trim()) Then
                Dim _Cmd_b() As Byte = Nothing
                If Not IsMailData Then ''�����ʼ�����
                    Cmd += _CRLF
                End If
                _Logs += Cmd
                ''��ʼд���ʼ�����
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
                ''��ȡ��������Ӧ
                Dim _Reader As System.IO.StreamReader = New System.IO.StreamReader(_NetworkStream)
                Dim _Response As String = _Reader.ReadLine()
                _Log.Info(_Mail.Receivers(0).Address.ToString + "[Command]" + Cmd)
                _Log.Debug(_Mail.Receivers(0).Address.ToString + "[��ⷵ��״̬]" + _Response)
                _Logs += _Response + _CRLF
                ''���״̬��
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
                ''����״̬����������һ���Ķ���
                If _IsExist Then ''���ǺϷ���SMTP����
                    SetError("���ǺϷ���SMTP����,��������ܾ�����")
                ElseIf _IsErrorCode Then '����û�ܳɹ�ִ��
                    SetError(_StatusCode + ":" + _ErrorCodes(_StatusCode))
                ElseIf _IsRightCode Then ' ����ɹ�ִ��
                    SetSuccess(_StatusCode + ":" + _RightCodes(_StatusCode))
                    '_ErrorMessage = ""
                End If
                Return _Response
            Else
                Return ""
            End If
        End Function

        ''' <summary>
        ''' ͨ��auth login��ʽ��½smtp������
        ''' </summary>
        Private Sub LandingByLogin()
            Dim _Base64UserID As String = ConvertBase64String(_UserID, "ASCII")
            Dim _Base64Pass As String = ConvertBase64String(_Password, "ASCII")
            ''����
            SendCommand("EHLO " + _ServerName, False) ' 
            ''��ʼ��½
            SendCommand("AUTH LOGIN", False)
            SendCommand(_Base64UserID, False)
            SendCommand(_Base64Pass, False)
        End Sub

        ''' <summary>
        ''' ͨ��auth plain��ʽ��½������
        ''' </summary>
        Private Sub LandingByPlain()
            Dim _NULL As String = 0.ToString()
            Dim _LoginStr As String = _NULL + _UserID + _NULL + _Password
            Dim _Base64LoginStr As String = ConvertBase64String(_LoginStr, "ASCII")
            ''����
            SendCommand("EHLO " + _ServerName, False)
            ''��½
            SendCommand(_Base64LoginStr, False)
        End Sub

        Private _Mail As Email.SendEmail.Mail
        '<Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)> _
        ''' <summary>
        ''' �����ʼ�
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
                ''��ⷢ���ʼ��ı�Ҫ����
                If Mail.Sender.Equals(Nothing) Then
                    SetError("û�����÷�����")
                    _IsSended = False
                    Exit Sub
                End If
                If Mail.Receivers.Count = 0 Then
                    SetError("����Ҫ��һ���ռ���")
                    _IsSended = False
                    Exit Sub
                End If
                If _SmtpValidateType <> SmtpValidateTypes.None Then
                    If String.IsNullOrEmpty(_UserID) Or String.IsNullOrEmpty(_Password) Then
                        SetError("��ǰ������Ҫsmtp��֤,����û�и�����½�ʺ�")
                        _IsSended = False
                        Exit Sub
                    End If
                End If
                ''��ʼ��½
                Select Case _SmtpValidateType
                    Case SmtpValidateTypes.None
                        SendCommand("EHLO " + _ServerName, False)
                    Case SmtpValidateTypes.Login
                        LandingByLogin()
                    Case SmtpValidateTypes.Plain
                        LandingByPlain()
                    Case Else
                End Select
                ''��ʼ���ʼ��Ự(��ӦSMTP����mail) ��Ҫ���������ַ������ط��޸��޸�  �����һ�ѽ��շ����ʼ���ַ�еõ�һ����ַд��!!!!!!
                SendCommand("MAIL FROM:<" + Mail.Sender.Address + ">", False)
                ''��ʶ�ռ���(��ӦSMTP����Rcpt)
                For Each _Receive As System.Net.Mail.MailAddress In Mail.Receivers
                    SendCommand("RCPT TO:<" + _Receive.Address + ">", False)
                    If _StatusCode <> "250" Then
                        _IsSended = False
                        Exit Sub
                    End If
                Next
                ''��ʶ��ʼ�����ʼ�����(Data)
                SendCommand("DATA", False)
                ''��ʼ��д�ʼ�����
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
                If Mail.Attachments.Count = 0 Then ''û�и���
                    If Mail.MailType = MailTypes.Text Then ''�ı���ʽ
                        _Message += "Content-Type:text/plain;" + _CRLF + " ".PadRight(8, " ") + "charset=""" + Mail.MailEncoding.ToString + """" + _CRLF
                        _Message += "Content-Transfer-Encoding:base64" + _CRLF + _CRLF
                        If Not Mail.MailBody Is Nothing Then
                            _Message += Convert.ToBase64String(Mail.MailBody, 0, Mail.MailBody.Length) + _CRLF + _CRLF
                        End If
                        _Message += _CRLF + "." + _CRLF
                    Else 'Html��ʽalertnative
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
                Else '�и���
                    '����Ҫ���ʼ�����ʾ��ÿ������������
                    Dim _AttatchmentDatas As System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection()
                    For Each path As String In Mail.Attachments
                        If Not File.Exists(path) Then
                            SetError("ָ���ĸ���û���ҵ�" + path)
                        Else
                            Dim _File As FileInfo = New FileInfo(path)
                            Dim _FileStream As FileStream = New FileStream(path, FileMode.Open, FileAccess.Read)
                            If _FileStream.Length > CLng(Integer.MaxValue) Then
                                SetError("�����Ĵ�С�������������")
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
                    '�����ʼ���Ϣ
                    If Mail.MailType = MailTypes.Text Then '�ı���ʽ
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
                    Else 'Html��ʽ
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
                ''�����ʼ�����
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
                ''�����־�ļ�
                ' _Log.Debug(_Logs)
            End Try
            'Return _IsSended
        End Sub

        ''' <summary>
        ''' �첽д������
        ''' </summary>
        ''' <param name="result"></param>
        Private Sub AsyncWriterCallBack(ByVal Result As System.IAsyncResult)
            If Result.IsCompleted Then
                _IsSendComplete = True
            End If
        End Sub

        ''' <summary>
        ''' �첽��ȡ����
        ''' </summary>
        ''' <param name="result"></param>
        Private Sub AsyncReaderCallBack(ByVal Result As System.IAsyncResult)
            If Result.IsCompleted Then
                _IsSendComplete = True
            End If
        End Sub

        ''' <summary>
        ''' �ر�����
        ''' </summary> 
        Private Sub Disconnect()
            Try
                _NetworkStream.Close()
                _TcpClient.Close()
            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' ���ó��ִ���ʱ�Ķ���
        ''' </summary>
        ''' <param name="errorStr"></param>
        Private Sub SetError(ByVal ErrorStr As String)
            _ErrorMessage = ErrorStr
            '_Log.Error(ErrorStr)
            _Logs += "����:" + _CRLF + _ErrorMessage + _CRLF + "���ʼ���������ֹ��" + _CRLF
        End Sub

        ''' <summary>
        ''' ���ó��ֳɹ��Ķ���
        ''' </summary>
        ''' <param name="SuccessStr"></param>
        Private Sub SetSuccess(ByVal SuccessStr As String)
            _SuccessMessage = SuccessStr
            '_Log.Info(SuccessStr)
            _Logs += "�ɹ�:" + _CRLF + _SuccessMessage + _CRLF + "���ʼ��������ɹ���" + _CRLF
        End Sub

        ''' <summary>
        '''���ַ���ת��Ϊbase64
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

