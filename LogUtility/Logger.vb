Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.IO

Namespace LogUtility
    Public Class Logger
        Private _LogLevel As Integer = LogUtility.LogLevel.ALL
        Private _CurrentLevel As Integer
        Private _MessagePattern As String = "[yyyy-MM-dd HH:mm:ss]"
        Private _LogFile As String
        Private _LogPath As String
        Private _Catalog As String

        Public Sub New()
            Me._LogLevel = LogUtility.LogLevel.GetLevelByName(System.Configuration.ConfigurationSettings.AppSettings("LogLevel"))
            Dim LogFilePath = System.Configuration.ConfigurationSettings.AppSettings("LogFile")
            If String.IsNullOrEmpty(LogFilePath) Then
                LogFilePath = "c:\Log\$Catalog\Log_{yyyyMMdd}.log"
            End If
            Me._LogFile = LogFilePath.Replace("$Catalog", "Admin")
            Me._Catalog = "Admin"
            Dim PathArray() As String
            PathArray = Split(Me._LogFile, "\")
            Me._LogPath = LogFile.Replace(PathArray(UBound(PathArray)), "")
        End Sub

        Public Sub New(ByVal Catalog As String)
            Me._LogLevel = LogUtility.LogLevel.GetLevelByName(System.Configuration.ConfigurationSettings.AppSettings("LogLevel"))
            Dim LogFilePath = System.Configuration.ConfigurationSettings.AppSettings("LogFile")
            If String.IsNullOrEmpty(LogFilePath) Then
                LogFilePath = "c:\Log\$Catalog\Log_{yyyyMMdd}.log" 
            End If
            Me._LogFile = LogFilePath.Replace("$Catalog", Catalog)
            Me._Catalog = Catalog
            Dim PathArray() As String
            PathArray = Split(Me._LogFile, "\")
            Me._LogPath = LogFile.Replace(PathArray(UBound(PathArray)), "")
        End Sub

        Public Sub New(ByVal LoggerLevel As Integer, ByVal LoggerFile As String)
            Me._LogLevel = LoggerLevel
            Me._LogFile = LoggerFile
            Dim PathArray() As String
            PathArray = Split(Me._LogFile, "\")
            Me._LogPath = LogFile.Replace(PathArray(UBound(PathArray)), "")
        End Sub

        Public Property LogLevelName() As String
            Get
                Return LogUtility.LogLevel.GetLevelByInt(_LogLevel)
            End Get
            Set(ByVal value As String)
                Me._LogLevel = LogUtility.LogLevel.GetLevelByName(value)
            End Set
        End Property

        Public Property LogLevel() As Integer
            Get
                Return _LogLevel
            End Get
            Set(ByVal value As Integer)
                _LogLevel = value
            End Set
        End Property

        Public Property MessagePattern() As String
            Get
                Return _MessagePattern
            End Get
            Set(ByVal value As String)
                _MessagePattern = value
            End Set
        End Property

        Public Property LogPath() As String
            Get
                Return Me._LogPath
            End Get
            Set(ByVal value As String)
                If value.IndexOf("\", Len(value) - 1) < 1 Then
                    value = value + "\"
                End If
                Me._LogFile = value + _Catalog + "\{yyyyMMdd}.log"
                Dim PathArray() As String
                PathArray = Split(Me._LogFile, "\")
                Dim a As Integer = UBound(PathArray)
                Me._LogPath = LogFile.Replace(PathArray(UBound(PathArray)), "")
            End Set
        End Property

        Public Property LogFile() As String
            Get
                Return _LogFile
            End Get
            Set(ByVal value As String)
                _LogFile = FormatLogFile(value)
            End Set
        End Property

        Public Sub Fatal(ByVal Message As Object)
            If _LogLevel >= LogUtility.LogLevel.FATAL Then
                _CurrentLevel = LogUtility.LogLevel.FATAL
                If IsNothing(Message) Then
                    Log("NULL")
                Else
                    Log(Message.ToString)
                End If
            End If
        End Sub

        Public Sub [Error](ByVal Message As Object)
            If _LogLevel >= LogUtility.LogLevel.[ERROR] Then
                _CurrentLevel = LogUtility.LogLevel.[ERROR]
                If IsNothing(Message) Then
                    Log("NULL")
                Else
                    Log(Message.ToString)
                End If
            End If
        End Sub

        Public Sub Warn(ByVal Message As Object)
            If _LogLevel >= LogUtility.LogLevel.WARN Then
                _CurrentLevel = LogUtility.LogLevel.WARN
                If IsNothing(Message) Then
                    Log("NULL")
                Else
                    Log(Message.ToString)
                End If
            End If
        End Sub

        Public Sub Info(ByVal Message As Object)
            If _LogLevel >= LogUtility.LogLevel.INFO Then
                _CurrentLevel = LogUtility.LogLevel.INFO
                If IsNothing(Message) Then
                    Log("NULL")
                Else
                    Log(Message.ToString)
                End If
            End If
        End Sub

        Public Sub Debug(ByVal Message As Object)
            If _LogLevel >= LogUtility.LogLevel.DEBUG Then
                _CurrentLevel = LogUtility.LogLevel.DEBUG
                If IsNothing(Message) Then
                    Log("NULL")
                Else
                    Log(Message.ToString)
                End If
            End If
        End Sub

        ''' <summary>
        ''' 监测文件存在否
        ''' </summary>
        ''' <param name="FilePath">文件路径</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ExistsFile(ByVal FilePath As String) As Boolean
            If IO.File.Exists(FilePath) = False Then
                'IO.Directory.CreateDirectory(fPathTemp)
                Dim _PathArray() As String
                Dim _PathTemp As String
                _PathArray = Split(FilePath, "\")
                _PathTemp = FilePath.Replace(_PathArray(UBound(_PathArray)), "")
                ExistsPath(_PathTemp)
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' 监测是否存在路径
        ''' </summary>
        ''' <param name="Path">路径</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ExistsPath(ByVal Path As String) As Boolean
            'Dim PathArray() As String
            'Dim fPathTemp As String
            'PathArray = Split(fPathData, "\")
            'fPathTemp = fPathData.Replace(PathArray(UBound(PathArray)), "")
            If IO.Directory.Exists(Path) = False Then
                'IO.Directory.CreateDirectory(Path)
                Return False
            Else
                Return True
            End If
        End Function


        ''' <summary>
        ''' 创建路径
        ''' </summary>
        ''' <param name="Path">路径</param>
        ''' <remarks></remarks>
        Private Sub CreaterPath(ByVal Path As String)
            'Dim PathArray() As String
            'Dim fPathTemp As String
            'PathArray = Split(fPathData, "\")
            'fPathTemp = fPathData.Replace(PathArray(UBound(PathArray)), "")
            If IO.Directory.Exists(Path) = False Then
                IO.Directory.CreateDirectory(Path)
            End If
        End Sub

        <Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)> _
        Public Sub Log(ByVal Message As String)
            Dim _FileStream As FileStream = Nothing
            If Not ExistsPath(Me._LogPath) Then CreaterPath(Me._LogPath)
            Try
                _FileStream = New FileStream(FormatLogFile(_LogFile), FileMode.OpenOrCreate)
                _FileStream.Position = _FileStream.Length
                Dim _StreamWriter As StreamWriter = New StreamWriter(_FileStream)
                _StreamWriter.WriteLine(FormatMessage(Message))
                _StreamWriter.Close()
            Catch ex As Exception

            Finally
                If Not IsNothing(_FileStream) Then
                    _FileStream.Close()
                End If
            End Try
        End Sub

        Private Function FormatMessage(ByVal Message As String)
            Message = System.DateTime.Now.ToString(_MessagePattern) + "[" + LogUtility.LogLevel.LEVEL_STRING([_CurrentLevel]) + "]" + Message
            Return Message
        End Function

        Private Function FormatLogFile(ByVal LogFile As String)
            Dim _DateFormatStartPos As Integer = LogFile.IndexOf("{")
            Dim _DateFormatEndPos As Integer = LogFile.IndexOf("}")
            Dim _DateFormat As String = System.DateTime.Now.ToString(LogFile.Substring(_DateFormatStartPos + 1, _DateFormatEndPos - _DateFormatStartPos - 1))
            Dim _FormattedLogFile As String = LogFile.Substring(0, _DateFormatStartPos) + _DateFormat + LogFile.Substring(_DateFormatEndPos + 1)
            Return _FormattedLogFile
        End Function

    End Class
End Namespace


