Namespace Utility

    ''' <summary>
    ''' 上传工具类
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Upload


        ''' <summary>
        ''' 删除文件
        ''' </summary>
        ''' <param name="FilePath">文件路径</param>
        ''' <remarks></remarks>
        Public Shared Sub DeleteFile(ByVal FilePath As String)
            If ExistsFile(FilePath) Then
                IO.File.Delete(FilePath)
            End If
        End Sub

        ''' <summary>
        ''' 删除路径
        ''' </summary>
        ''' <param name="FilePath">文件路径</param>
        ''' <remarks></remarks>
        Public Shared Sub DeletePath(ByVal FilePath As String)
            If ExistsPath(FilePath) Then
                Dim _PathArray() As String
                Dim _PathTemp As String
                _PathArray = Split(FilePath, "\")
                _PathTemp = FilePath 'PictruePath.Replace(PathArray(UBound(PathArray)), "")
                IO.Directory.Delete(_PathTemp, True)
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
        Public Shared Sub CreaterPath(ByVal Path As String)
            'Dim PathArray() As String
            'Dim fPathTemp As String
            'PathArray = Split(fPathData, "\")
            'fPathTemp = fPathData.Replace(PathArray(UBound(PathArray)), "")
            If IO.Directory.Exists(Path) = False Then
                IO.Directory.CreateDirectory(Path)
            End If
        End Sub


        '设置文件存贮目录

        Private _File_Path As String
        ''' <summary>
        ''' 文件存贮目录
        ''' </summary>
        ''' <value></value>
        ''' <remarks></remarks>
        Public WriteOnly Property FilePath() As String
            Set(ByVal Value As String)
                _File_Path = Value
            End Set
        End Property

        '返回文件名
        Private _File_Name As String
        ''' <summary>
        ''' 文件名
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FileName() As String
            Get
                Return _File_Name
            End Get
            Set(ByVal Value As String)
                _File_Name = Value
            End Set
        End Property


        '返回文件后缀名
        Private _File_Post_Fix_Name As String
        ''' <summary>
        ''' 文件后缀名
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FilePostFixName() As String
            Get
                Return _File_Post_Fix_Name
            End Get
        End Property

        '返回文件类型
        Private _File_Type As String
        ''' <summary>
        ''' 文件类型
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FileType() As String
            Get
                Return _File_Type
            End Get
        End Property

        '返回文件尺寸
        Private _File_Size As Integer
        ''' <summary>
        ''' 文件尺寸
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FileSize() As Integer
            Get
                Return _File_Size
            End Get
        End Property

        '返回信息
        Private _File_Upload_Msg As String
        ''' <summary>
        ''' 返回信息
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FileUploadMsg() As String
            Get
                Return _File_Upload_Msg
            End Get
        End Property

        '文件的内容
        Private FileContentData As Byte()

        '文件
        Private _File_Data As System.Web.UI.HtmlControls.HtmlInputFile
        Public WriteOnly Property FileData() As System.Web.UI.HtmlControls.HtmlInputFile
            Set(ByVal value As System.Web.UI.HtmlControls.HtmlInputFile)
                _File_Type = value.PostedFile.ContentType
                Dim _FileSplit() As String = Split(value.PostedFile.FileName, ".")
                _File_Post_Fix_Name = "." + _FileSplit(_FileSplit.Length - 1)
                _FileSplit = Split(value.PostedFile.FileName, "\")
                _File_Name = _FileSplit(_FileSplit.Length - 1)
                _File_Size = value.PostedFile.ContentLength
                _File_Data = value
            End Set
        End Property

        'Enum storeType
        '    DataBase = 0
        '    File = 1
        'End Enum


        Public Sub New()

        End Sub

        Public Sub New(ByVal UploadFile As Web.UI.HtmlControls.HtmlInputFile)
            _File_Type = UploadFile.PostedFile.ContentType
            Dim _FileSplit() As String = Split(UploadFile.PostedFile.FileName, ".")
            _File_Post_Fix_Name = "." + _FileSplit(_FileSplit.Length - 1)
            _FileSplit = Split(UploadFile.PostedFile.FileName, "\")
            _File_Name = _FileSplit(_FileSplit.Length - 1)
            _File_Size = UploadFile.PostedFile.ContentLength
            _File_Data = UploadFile
        End Sub



        '文件保存
        Public Overloads Sub SaveFile()
            With _File_Data
                Try
                    '得到文件的具体信息
                    'fileNamedata = Format(Now(), "yyyyMMddHmmss") + CInt(Int((9 * Rnd()) + 0)).ToString + "." + filePostfixNamedata
                    ReDim FileContentData(.PostedFile.InputStream.Length)
                    .PostedFile.InputStream.Read(FileContentData, 0, .PostedFile.InputStream.Length)
                    '文件保存
                    'Select Case storeType
                    '   Case storeType.DataBase
                    'fileDirdata = ""
                    'Me.saveFileInformation(pid, uid, storeType)
                    'fileUploadMsgdata = "保存成功"
                    '   Case storeType.File

                    If Not IsNothing(_File_Path) Then
                        If Not ExistsPath(_File_Path) Then CreaterPath(_File_Path)

                        .PostedFile.SaveAs(_File_Path + _File_Name)
                        ' Me.saveFileInformation(pid, uid, storeType)
                        _File_Upload_Msg = "保存成功"

                    Else
                        _File_Upload_Msg = "存储为文件时你没有设置目录"
                    End If
                    'End Select
                Catch e As Exception
                    _File_Upload_Msg = e.ToString
                End Try
            End With
        End Sub


    End Class
End Namespace