
Namespace DatabaseOperation



    Public Class DataBaseFactory



        ''' <summary>
        ''' 获取一个数据库操作对象
        ''' </summary>
        ''' <param name="DataBaseEnum">操作数据库类型</param>
        ''' <returns>操作数据库对象</returns>
        ''' <remarks></remarks>
        Public Overloads Function GetDataBase(ByVal DataBaseEnum As DataBaseEnum) As DataBaseClass
            Dim _DataBase As DataBaseClass = Nothing
            Select Case DataBaseEnum
                Case DataBaseEnum.SQLServer 'SQL Server
                    _DataBase = New DataBase_SQLServer()
                    'Case DatabaseOperation.DataBaseEnum.MySQLServer
                    '    _DataBase = New DataBase_MySQLServer()
            End Select
            Return _DataBase
        End Function


        '''' <summary>
        '''' 获取一个数据库操作对象
        '''' </summary>
        '''' <param name="ConnectionString">数据库连接字符串</param>
        '''' <returns>操作数据库对象</returns>
        '''' <remarks></remarks>
        'Public Overloads Function GetDataBase(ByVal DataBaseEnum As DataBaseEnum, ByVal ConnectionString As String) As DataBaseClass
        '    If ConnectionString.IndexOf("provider=") < 0 Then
        '        ' 如果连接字符串中没有provider的标志，认为是SQL Server数据库
        '        Return New DataBase_SQLServer(ConnectionString) ' SqlDBOperator(strConnection);
        '    Else
        '        ' 如果连接字符串中有provider的标志，认为是OLE DB数据库
        '        Return Nothing
        '    End If

        'End Function


        ''' <summary>
        ''' 获取一个数据库操作对象
        ''' </summary>
        ''' <param name="ConnectionString">数据库连接字符串</param>
        ''' <param name="DataBaseEnum">操作数据库类型</param>
        ''' <returns>操作数据库对象</returns>
        ''' <remarks></remarks>
        Public Overloads Function GetDataBase(ByVal ConnectionString As String, ByVal DataBaseEnum As DataBaseEnum) As DataBaseClass
            Dim _DataBase As DataBaseClass = Nothing
            Select Case DataBaseEnum
                Case DataBaseEnum.SQLServer 'SQL Server
                    _DataBase = New DataBase_SQLServer(ConnectionString)
                    'Case DatabaseOperation.DataBaseEnum.MySQLServer
                    '    _DataBase = New DataBase_MySQLServer(ConnectionString)
            End Select
            Return _DataBase
        End Function
    End Class


End Namespace