
Namespace DatabaseOperation



    Public Class DataBaseFactory



        ''' <summary>
        ''' ��ȡһ�����ݿ��������
        ''' </summary>
        ''' <param name="DataBaseEnum">�������ݿ�����</param>
        ''' <returns>�������ݿ����</returns>
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
        '''' ��ȡһ�����ݿ��������
        '''' </summary>
        '''' <param name="ConnectionString">���ݿ������ַ���</param>
        '''' <returns>�������ݿ����</returns>
        '''' <remarks></remarks>
        'Public Overloads Function GetDataBase(ByVal DataBaseEnum As DataBaseEnum, ByVal ConnectionString As String) As DataBaseClass
        '    If ConnectionString.IndexOf("provider=") < 0 Then
        '        ' ��������ַ�����û��provider�ı�־����Ϊ��SQL Server���ݿ�
        '        Return New DataBase_SQLServer(ConnectionString) ' SqlDBOperator(strConnection);
        '    Else
        '        ' ��������ַ�������provider�ı�־����Ϊ��OLE DB���ݿ�
        '        Return Nothing
        '    End If

        'End Function


        ''' <summary>
        ''' ��ȡһ�����ݿ��������
        ''' </summary>
        ''' <param name="ConnectionString">���ݿ������ַ���</param>
        ''' <param name="DataBaseEnum">�������ݿ�����</param>
        ''' <returns>�������ݿ����</returns>
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