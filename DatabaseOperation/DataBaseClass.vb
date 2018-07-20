

Imports LogUtility
Namespace DatabaseOperation


    ''' <summary>
    ''' ���ݿ����
    ''' </summary>
    Public MustInherit Class DataBaseClass

        Protected _Log As LogUtility.Logger
        Public Property Log() As LogUtility.Logger
            Get
                Return _Log
            End Get
            Set(ByVal value As LogUtility.Logger)
                _Log = value
            End Set
        End Property


        ''' <summary>
        ''' �������ݿ��ַ���
        ''' </summary>
        Public MustOverride Property Connection() As String


        ''' <summary>
        ''' �����ݿ�����
        ''' </summary>
        Public MustOverride Sub Open()

        ''' <summary>
        ''' �����ݿ�����
        ''' </summary>
        ''' <param name="ConnectionString">���ݿ������ַ���</param>
        ''' <remarks></remarks>
        Public MustOverride Sub Open(ByVal ConnectionString As String)


        ''' <summary>
        ''' �ر����ݿ�����
        ''' </summary>
        Public MustOverride Sub Close()







        ''' <summary>
        ''' �����Ƿ��������
        ''' </summary>
        ''' <param name="SqlString">Sql���</param>
        ''' <returns>���� TRUE FALSE</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExistDate(ByVal SqlString As String) As Boolean


        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="Parameters">��������</param>
        ''' <returns>�Ƿ��������</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExistDate(ByVal SqlString As String, ByVal Parameters As SqlClient.SqlParameter()) As Boolean


        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="Parameters">��������</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSqlNonQuery(ByVal SqlString As String, ByVal Parameters As SqlClient.SqlParameter())

        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <returns>������ֵ</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSql(ByVal SqlString As String) As DataTable


        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="DataTableObj">DataTable����</param>
        ''' <returns>����DataSet����</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As DataTable) As DataTable


        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="DataTableObj">DataTable����</param>
        ''' <param name="Parameters">��������</param>
        ''' <returns>����DataTable����</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As DataTable, ByVal Parameters As SqlClient.SqlParameter()) As DataTable

        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="DataTableObj">DataTable����</param>
        ''' <param name="PageSize">ÿҳ����</param>
        ''' <param name="PageNo">ҳ��</param>
        ''' <returns>����DataTable����</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As System.Data.DataTable, ByVal PageSize As Integer, ByVal PageNo As Integer) As System.Data.DataTable

        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="DataTableObj">DataTable����</param>
        ''' <param name="PageSize">ÿҳ����</param>
        ''' <param name="PageNo">ҳ��</param>
        ''' <param name="Parameters">��������</param>
        ''' <returns>����DataTable����</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As System.Data.DataTable, ByVal PageSize As Integer, ByVal PageNo As Integer, ByVal Parameters As SqlClient.SqlParameter()) As System.Data.DataTable

        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <returns>����SqlDataReader����</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSqlReturnDataReader(ByVal SqlString As String) As System.Data.SqlClient.SqlDataReader

        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="Parameters">��������</param>
        ''' <returns>����SqlDataReader����</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSqlReturnDataReader(ByVal SqlString As String, ByVal Parameters As SqlClient.SqlParameter()) As System.Data.SqlClient.SqlDataReader

        '''' <summary>
        '''' ִ��SQL���
        '''' </summary>
        '''' <param name="SqlString">SQL���</param>
        '''' <param name="PageSize">ÿҳ����</param>
        '''' <param name="PageNo">ҳ��</param>
        '''' <returns>����SqlDataReader����</returns>
        '''' <remarks></remarks>
        'Public MustOverride Function ExecuteSqlReturnDataReader(ByVal SqlString As String, ByVal PageSize As Integer, ByVal PageNo As Integer) As System.Data.SqlClient.SqlDataReader

        '''' <summary>
        '''' ִ��SQL���
        '''' </summary>
        '''' <param name="SqlString">SQL���</param>
        '''' <param name="PageSize">ÿҳ����</param>
        '''' <param name="PageNo">ҳ��</param>
        '''' <param name="Parameters">��������</param>
        '''' <returns>����SqlDataReader����</returns>
        '''' <remarks></remarks>
        'Public MustOverride Function ExecuteSqlReturnDataReader(ByVal SqlString As String, ByVal PageSize As Integer, ByVal PageNo As Integer, ByVal Parameters As SqlClient.SqlParameter()) As System.Data.SqlClient.SqlDataReader


        'Public MustOverride Sub AddTable(ByVal SqlString As String)

    End Class


End Namespace