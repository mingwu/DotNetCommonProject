

Imports LogUtility
Namespace DatabaseOperation


    ''' <summary>
    ''' 数据库基类
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
        ''' 连接数据库字符串
        ''' </summary>
        Public MustOverride Property Connection() As String


        ''' <summary>
        ''' 打开数据库连接
        ''' </summary>
        Public MustOverride Sub Open()

        ''' <summary>
        ''' 打开数据库连接
        ''' </summary>
        ''' <param name="ConnectionString">数据库连接字符串</param>
        ''' <remarks></remarks>
        Public MustOverride Sub Open(ByVal ConnectionString As String)


        ''' <summary>
        ''' 关闭数据库连接
        ''' </summary>
        Public MustOverride Sub Close()







        ''' <summary>
        ''' 检验是否存在数据
        ''' </summary>
        ''' <param name="SqlString">Sql语句</param>
        ''' <returns>返回 TRUE FALSE</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExistDate(ByVal SqlString As String) As Boolean


        ''' <summary>
        ''' 执行SQL语句
        ''' </summary>
        ''' <param name="SqlString">SQL语句</param>
        ''' <param name="Parameters">参数集合</param>
        ''' <returns>是否存在数据</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExistDate(ByVal SqlString As String, ByVal Parameters As SqlClient.SqlParameter()) As Boolean


        ''' <summary>
        ''' 执行SQL语句
        ''' </summary>
        ''' <param name="SqlString">SQL语句</param>
        ''' <param name="Parameters">参数集合</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSqlNonQuery(ByVal SqlString As String, ByVal Parameters As SqlClient.SqlParameter())

        ''' <summary>
        ''' 执行SQL语句
        ''' </summary>
        ''' <param name="SqlString">SQL语句</param>
        ''' <returns>不返回值</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSql(ByVal SqlString As String) As DataTable


        ''' <summary>
        ''' 执行SQL语句
        ''' </summary>
        ''' <param name="SqlString">SQL语句</param>
        ''' <param name="DataTableObj">DataTable对象</param>
        ''' <returns>返回DataSet对象</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As DataTable) As DataTable


        ''' <summary>
        ''' 执行SQL语句
        ''' </summary>
        ''' <param name="SqlString">SQL语句</param>
        ''' <param name="DataTableObj">DataTable对象</param>
        ''' <param name="Parameters">参数集合</param>
        ''' <returns>返回DataTable对象</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As DataTable, ByVal Parameters As SqlClient.SqlParameter()) As DataTable

        ''' <summary>
        ''' 执行SQL语句
        ''' </summary>
        ''' <param name="SqlString">SQL语句</param>
        ''' <param name="DataTableObj">DataTable对象</param>
        ''' <param name="PageSize">每页条数</param>
        ''' <param name="PageNo">页码</param>
        ''' <returns>返回DataTable对象</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As System.Data.DataTable, ByVal PageSize As Integer, ByVal PageNo As Integer) As System.Data.DataTable

        ''' <summary>
        ''' 执行SQL语句
        ''' </summary>
        ''' <param name="SqlString">SQL语句</param>
        ''' <param name="DataTableObj">DataTable对象</param>
        ''' <param name="PageSize">每页条数</param>
        ''' <param name="PageNo">页码</param>
        ''' <param name="Parameters">参数集合</param>
        ''' <returns>返回DataTable对象</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As System.Data.DataTable, ByVal PageSize As Integer, ByVal PageNo As Integer, ByVal Parameters As SqlClient.SqlParameter()) As System.Data.DataTable

        ''' <summary>
        ''' 执行SQL语句
        ''' </summary>
        ''' <param name="SqlString">SQL语句</param>
        ''' <returns>返回SqlDataReader对象</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSqlReturnDataReader(ByVal SqlString As String) As System.Data.SqlClient.SqlDataReader

        ''' <summary>
        ''' 执行SQL语句
        ''' </summary>
        ''' <param name="SqlString">SQL语句</param>
        ''' <param name="Parameters">参数集合</param>
        ''' <returns>返回SqlDataReader对象</returns>
        ''' <remarks></remarks>
        Public MustOverride Function ExecuteSqlReturnDataReader(ByVal SqlString As String, ByVal Parameters As SqlClient.SqlParameter()) As System.Data.SqlClient.SqlDataReader

        '''' <summary>
        '''' 执行SQL语句
        '''' </summary>
        '''' <param name="SqlString">SQL语句</param>
        '''' <param name="PageSize">每页条数</param>
        '''' <param name="PageNo">页码</param>
        '''' <returns>返回SqlDataReader对象</returns>
        '''' <remarks></remarks>
        'Public MustOverride Function ExecuteSqlReturnDataReader(ByVal SqlString As String, ByVal PageSize As Integer, ByVal PageNo As Integer) As System.Data.SqlClient.SqlDataReader

        '''' <summary>
        '''' 执行SQL语句
        '''' </summary>
        '''' <param name="SqlString">SQL语句</param>
        '''' <param name="PageSize">每页条数</param>
        '''' <param name="PageNo">页码</param>
        '''' <param name="Parameters">参数集合</param>
        '''' <returns>返回SqlDataReader对象</returns>
        '''' <remarks></remarks>
        'Public MustOverride Function ExecuteSqlReturnDataReader(ByVal SqlString As String, ByVal PageSize As Integer, ByVal PageNo As Integer, ByVal Parameters As SqlClient.SqlParameter()) As System.Data.SqlClient.SqlDataReader


        'Public MustOverride Sub AddTable(ByVal SqlString As String)

    End Class


End Namespace