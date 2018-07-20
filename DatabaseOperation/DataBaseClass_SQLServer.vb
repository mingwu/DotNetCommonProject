
Namespace DatabaseOperation


    'ͨ�����ݿ���
    ''' <summary>
    ''' SQL Server ���ݿ������
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DataBase_SQLServer
        Inherits DataBaseClass



        ''' <summary>
        ''' �����ַ�����˽��
        ''' </summary>
        ''' <remarks></remarks>
        Private _ConnectionString As String

        ''' <summary>
        ''' ���ݿ����Ӷ���˽��
        ''' </summary>
        ''' <remarks></remarks>
        Private _ConnectionObject As SqlClient.SqlConnection = Nothing


        ''' <summary>
        ''' �����ַ���������
        ''' </summary>
        ''' <value>�����ַ���</value>
        ''' <returns>�����ַ���</returns>
        ''' <remarks></remarks>
        Public Overrides Property Connection() As String
            Get
                Return _ConnectionString
            End Get
            Set(ByVal value As String)
                _ConnectionString = value
            End Set
        End Property


        ''' <summary>
        ''' SQL Server ���ݿ�
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            _ConnectionString = System.Configuration.ConfigurationSettings.AppSettings("DBString")
            _Log = New LogUtility.Logger("DatabaseLog")
        End Sub

        ''' <summary>
        ''' SQL Server ���ݿ�
        ''' </summary>
        ''' <param name="ConnectionString">���ݿ������ַ���</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal ConnectionString As String)
            _ConnectionString = ConnectionString
            _Log = New LogUtility.Logger("DatabaseLog")
        End Sub


        ''' <summary>
        ''' �����ݿ�����
        ''' </summary>
        Public Overloads Overrides Sub Open()
            _ConnectionObject = New SqlClient.SqlConnection(_ConnectionString)
            If _ConnectionObject.State <> ConnectionState.Open Then
                Try
                    _ConnectionObject.Open()
                Catch ex As Exception
                    'Throw New Exception("Error :" + ex.Message)
                    _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute Open() Error:" + ex.Message + ex.StackTrace)
                End Try
            End If
        End Sub


        ''' <summary>
        ''' �����ݿ�����
        ''' </summary>
        ''' <param name="ConnectionString">���ݿ������ַ���</param>
        ''' <remarks></remarks>
        Public Overloads Overrides Sub Open(ByVal ConnectionString As String)
            _ConnectionObject = New SqlClient.SqlConnection(ConnectionString)
            If _ConnectionObject.State <> ConnectionState.Open Then
                Try
                    _ConnectionObject.Open()
                Catch ex As Exception
                    'Throw New Exception("Error :" + ex.Message)
                    _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute Open(ConnectionString) Error:" + ex.Message + ex.StackTrace)
                End Try
            End If
        End Sub


        ''' <summary>
        ''' �ر����ݿ�����
        ''' </summary>
        Public Overrides Sub Close()
            If _ConnectionObject.State = ConnectionState.Open Then
                Try
                    _ConnectionObject.Close()
                    _ConnectionObject.Dispose()
                Catch ex As Exception
                    'Throw New Exception("Error :" + ex.Message)
                    _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute Close() Error:" + ex.Message + ex.StackTrace)
                End Try
            End If
            'If Convert.IsDBNull(_ConnectionObject) Then
            '    _ConnectionObject.Close()
            '    _ConnectionObject.Dispose()
            'End If
            GC.Collect()
        End Sub

        ''' <summary>
        ''' ����Command����˽��
        ''' </summary>
        ''' <param name="SqlString">Sql���</param>
        ''' <param name="ConnectionObject">Connection����</param>
        ''' <returns>Command����</returns>
        ''' <remarks></remarks>
        Protected Function CreateCmd(ByVal SqlString As String, ByVal ConnectionObject As SqlClient.SqlConnection) As SqlClient.SqlCommand
            Try
                Dim _Command As New SqlClient.SqlCommand(SqlString, ConnectionObject)
                Return _Command
            Catch ex As Exception
                _Log.Fatal("[DatabaseOperation.DataBaseClass]Execute CreateCmd(SqlString,ConnectionObject) Error:" + ex.Message + ex.StackTrace)
            End Try

        End Function

        ''' <summary>
        ''' ����Command����˽��
        ''' </summary>
        ''' <param name="SqlString">Sql���</param>
        ''' <param name="ConnectionObject">Connection����</param>
        ''' <param name="Parameters">��������</param>
        ''' <param name="CommondType">CommondText����</param>
        ''' <returns>Command����</returns>
        ''' <remarks></remarks>
        Protected Function CreateCmd(ByVal SqlString As String, ByVal ConnectionObject As SqlClient.SqlConnection, ByVal Parameters() As SqlClient.SqlParameter, ByVal CommondType As Data.CommandType) As SqlClient.SqlCommand
            Try
                Dim _Cmd As New SqlClient.SqlCommand(SqlString, ConnectionObject)
                _Cmd.CommandType = CommondType
                _Cmd.Parameters.Clear()
                If (Not IsDBNull(Parameters)) And (Not IsNothing(Parameters)) Then
                    Dim _Parameter As SqlClient.SqlParameter
                    For Each _Parameter In Parameters
                        If Not IsDBNull(_Parameter) And (Not IsNothing(_Parameter)) Then
                            _Cmd.Parameters.AddWithValue(_Parameter.ParameterName, _Parameter.Value)
                        End If
                    Next
                End If
                Return _Cmd
            Catch ex As Exception
                _Log.Fatal("[DatabaseOperation.DataBaseClass]Execute CreateCmd(SqlString,ConnectionObject,Parameters(),CommondType) Error:" + ex.Message + ex.StackTrace)
            End Try

        End Function

        ''' <summary>
        ''' �����Ƿ��������
        ''' </summary>
        ''' <param name="SqlString">Sql���</param>
        ''' <returns>���� TRUE FALSE</returns>
        ''' <remarks></remarks>
        Public Overrides Function ExistDate(ByVal SqlString As String) As Boolean
            Me.Open(_ConnectionString)
            Dim dr As SqlClient.SqlDataReader
            dr = CreateCmd(SqlString, _ConnectionObject).ExecuteReader
            If dr.Read() Then
                Me.Close()
                Return True
            Else
                Me.Close()
                Return False
            End If
        End Function


        ''' <summary>
        ''' �����Ƿ��������
        ''' </summary>
        ''' <param name="SqlString">Sql���</param>
        ''' <param name="Parameters">��������</param>
        ''' <returns>���� TRUE FALSE</returns>
        ''' <remarks></remarks>
        Public Overrides Function ExistDate(ByVal SqlString As String, ByVal Parameters As SqlClient.SqlParameter()) As Boolean
            Me.Open(_ConnectionString)
            Dim dr As SqlClient.SqlDataReader
            dr = CreateCmd(SqlString, _ConnectionObject, Parameters, CommandType.Text).ExecuteReader
            If dr.Read() Then
                Me.Close()
                Return True
            Else
                Me.Close()
                Return False
            End If
        End Function

        ''' <summary>
        ''' ����SQL���
        ''' </summary>
        ''' <param name="SqlString">Sql���</param>
        ''' <remarks></remarks>
        Public Overrides Function ExecuteSql(ByVal SqlString As String) As System.Data.DataTable
            Me.Open(_ConnectionString)
            'Dim _CommandObj As SqlClient.SqlCommand
            Dim da As SqlClient.SqlDataAdapter
            da = New SqlClient.SqlDataAdapter(SqlString, _ConnectionObject)
            Dim DataTableObj As New System.Data.DataTable

            'Dim _SqlTransaction As SqlClient.SqlTransaction = _ConnectionObject.BeginTransaction
            '_CommandObj.Transaction = _SqlTransaction New
            Try
                '_CommandObj = CreateCmd(SqlString, _ConnectionObject)
                '_CommandObj.ExecuteNonQuery()
                da.Fill(DataTableObj)
                '_SqlTransaction.Commit()
            Catch ex As Exception
                '_SqlTransaction.Rollback()
                'Throw New Exception("Error :" + ex.Message)
                _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute ExecuteSql(SqlString) Sql:" + SqlString + " Error:" + ex.Message + ex.StackTrace)
            End Try
            Me.Close()
            'Return Nothing
            Return DataTableObj
        End Function

        ''' <summary>
        ''' ����SQL���
        ''' </summary>
        ''' <param name="SqlString">Sql���</param>
        ''' <remarks></remarks>
        Public Overrides Function ExecuteSqlNonQuery(ByVal SqlString As String, ByVal Parameters() As SqlClient.SqlParameter)
            Me.Open(_ConnectionString)
            Dim _CommandObj As SqlClient.SqlCommand
            'Dim _SqlTransaction As SqlClient.SqlTransaction = _ConnectionObject.BeginTransaction
            '_CommandObj.Transaction = _SqlTransaction New
            Try
                _CommandObj = CreateCmd(SqlString, _ConnectionObject, Parameters, CommandType.Text)
                _CommandObj.ExecuteNonQuery()
                '_SqlTransaction.Commit()
            Catch ex As Exception
                '_SqlTransaction.Rollback()
                'Throw New Exception("Error :" + ex.Message)
                _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute ExecuteSql(SqlString) Sql:" + SqlString + " Error:" + ex.Message + ex.StackTrace)
            End Try
            Me.Close()
            Return Nothing
        End Function


        '''' <summary>
        '''' ִ��SQL���
        '''' </summary>
        '''' <param name="SqlString">SQL���</param>
        '''' <param name="DataSetObj">DataSet����</param>
        '''' <returns>����DataSet����</returns>
        '''' <remarks></remarks>
        'Public Overloads Overrides Function ExecuteSql(ByVal SqlString As String, ByVal DataSetObj As System.Data.DataSet) As System.Data.DataSet
        '    Me.Open(_ConnectionString)
        '    Dim da As SqlClient.SqlDataAdapter
        '    da = New SqlClient.SqlDataAdapter(SqlString, _ConnectionObject)
        '    Try
        '        da.Fill(DataSetObj)
        '    Catch ex As Exception
        '        Throw New Exception("Error :" + ex.Message)
        '    End Try
        '    Me.Close()
        '    Return DataSetObj
        'End Function


        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="DataTableObj">DataTable����</param>
        ''' <returns>����DataTable����</returns>
        ''' <remarks></remarks>
        Public Overloads Overrides Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As System.Data.DataTable) As System.Data.DataTable
            Me.Open(_ConnectionString)
            Dim da As SqlClient.SqlDataAdapter
            da = New SqlClient.SqlDataAdapter(SqlString, _ConnectionObject)
            Try
                da.Fill(DataTableObj)
            Catch ex As Exception
                'Throw New Exception("Error :" + ex.Message)
                _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute ExecuteSql(SqlString,DataTableObj) Sql:" + SqlString + " Error:" + ex.Message + ex.StackTrace)
            End Try
            Me.Close()
            Return DataTableObj
        End Function

 

        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="DataTableObj">DataTable����</param>
        ''' <returns>����DataTable����</returns>
        ''' <remarks></remarks>
        Public Overloads Overrides Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As System.Data.DataTable, ByVal Parameters() As SqlClient.SqlParameter) As System.Data.DataTable
            Me.Open(_ConnectionString)
            Dim da As New SqlClient.SqlDataAdapter
            'da = New SqlClient.SqlDataAdapter(SqlString, _ConnectionObject)
            Dim _CommandObj As SqlClient.SqlCommand
            Try
                _CommandObj = CreateCmd(SqlString, _ConnectionObject, Parameters, CommandType.Text)
                da.SelectCommand = _CommandObj
                da.Fill(DataTableObj)
            Catch ex As Exception
                'Throw New Exception("Error :" + ex.Message)
                _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute ExecuteSql(SqlString,DataTableObj,Parameters()) Sql:" + SqlString + " Error:" + ex.Message + ex.StackTrace)
            End Try
            Me.Close()
            Return DataTableObj
        End Function

        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="DataTableObj">DataTable����</param>
        ''' <param name="PageSize">ÿҳ����</param>
        ''' <param name="PageNo">ҳ��</param>
        ''' <returns>����DataTable����</returns>
        ''' <remarks></remarks>
        Public Overloads Overrides Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As System.Data.DataTable, ByVal PageSize As Integer, ByVal PageNo As Integer) As System.Data.DataTable
            Me.Open(_ConnectionString)
            Dim da As SqlClient.SqlDataAdapter
            da = New SqlClient.SqlDataAdapter(SqlString, _ConnectionObject)
            Dim StartID As Integer
            StartID = PageNo * PageSize - (PageSize - 1)
            If StartID - 1 < 1 Then StartID = 1
            Try
                da.Fill(StartID - 1, PageSize, DataTableObj)
            Catch ex As Exception
                ' Throw New Exception("Error :" + ex.Message)
                _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute ExecuteSql(SqlString,DataTableObj,PageSize,PageNo) Sql:" + SqlString + " Error:" + ex.Message + ex.StackTrace)
            End Try
            Me.Close()
            Return DataTableObj
        End Function

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
        Public Overloads Overrides Function ExecuteSql(ByVal SqlString As String, ByVal DataTableObj As System.Data.DataTable, ByVal PageSize As Integer, ByVal PageNo As Integer, ByVal Parameters() As SqlClient.SqlParameter) As System.Data.DataTable
            Me.Open(_ConnectionString)
            Dim da As New SqlClient.SqlDataAdapter
            'da = New SqlClient.SqlDataAdapter(SqlString, _ConnectionObject)
            Dim _CommandObj As SqlClient.SqlCommand
            Dim StartID As Integer
            StartID = PageNo * PageSize - (PageSize - 1)
            If StartID - 1 < 1 Then StartID = 1
            Try
                _CommandObj = CreateCmd(SqlString, _ConnectionObject, Parameters, CommandType.Text)
                da.SelectCommand = _CommandObj
                da.Fill(StartID - 1, PageSize, DataTableObj)
            Catch ex As Exception
                ' Throw New Exception("Error :" + ex.Message)
                _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute ExecuteSql(SqlString,DataTableObj,PageSize,PageNo,Parameters) Sql:" + SqlString + " Error:" + ex.Message + ex.StackTrace)
            End Try
            Me.Close()
            Return DataTableObj
        End Function


        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <returns>����SqlDataReader����</returns>
        ''' <remarks></remarks>
        Public Overloads Overrides Function ExecuteSqlReturnDataReader(ByVal SqlString As String) As System.Data.SqlClient.SqlDataReader
            Me.Open(_ConnectionString)
            Dim dr As SqlClient.SqlDataReader = Nothing
            Try
                Dim _CommandObj As SqlClient.SqlCommand
                _CommandObj = CreateCmd(SqlString, _ConnectionObject)
                dr = _CommandObj.ExecuteReader()
                Return dr
            Catch ex As Exception
                _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute ExecuteSqlReturnDataReader(SqlString) Sql:" + SqlString + " Error:" + ex.Message + ex.StackTrace)
            End Try
            Return dr
            Me.Close()
        End Function

        ''' <summary>
        ''' ִ��SQL���
        ''' </summary>
        ''' <param name="SqlString">SQL���</param>
        ''' <param name="Parameters">��������</param>
        ''' <returns>����SqlDataReader����</returns>
        ''' <remarks></remarks>
        Public Overloads Overrides Function ExecuteSqlReturnDataReader(ByVal SqlString As String, ByVal Parameters As SqlClient.SqlParameter()) As System.Data.SqlClient.SqlDataReader
            Me.Open(_ConnectionString)
            Dim dr As SqlClient.SqlDataReader = Nothing
            Try
                Dim _CommandObj As SqlClient.SqlCommand
                _CommandObj = CreateCmd(SqlString, _ConnectionObject, Parameters, CommandType.Text)
                dr = _CommandObj.ExecuteReader
            Catch ex As Exception
                'Throw New Exception("Error :" + ex.Message)
                _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute ExecuteSql(SqlString,DataTableObj,Parameters()) Sql:" + SqlString + " Error:" + ex.Message + ex.StackTrace)
            End Try
            Return dr
            Me.Close()
        End Function

        '''' <summary>
        '''' ִ��SQL���
        '''' </summary>
        '''' <param name="SqlString">SQL���</param>
        '''' <param name="PageSize">ÿҳ����</param>
        '''' <param name="PageNo">ҳ��</param>
        '''' <returns>����SqlDataReader����</returns>
        '''' <remarks></remarks>
        'Public Overloads Overrides Function ExecuteSqlReturnDataReader(ByVal SqlString As String, ByVal PageSize As Integer, ByVal PageNo As Integer) As System.Data.SqlClient.SqlDataReader
        '    Me.Open(_ConnectionString)
        '    Dim dr As SqlClient.SqlDataReader = Nothing
        '    Dim da As SqlClient.SqlDataAdapter
        '    da = New SqlClient.SqlDataAdapter(SqlString, _ConnectionObject)
        '    Dim StartID As Integer
        '    StartID = PageNo * PageSize - (PageSize - 1)
        '    If StartID - 1 < 1 Then StartID = 1
        '    Try
        '        da.Fill(StartID - 1, PageSize, DataTableObj)
        '    Catch ex As Exception
        '        ' Throw New Exception("Error :" + ex.Message)
        '        _Log.Fatal("[DatabaseOperation.DataBase_SQLServer]Execute ExecuteSql(SqlString,DataTableObj,PageSize,PageNo) Error:" + ex.Message + ex.StackTrace)
        '    End Try
        '    Me.Close()
        '    Return DataTableObj
        'End Function

        '''' <summary>
        '''' ִ��SQL���
        '''' </summary>
        '''' <param name="SqlString">SQL���</param>
        '''' <param name="PageSize">ÿҳ����</param>
        '''' <param name="PageNo">ҳ��</param>
        '''' <param name="Parameters">��������</param>
        '''' <returns>����SqlDataReader����</returns>
        '''' <remarks></remarks>
        'Public Overloads Overrides Function ExecuteSqlReturnDataReader(ByVal SqlString As String, ByVal PageSize As Integer, ByVal PageNo As Integer, ByVal Parameters As SqlClient.SqlParameter()) As System.Data.SqlClient.SqlDataReader

        'End Function




        '  /// <summary>
        '   /// ����SQL��䷵��DataReader
        '  /// </summary>
        '    /// <param name="SQL"></param>
        '    /// <returns>SqlDataReader����.</returns>
        '  public SqlDataReader RunProcGetReader(string SQL) 
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   SqlCommand Cmd ;
        '   Cmd = CreateCmd(SQL, Conn);
        '   SqlDataReader Dr;
        '   try
        '   {
        '    Dr = Cmd.ExecuteReader(CommandBehavior.Default);
        '   }
        '   catch
        '   {
        '    throw new Exception(SQL);
        '   }
        '   //Dispose(Conn);
        '   return Dr;
        '  }



        '  /// <summary>
        '  /// ����Command����
        '  /// </summary>
        '  /// <param name="SQL"></param>
        '  /// <returns></returns>
        '  public SqlCommand CreateCmd(string SQL)
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   SqlCommand Cmd ;
        '   Cmd = new SqlCommand(SQL, Conn);
        '   return Cmd;
        '  }

        '  /// <summary>
        '  /// ����adapter����
        '  /// </summary>
        '  /// <param name="SQL"></param>
        '  /// <param name="Conn"></param>
        '  /// <returns></returns>
        '  public SqlDataAdapter CreateDa(string SQL) 
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   SqlDataAdapter Da;
        '   Da = new SqlDataAdapter(SQL, Conn);
        '   return Da;
        '  }

        '  /// <summary>
        '  /// ����SQL���,����DataSet����
        '  /// </summary>
        '  /// <param name="procName">SQL���</param>
        '  /// <param name="prams">DataSet����</param>
        '  public DataSet RunProc(string SQL ,DataSet Ds) 
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   SqlDataAdapter Da;
        '   //Da = CreateDa(SQL, Conn);
        '   Da = new SqlDataAdapter(SQL,Conn);
        '   try
        '   {
        '    Da.Fill(Ds);
        '   }
        '   catch(Exception Err)
        '   {
        '    throw Err;
        '   }
        '   Dispose(Conn);
        '   return Ds;
        '  }

        '  /// <summary>
        '  /// ����SQL���,����DataSet����
        '  /// </summary>
        '  /// <param name="procName">SQL���</param>
        '  /// <param name="prams">DataSet����</param>
        '  /// <param name="dataReader">����</param>
        '  public DataSet RunProc(string SQL ,DataSet Ds,string tablename) 
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   SqlDataAdapter Da;
        '   Da = CreateDa(SQL);
        '   try
        '   {
        '    Da.Fill(Ds,tablename);
        '   }
        '   catch(Exception Ex)
        '   {
        '    throw Ex;
        '   }
        '   Dispose(Conn);
        '   return Ds;
        '  }

        '  /// <summary>
        '  /// ����SQL���,����DataSet����
        '  /// </summary>
        '  /// <param name="procName">SQL���</param>
        '  /// <param name="prams">DataSet����</param>
        '  /// <param name="dataReader">����</param>
        '  public DataSet RunProc(string SQL , DataSet Ds ,int  StartIndex ,int PageSize, string tablename )
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   SqlDataAdapter Da ;
        '   Da = CreateDa(SQL);
        '   try
        '   {
        '    Da.Fill(Ds, StartIndex, PageSize, tablename);
        '   }
        '   catch(Exception Ex)
        '   {
        '    throw Ex;
        '   }
        '   Dispose(Conn);
        '   return Ds;
        '  }

        '  /// <summary>
        '  /// �����Ƿ��������
        '  /// </summary>
        '  /// <returns></returns>
        '  public bool ExistDate(string SQL) 
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   SqlDataReader Dr ;
        '   Dr = CreateCmd(SQL,Conn).ExecuteReader();
        '   if (Dr.Read())
        '   {
        '    Dispose(Conn);
        '    return true;
        '   }
        '   else
        '   {
        '    Dispose(Conn);
        '    return false;
        '   }
        '  }

        '  /// <summary>
        '  /// ����SQL���ִ�н���ĵ�һ�е�һ��
        '  /// </summary>
        '  /// <returns>�ַ���</returns>
        '  public string ReturnValue(string SQL) 
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   string result;
        '   SqlDataReader Dr ;
        '   try
        '   {
        '    Dr = CreateCmd(SQL,Conn).ExecuteReader();
        '    if (Dr.Read())
        '    {
        '     result = Dr[0].ToString();
        '     Dr.Close(); 
        '    }
        '    else
        '    {
        '     result = "";
        '     Dr.Close(); 
        '    }
        '   }
        '   catch
        '   {
        '    throw new Exception(SQL);
        '   }
        '   Dispose(Conn);
        '   return result;
        '  }

        '  /// <summary>
        '  /// ����SQL����һ��,��ColumnI��,
        '  /// </summary>
        '  /// <returns>�ַ���</returns>
        '  public string ReturnValue(string SQL, int ColumnI) 
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   string result;
        '   SqlDataReader Dr ;
        '   try
        '   {
        '    Dr = CreateCmd(SQL,Conn).ExecuteReader();
        '   }
        '   catch
        '   {
        '    throw new Exception(SQL);
        '   }
        '   if (Dr.Read())
        '   {
        '    result = Dr[ColumnI].ToString();
        '   }
        '   else
        '   {
        '    result = "";
        '   }
        '   Dr.Close();
        '   Dispose(Conn);
        '   return result;
        '  }

        '  /// <summary>
        '  /// ����һ���洢����ʹ�õ�sqlcommand.
        '  /// </summary>
        '  /// <param name="procName">�洢������.</param>
        '  /// <param name="prams">�洢�����������.</param>
        '  /// <returns>sqlcommand����.</returns>
        '  public SqlCommand CreateCmd(string procName, SqlParameter[] prams) 
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   SqlCommand Cmd = new SqlCommand(procName, Conn);
        '   Cmd.CommandType = CommandType.StoredProcedure;
        '   if (prams != null) 
        '   {
        '    foreach (SqlParameter parameter in prams)
        '    {
        '     if(parameter != null)
        '     {
        '      Cmd.Parameters.Add(parameter);
        '     }
        '    }
        '   }
        '   return Cmd;
        '  }

        '  /// <summary>
        '  /// Ϊ�洢��������һ��SqlCommand����
        '  /// </summary>
        '  /// <param name="procName">�洢������</param>
        '  /// <param name="prams">�洢���̲���</param>
        '  /// <returns>SqlCommand����</returns>
        '  private SqlCommand CreateCmd(string procName, SqlParameter[] prams,SqlDataReader Dr) 
        '  {
        '   SqlConnection Conn;
        '   Conn = new SqlConnection(ConnStr);
        '   Conn.Open();
        '   SqlCommand Cmd = new SqlCommand(procName, Conn);
        '   Cmd.CommandType = CommandType.StoredProcedure;
        '   if (prams != null) 
        '   {
        '    foreach (SqlParameter parameter in prams)
        '     Cmd.Parameters.Add(parameter);
        '   }
        '   Cmd.Parameters.Add(
        '    new SqlParameter("ReturnValue", SqlDbType.Int, 4,
        '    ParameterDirection.ReturnValue, false, 0, 0,
        '    string.Empty, DataRowVersion.Default, null));

        '   return Cmd;
        '  }

        '  /// <summary>
        '  /// ���д洢����,����.
        '  /// </summary>
        '  /// <param name="procName">�洢������</param>
        '  /// <param name="prams">�洢���̲���</param>
        '  /// <param name="dataReader">SqlDataReader����</param>
        '  public void RunProc(string procName, SqlParameter[] prams, SqlDataReader Dr) 
        '  {

        '   SqlCommand Cmd = CreateCmd(procName, prams, Dr);
        '   Dr = Cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
        '   return;
        '  }

        '  /// <summary>
        '  /// ���д洢����,����.
        '  /// </summary>
        '  /// <param name="procName">�洢������</param>
        '  /// <param name="prams">�洢���̲���</param>
        '  public string RunProc(string procName, SqlParameter[] prams) 
        '  {
        '   SqlDataReader Dr;
        '   SqlCommand Cmd = CreateCmd(procName, prams);
        '   Dr = Cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection);
        '   if(Dr.Read())
        '   {
        '    return Dr.GetValue(0).ToString();
        '   }
        '   else
        '   {
        '    return "";
        '   }
        '  }

        '  /// <summary>
        '  /// ���д洢����,����dataset.
        '  /// </summary>
        '  /// <param name="procName">�洢������.</param>
        '  /// <param name="prams">�洢�����������.</param>
        '  /// <returns>dataset����.</returns>
        '  public DataSet RunProc(string procName,SqlParameter[] prams,DataSet Ds)
        '  {
        '   SqlCommand Cmd = CreateCmd(procName,prams);
        '   SqlDataAdapter Da = new SqlDataAdapter(Cmd);
        '   try
        '   {
        '    Da.Fill(Ds);
        '   } 
        '   catch(Exception Ex)
        '   {
        '    throw Ex;
        '   }
        '   return Ds;
        '  }

        ' }
        '}



    End Class


End Namespace