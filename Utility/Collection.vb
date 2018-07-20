Namespace Utility
    Public Class Collection
        Inherits Hashtable
        ' Private _HashTable As New Hashtable

        ''' <summary>
        ''' ����Key��ArrayList
        ''' </summary>
        ''' <remarks></remarks>
        Private _Keys As New ArrayList

        ''' <summary>
        ''' ����Key����Ӧ�Ķ���
        ''' </summary>
        ''' <param name="Key">Key</param>
        ''' <returns>����Key����Ӧ�Ķ���</returns>
        ''' <remarks></remarks>
        Public Function GetItem(ByVal Key As String)
            If MyBase.ContainsKey(Key) Then
                If IsDBNull(MyBase.Item(Key)) Then
                    Return Nothing
                Else
                    Return MyBase.Item(Key)
                End If
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' ����Key����
        ''' </summary>
        ''' <value></value>
        ''' <returns>����Key����</returns>
        ''' <remarks></remarks>
        ReadOnly Property GetKeys() As ArrayList
            Get
                Return _Keys '_HashTable.Keys
            End Get
        End Property

        ''' <summary>
        ''' ��Ӷ���
        ''' </summary>
        ''' <param name="Key">����Key</param>
        ''' <param name="Value">����ֵ</param>
        ''' <remarks></remarks>
        Public Overloads Sub Add(ByVal Key As String, ByVal Value As Object)
            MyBase.Add(Key, Value)
            _Keys.Add(Key)
        End Sub

        ''' <summary>
        ''' �����������
        ''' </summary>
        ''' <remarks></remarks>
        Public Overloads Sub Clear()
            MyBase.Clear()
            _Keys.Clear()
        End Sub

        ''' <summary>
        ''' �Ƴ�Keyָ���Ķ���
        ''' </summary>
        ''' <param name="Key">����Key</param>
        ''' <remarks></remarks>
        Public Overloads Sub Remove(ByVal Key As String)
            MyBase.Remove(Key)
            _Keys.Remove(Key)
        End Sub


    End Class

    Public Class Populate
        Public Function PopulateItems(ByVal DataTable As DataTable) As Utility.Collection()
            Dim _Collection() As Utility.Collection
            Dim i, j As Integer
            _Collection = Nothing
            Array.Resize(_Collection, 0)
            Array.Resize(_Collection, DataTable.Rows.Count)
            For i = 0 To DataTable.Rows.Count - 1
                _Collection(i) = New Utility.Collection
                For j = 0 To DataTable.Columns.Count - 1
                    _Collection(i).Add(DataTable.Columns(j).ColumnName, DataTable.Rows(i).Item(DataTable.Columns(j).ColumnName))
                Next
            Next
            If DataTable.Rows.Count > 0 Then
                Return _Collection
            Else
                Array.Resize(_Collection, 0)
                Return _Collection
            End If
        End Function

        Public Function PopulateItem(ByVal DataTable As DataTable) As Utility.Collection
            Dim _Collection As New Utility.Collection
            Dim j As Integer
            If DataTable.Rows.Count > 0 Then
                _Collection = New Utility.Collection
                For j = 0 To DataTable.Columns.Count - 1
                    _Collection.Add(DataTable.Columns(j).ColumnName, DataTable.Rows(0).Item(DataTable.Columns(j).ColumnName))
                Next
                Return _Collection
            Else
                Return New Utility.Collection
            End If

        End Function
    End Class

    'Public Class Convert
    '    Public Function ConvertToDateTable(ByVal Collections() As Utility.Collection) As Data.DataTable
    '        Dim _DataTableRow As DataRow
    '        Dim _DataTable As New DataTable
    '        Dim i As Integer
    '        Dim _KeyCollection As IDictionaryEnumerator


    '        _KeyCollection = Collections.GetEnumerator
    '        While _KeyCollection.MoveNext()
    '            _DataTable.Columns.Add(New DataColumn(_KeyCollection.Key))
    '        End While

    '        For i = 0 To Collections.Length - 1
    '            _DataTableRow = _DataTable.NewRow
    '            _KeyCollection = Collections(i).GetEnumerator
    '            While _KeyCollection.MoveNext()
    '                _DataTableRow.Item(_KeyCollection.Key) = _KeyCollection.Value 'Collections(i).GetItem("content_type_name").ToString
    '            End While
    '            _DataTable.Rows.Add(_DataTableRow)
    '        Next
    '        Return _DataTable
    '    End Function
    'End Class
End Namespace
