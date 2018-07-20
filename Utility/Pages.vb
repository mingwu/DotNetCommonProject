Namespace Utility
    ''' <summary>
    ''' 分页类
    ''' </summary>
    ''' <remarks></remarks>
    Public Class Page

        Private Shared _ParameterStr As String = ""
        ''' <summary>
        ''' 清除参数
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub ClearParameter()
            _ParameterStr = ""
        End Sub

        ''' <summary>
        ''' 添加参数
        ''' </summary>
        ''' <param name="ParameterName">参数名</param>
        ''' <param name="ParameterValue">参数值</param>
        ''' <remarks></remarks>
        Public Shared Sub AddParameter(ByVal ParameterName As String, ByVal ParameterValue As String)
            _ParameterStr += "'" + ParameterName + "','" + ParameterValue + "' , "
        End Sub

        Public Enum languageEnum
            zhcn
            en
        End Enum

        Public Shared language As languageEnum

        'link class
        '翻页字符串　$pageno $pagesize $nextpage $previouspage $totalpage $currentpage $firstpage $lastpage

        ''' <summary>
        ''' 生成翻页
        ''' </summary>
        ''' <param name="PageNo">当前页码</param>
        ''' <param name="RecordCount">总记录数</param>
        ''' <returns>返回翻页</returns>
        ''' <remarks></remarks>
        Public Shared Function ChangePageString(ByVal PageNo As Integer, ByVal PageSize As Integer, ByVal RecordCount As Integer)
            'Dim _CultureInfo As System.Globalization.CultureInfo
            '_CultureInfo = New System.Globalization.CultureInfo(System.Globalization.CultureInfo.CurrentCulture.Name)
            'System.Threading.Thread.CurrentThread.CurrentCulture = _CultureInfo
            'System.Threading.Thread.CurrentThread.CurrentUICulture = _CultureInfo

            My.Application.ChangeCulture(System.Globalization.CultureInfo.CurrentCulture.Name)
            My.Application.ChangeUICulture(System.Globalization.CultureInfo.CurrentCulture.Name)
            If _ParameterStr <> "" Then
                _ParameterStr = Left(_ParameterStr, Len(_ParameterStr) - 2)
                _ParameterStr = " , " + _ParameterStr
            End If
            '当前第几页／共几页　　＿页ｇｏ　　　首页　上一页　下一页　　尾页
            '當前第幾頁/共幾頁     _頁  首頁 上一頁 下一頁 尾頁
            'First page   Previous page   Next page   Last page   Total:999  Current 8 page  

            Dim PageCount As Integer = Math.Round(RecordCount / PageSize + 0.455)
            Dim FirstPageString, LastPageString, ActionPageString As String
            If PageCount < 1 Then
                Return ""
            End If


            Select Case language
                Case languageEnum.zhcn
                    If PageNo < 1 Then PageNo = 1

                    If PageNo = 0 Or PageNo = 1 Then
                        FirstPageString = "首页 上一页"
                    Else
                        FirstPageString = "<a class=org href=""javascript:redirect(['PageNo','1' " + _ParameterStr + "])"">首页</a> <a class=org href=""javascript:redirect(['PageNo','" + (PageNo - 1).ToString + "' " + _ParameterStr + "])"">上一页</a>"
                    End If

                    If PageNo >= PageCount Then
                        LastPageString = " 下一页 尾页"
                    Else
                        LastPageString = "<a class=org href=""javascript:redirect(['PageNo','" + (PageNo + 1).ToString + "' " + _ParameterStr + "])"">下一页</a> <a class=org href=""javascript:redirect(['PageNo','" + PageCount.ToString + "' " + _ParameterStr + "])"">尾页</a>"
                    End If


                    ActionPageString = " <input type=""text"" id=""GoToPage""  name=""GoToPage"" value="""" class=""formtext"" style=""WIDTH: 20px""> 页 <a href=""javascript:redirect(['PageNo',document.getElementById('GoToPage').value " + _ParameterStr + "])"">GO</a>"

                    ChangePageString = "当前第 " + PageNo.ToString + " 页/共 " + PageCount.ToString + " 页　　" + ActionPageString + "　　　" + FirstPageString + "　" + LastPageString
                Case languageEnum.en
                    If PageNo < 1 Then PageNo = 1

                    If PageNo = 0 Or PageNo = 1 Then
                        FirstPageString = "First Page Previous Page"
                    Else
                        FirstPageString = "<a class=org href=""javascript:redirect(['PageNo','1' " + _ParameterStr + "])"">First Page</a> <a class=org href=""javascript:redirect(['PageNo','" + (PageNo - 1).ToString + "' " + _ParameterStr + "])"">Previous Page</a>"
                    End If

                    If PageNo >= PageCount Then
                        LastPageString = " Next Page Last Page"
                    Else
                        LastPageString = "<a class=org href=""javascript:redirect(['PageNo','" + (PageNo + 1).ToString + "' " + _ParameterStr + "])"">Next Page</a> <a class=org href=""javascript:redirect(['PageNo','" + PageCount.ToString + "' " + _ParameterStr + "])"">Last Page</a>"
                    End If


                    ActionPageString = " <input type=""text"" id=""GoToPage""  name=""GoToPage"" value="""" class=""formtext"" style=""WIDTH: 20px""> <a href=""javascript:redirect(['PageNo',document.getElementById('GoToPage').value " + _ParameterStr + "])"">GO</a>"

                    ChangePageString = "Current page: " + PageNo.ToString + " / Total Page: " + PageCount.ToString + "　　" + ActionPageString + "　　　" + FirstPageString + "　" + LastPageString
                Case Else
                    If PageNo < 1 Then PageNo = 1

                    If PageNo = 0 Or PageNo = 1 Then
                        FirstPageString = "首页 上一页"
                    Else
                        FirstPageString = "<a class=org href=""javascript:redirect(['PageNo','1' " + _ParameterStr + "])"">首页</a> <a class=org href=""javascript:redirect(['PageNo','" + (PageNo - 1).ToString + "' " + _ParameterStr + "])"">上一页</a>"
                    End If

                    If PageNo >= PageCount Then
                        LastPageString = " 下一页 尾页"
                    Else
                        LastPageString = "<a class=org href=""javascript:redirect(['PageNo','" + (PageNo + 1).ToString + "' " + _ParameterStr + "])"">下一页</a> <a class=org href=""javascript:redirect(['PageNo','" + PageCount.ToString + "' " + _ParameterStr + "])"">尾页</a>"
                    End If


                    ActionPageString = " <input type=""text"" id=""GoToPage""  name=""GoToPage"" value="""" class=""formtext"" style=""WIDTH: 20px""> 页 <a href=""javascript:redirect(['PageNo',document.getElementById('GoToPage').value " + _ParameterStr + "])"">GO</a>"

                    ChangePageString = "当前第 " + PageNo.ToString + " 页/共 " + PageCount.ToString + " 页　　" + ActionPageString + "　　　" + FirstPageString + "　" + LastPageString
            End Select





            Return ChangePageString ' + " sss " +my.Resources.Pages. My.Resources.String1 + "  " + System.Globalization.CultureInfo.CurrentCulture.Name

        End Function

        'Public Shared Function PageChangeString_Type_1(ByVal PageNo As Integer, ByVal PageCount As Integer)

        '    If _ParameterStr <> "" Then
        '        _ParameterStr = Left(_ParameterStr, Len(_ParameterStr) - 2)
        '        _ParameterStr = " , " + _ParameterStr
        '    End If
        'If PageNo < 1 Then PageNo = 1
        '    Dim i, PageEnd, j As Integer

        '    '计算页码
        '    i = CInt(PageNo - 1) \ 5

        '    If i * 5 > 0 Then
        '        PageChangeString = "<a class=org href=""javascript:redirect(['PageNo','" + CStr(i * 5) + "' " + _ParameterStr + "])"">前5页</a> "
        '    End If

        '    If (i + 1) * 5 - 1 >= PageCount Then '计算是不是在最后
        '        PageEnd = PageCount - 1
        '    Else
        '        PageEnd = (i + 1) * 5 - 1
        '    End If


        '    For j = i * 5 To PageEnd '显示页码，要是是最后的就显示剩下的页码
        '        If CInt(PageNo) = j + 1 Then
        '            PageChangeString = PageChangeString() + "<b>[" + CStr(j + 1) + "]</b> "
        '        Else
        '            PageChangeString = PageChangeString() + "<a class=org href=""javascript:redirect(['PageNo','" + CStr(j + 1) + "' " + _ParameterStr + "])"">[" + CStr(j + 1) + "]</a> "
        '        End If
        '    Next

        '    If (i + 1) * 5 < PageCount Then
        '        PageChangeString = PageChangeString() + "<a class=org href=""javascript:redirect(['PageNo','" + CStr((i + 1) * 5 + 1) + "' " + _ParameterStr + "])"">后5页</a>"
        '    End If
        '    Return PageChangeString()
        'End Function
    End Class
End Namespace
