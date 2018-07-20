Namespace Utility
    Public Class GetHTMLFromUrl
        Dim _Log As New LogUtility.Logger("Utility")
        Public Function GetContentFromUrl(ByVal RequestUrl As String) As String
            Dim _StrResponse As String = ""
            'Try
            '    Dim _Buf(38191) As Byte
            '    Dim _WebRequest As System.Net.HttpWebRequest = DirectCast(Net.WebRequest.Create(RequestUrl), Net.HttpWebRequest)
            '    _WebRequest.Method = "GET"
            '    Dim _WebResponse As System.Net.WebResponse = _WebRequest.GetResponse()
            '    Dim _ResponseStream As IO.Stream = _WebResponse.GetResponseStream() ' New IO.StreamReader(_WebResponse.GetResponseStream(), System.Text.Encoding.GetEncoding("gb2312"))
            '    Dim _Count As Integer = _ResponseStream.Read(_Buf, 0, _Buf.Length) '
            '    _StrResponse = Text.Encoding.UTF8.GetString(_Buf, 0, _Count) ' _ResponseStream.ReadToEnd()
            '    _WebResponse.Close()
            '    _ResponseStream.Close()
            '    'Return _StrResponse
            '    _Log.Info("[Utility.GetHTMLFromUrl]Execute GetContentFromUrll()")
            'Catch ex As Exception
            '    _Log.Error("[Utility.GetHTMLFromUrl]Execute GetContentFromUrll() RequestUrl " + RequestUrl + " Error:" + ex.Message.ToString)
            'Finally

            'End Try
            Try
                Dim request As Net.WebRequest = Net.WebRequest.Create(RequestUrl)
                Dim response As Net.WebResponse = request.GetResponse()
                Dim reader As New IO.StreamReader(response.GetResponseStream(), Text.Encoding.Default)
                _StrResponse = reader.ReadToEnd()
                reader.Close()
                reader.Dispose()
                response.Close()
                _Log.Info("[Utility.GetHTMLFromUrl]Execute GetContentFromUrll(RequestUrl) RequestUrl: " + RequestUrl)
            Catch ex As Exception
                'tb.Text = ex.Message
                _Log.Error("[Utility.GetHTMLFromUrl]Execute GetContentFromUrll(RequestUrl) RequestUrl: " + RequestUrl + " Error:" + ex.Message.ToString)
            End Try
            Return _StrResponse
        End Function

        Public Function GetContentFromUrl(ByVal RequestUrl As String, ByVal Enconding As String) As String
            Dim _StrResponse As String = ""
            'Try
            '    Dim _Buf(38191) As Byte
            '    Dim _WebRequest As System.Net.HttpWebRequest = DirectCast(Net.WebRequest.Create(RequestUrl), Net.HttpWebRequest)
            '    _WebRequest.Method = "GET"
            '    Dim _WebResponse As System.Net.WebResponse = _WebRequest.GetResponse()
            '    Dim _ResponseStream As IO.Stream = _WebResponse.GetResponseStream() ' New IO.StreamReader(_WebResponse.GetResponseStream(), System.Text.Encoding.GetEncoding("gb2312"))
            '    Dim _Count As Integer = _ResponseStream.Read(_Buf, 0, _Buf.Length) '
            '    _StrResponse = Text.Encoding.UTF8.GetString(_Buf, 0, _Count) ' _ResponseStream.ReadToEnd()
            '    _WebResponse.Close()
            '    _ResponseStream.Close()
            '    'Return _StrResponse
            '    _Log.Info("[Utility.GetHTMLFromUrl]Execute GetContentFromUrll()")
            'Catch ex As Exception
            '    _Log.Error("[Utility.GetHTMLFromUrl]Execute GetContentFromUrll() RequestUrl " + RequestUrl + " Error:" + ex.Message.ToString)
            'Finally

            'End Try
            Try
                Dim request As Net.WebRequest = Net.WebRequest.Create(RequestUrl)
                Dim response As Net.WebResponse = request.GetResponse()
                Dim reader As New IO.StreamReader(response.GetResponseStream(), Text.Encoding.GetEncoding(Enconding))
                _StrResponse = reader.ReadToEnd()
                reader.Close()
                reader.Dispose()
                response.Close()
                _Log.Info("[Utility.GetHTMLFromUrl]Execute GetContentFromUrll(RequestUrl,Enconding) RequestUrl: " + RequestUrl + " Enconding:" + Enconding)
            Catch ex As Exception
                'tb.Text = ex.Message
                _Log.Error("[Utility.GetHTMLFromUrl]Execute GetContentFromUrll(RequestUrl,Enconding) RequestUrl: " + RequestUrl + " Enconding:" + Enconding + " Error:" + ex.Message.ToString)
            End Try
            Return _StrResponse
        End Function

        'Public Function GetContentFromUrlByPOST(ByVal RequestUrl As String, ByVal Parameters As System.Collections.Specialized.NameValueCollection)   tnf用这个方法得不到内容 
        '    Dim _StrResponse As String = ""
        '    Dim _StrParameters As String = ""
        '    Try
        '        For Each _ParameterName As String In Parameters.AllKeys
        '            _StrParameters = _StrParameters + "&" + _ParameterName + "=" + Parameters(_ParameterName)
        '        Next
        '        _StrParameters = Right(_StrParameters, Len(_StrParameters) - 1)
        '        Dim _ParameterByte() As Byte = Text.Encoding.ASCII.GetBytes(_StrParameters)
        '        Dim request As Net.HttpWebRequest = Net.HttpWebRequest.Create(RequestUrl)
        '        request.Method = "POST"
        '        request.ContentLength = _ParameterByte.Length
        '        request.ContentType = "application/x-www-form-urlencoded"
        '        'request.Headers.Add("HTTP_REFERER", "http://www.thenorthface.com/catalog/sc-gear/boys-shirts.html")
        '        request.Referer = "http://www.thenorthface.com/catalog/sc-gear/boys-shirts.html"
        '        request.UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; MyIE2; .NET CLR 1.1.4322)"

        '        Dim _requestStream As IO.Stream = request.GetRequestStream
        '        _requestStream.Write(_ParameterByte, 0, _ParameterByte.Length)
        '        Dim response As Net.WebResponse = request.GetResponse()
        '        Dim reader As New IO.StreamReader(response.GetResponseStream(), Text.Encoding.Default)
        '        _StrResponse = reader.ReadToEnd()
        '        reader.Close()
        '        reader.Dispose()
        '        response.Close()
        '        _Log.Info("[Utility.GetHTMLFromUrl]Execute GetContentFromUrlByPOST(RequestUrl,Parameters)")
        '    Catch ex As Exception
        '        'tb.Text = ex.Message
        '        _Log.Error("[Utility.GetHTMLFromUrl]Execute GetContentFromUrlByPOST(RequestUrl,Parameters) RequestUrl " + RequestUrl + " Error:" + ex.Message.ToString)
        '    End Try
        '    Return _StrResponse
        'End Function


    End Class
End Namespace


