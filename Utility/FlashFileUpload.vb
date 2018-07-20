Namespace Utility
    Public Class FlashFileUpload
        Implements Web.IHttpHandler
        Public Sub ProcessRequest(ByVal Context As System.Web.HttpContext) Implements System.Web.IHttpHandler.ProcessRequest
            If Context.Request.Files.Count > 0 Then
                For j As Integer = 0 To Context.Request.Files.Count - 1
                    Dim _UploadFile As Web.HttpPostedFile = Context.Request.Files(j)
                    If _UploadFile.ContentLength > 0 Then
                        Dim _Log As New LogUtility.Logger("Upload")
                        Dim _FileSplit() As String
                        Dim _File_Name As String
                        _FileSplit = Split(_UploadFile.FileName, "\")
                        _File_Name = Guid.NewGuid.ToString + _FileSplit(_FileSplit.Length - 1)
                        _Log.Info(String.Format("ClassÉÏ´«{0}{1} name{2}", System.Configuration.ConfigurationSettings.AppSettings("DisplayImagePath"), _File_Name, Context.Request("name")))
                        _UploadFile.SaveAs(String.Format("{0}{1}", System.Configuration.ConfigurationSettings.AppSettings("DisplayImagePath"), _File_Name))
                    End If
                Next
            End If
        End Sub

        Public ReadOnly Property IsReusable() As Boolean Implements System.Web.IHttpHandler.IsReusable
            Get
                Return False
            End Get
        End Property

    End Class
End Namespace

