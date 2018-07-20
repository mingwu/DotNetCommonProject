Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Data

Namespace Utility.Pictrues
    Public Class Pictrue
        Public Shared Sub DeletePictrue(ByVal PictruePath As String)
            If isFile(PictruePath) Then
                IO.File.Delete(PictruePath)
            End If
        End Sub

        Public Shared Sub DeletePath(ByVal PictruePath As String)
            If isPath(PictruePath) Then
                Dim PathArray() As String
                Dim fPathTemp As String
                PathArray = Split(PictruePath, "\")
                fPathTemp = PictruePath 'PictruePath.Replace(PathArray(UBound(PathArray)), "")
                IO.Directory.Delete(fPathTemp, True)
            End If
        End Sub

        Public Shared Function isFile(ByVal fFileData As String) As Boolean
            If IO.File.Exists(fFileData) = False Then
                'IO.Directory.CreateDirectory(fPathTemp)
                Dim PathArray() As String
                Dim fPathTemp As String
                PathArray = Split(fFileData, "\")
                fPathTemp = fFileData.Replace(PathArray(UBound(PathArray)), "")
                isPath(fPathTemp)
                Return False
            Else
                Return True
            End If
        End Function
        Public Shared Function isPath(ByVal fPathData As String) As Boolean
            'Dim PathArray() As String
            'Dim fPathTemp As String
            'PathArray = Split(fPathData, "\")
            'fPathTemp = fPathData.Replace(PathArray(UBound(PathArray)), "")
            If IO.Directory.Exists(fPathData) = False Then
                IO.Directory.CreateDirectory(fPathData)
                Return False
            Else
                Return True
            End If
        End Function



        '��Ҫ�����Բ�Ǵ���
        Public Enum RoundRectanglePosition 
            '���Ͻ�   
            TopLeft = 1 
            '���Ͻ�  
            TopRight = 2
            '���½�  
            BottomLeft = 3
            '���½�   
            BottomRight = 4
        End Enum



        'Public Shared Function createRoundedCorner(ByVal image As Image, ByVal roundCorner As RoundRectanglePosition) As System.Drawing.Image
        '    Dim g As Graphics = Graphics.FromImage(image)
        '    '��֤ͼƬ���� 
        '    g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        '    g.InterpolationMode = InterpolationMode.HighQualityBicubic
        '    g.CompositingQuality = CompositingQuality.HighQuality
        '    Dim rect As Rectangle = New Rectangle(0, 0, image.Width, image.Height)
        '    '����Բ���ⲿ·��  
        '    Dim rectPath As Graphics = CreateRoundRectanglePath(rect, image.Width / 10, roundCorner)


        '    '    Graphics g = Graphics.FromImage(image);  
        '    '//��֤ͼƬ����  
        '    'g.SmoothingMode = SmoothingMode.HighQuality;  
        '    'g.InterpolationMode = InterpolationMode.HighQualityBicubic;  
        '    'g.CompositingQuality = CompositingQuality.HighQuality;  
        '    'Rectangle rect = new Rectangle(0, 0, image.Width, image.Height);  
        '    '//����Բ���ⲿ·��  
        '    'GraphicsPath rectPath = CreateRoundRectanglePath(rect, image.Width / 10, roundCorner);   
        '    '//Բ�Ǳ����ð�ɫ���  
        '    'Brush b = new SolidBrush(Color.White);  
        '    'g.DrawPath(new Pen(b), rectPath);  
        '    'g.FillPath(b, rectPath);  
        '    'g.Dispose();  
        '    'return image; 
        'End Function



    End Class
End Namespace
