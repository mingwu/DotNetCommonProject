Imports System.Drawing
Namespace Utility.Pictrues
    Public Class ReSizePictrue
        '�޸��ļ��Ĵ�С
        Inherits Pictrue

        Private _PictrueType As Drawing.Imaging.ImageFormat
        Public ReadOnly Property PictrueType() As Drawing.Imaging.ImageFormat
            Get
                Return _PictrueType
            End Get
        End Property



        Public Function ResizePictrue(ByVal PictruePath As String, ByVal Width As Integer, ByVal Height As Integer, ByVal DPI As Integer) As Bitmap

            If PictruePath <> "" Then

                Dim Img, ThumbnailImg As Drawing.Image
                Dim ImgHeight, ImgWidth As Integer
                Dim temp As System.Drawing.Image = System.Drawing.Image.FromFile(PictruePath)
                Dim callb As Drawing.Image.GetThumbnailImageAbort
                Img = temp


                ImgHeight = Img.Height
                ImgWidth = Img.Width


                If Width > ImgWidth Or Height > ImgHeight Then '�Ŵ�
                    If ImgWidth / ImgHeight > Width / Height Then
                        ImgWidth = Width
                        ImgHeight = Math.Round(Img.Height * Width / Img.Width, 0, MidpointRounding.AwayFromZero)
                    Else
                        ImgHeight = Height
                        ImgWidth = Math.Round(Img.Width * Height / Img.Height, 0, MidpointRounding.AwayFromZero)
                    End If
                Else
                    If ImgWidth / ImgHeight > Width / Height Then
                        ImgWidth = Width
                        ImgHeight = Math.Round(Img.Height * Width / Img.Width, 0, MidpointRounding.AwayFromZero)
                    Else
                        ImgHeight = Height
                        ImgWidth = Math.Round(Img.Width * Height / Img.Height, 0, MidpointRounding.AwayFromZero)
                    End If
                End If

                Dim imgOutput As New Bitmap(Img, ImgWidth, ImgHeight)

                'imgOutput = New Bitmap(imgOutput)
                If DPI <> 0 Then
                    imgOutput.SetResolution(DPI, DPI) '�趨dpi
                End If



                'If nWidth < myImage.Width Or nHeight < myImage.Height Then
                '    If myImage.Width / myImage.Height > nWidth / nHeight Then
                '        xMax = nWidth
                '        yMax = myImage.Height * nWidth \ myImage.Width
                '    Else
                '        yMax = nHeight
                '        xMax = myImage.Width * nHeight \ myImage.Height
                '    End If
                'Else
                '    xMax = myImage.Width
                '    yMax = myImage.Height
                'End If

                'ֻ����С
                'Dim i As Integer
                'For i = 0 To 1
                '    If Img.Width > Width Then
                '        ImgHeight = Int((Width * Img.Height) / Img.Width)
                '        ImgWidth = Width
                '    ElseIf Img.Height > Height Then
                '        ImgWidth = Int((Height * Img.Width) / Img.Height)
                '        ImgHeight = Height
                '    End If
                'Next


                _PictrueType = Img.RawFormat
                'ThumbnailImg = imgOutput.GetThumbnailImage(ImgWidth, ImgHeight, callb, New System.IntPtr)   '.Save(FilePathData + "Thumbnail_" + FileNameData, sPicType)
                temp.Dispose()
                Img.Dispose()
                Return imgOutput ' 'ThumbnailImg 
                imgOutput.Dispose()
                ThumbnailImg.Dispose()
                'Else
                '    Return DBNull.Value
            End If

        End Function

       
        Public Function ResizePictrueByZoomRotateTransform(ByVal PictruePath As String, ByVal Width As Integer, ByVal Height As Integer, ByVal Degrees As Single, ByVal IsFit As Boolean) As Bitmap

            If PictruePath <> "" Then

                Dim temp As System.Drawing.Image = System.Drawing.Image.FromFile(PictruePath)
                Return ResizePictrueByZoomRotateTransform(temp, Width, Height, Degrees, IsFit)

            End If
        End Function

        Public Function ResizePictrueByZoomRotateTransform(ByVal image As Image, ByVal Width As Integer, ByVal Height As Integer, ByVal Degrees As Single, ByVal IsFit As Boolean) As Bitmap
            Dim temp As System.Drawing.Image = image
            Dim Img As Drawing.Image
            Dim ImgHeight, ImgWidth As Integer
            Dim callb As Drawing.Image.GetThumbnailImageAbort
            Img = temp
            Dim rate As Double

            Dim _log As New LogUtility.Logger("ReSizePictrue")

            ImgHeight = Img.Height
            ImgWidth = Img.Width

            If IsFit Then
                If Img.Height / Height < Img.Width / Width Then '������Ŵ���С���� ȡС�� ������
                    rate = Convert.ToDouble(Width) / Convert.ToDouble(Img.Width)
                Else
                    rate = Convert.ToDouble(Height) / Convert.ToDouble(Img.Height)
                End If
            Else
                If Img.Height / Height > Img.Width / Width Then '������Ŵ���С���� ȡ��� ����
                    rate = Convert.ToDouble(Width) / Convert.ToDouble(Img.Width)
                Else
                    rate = Convert.ToDouble(Height) / Convert.ToDouble(Img.Height)
                End If
            End If
            '_log.Info("rate:" + rate.ToString())
            ImgHeight = Math.Round(Img.Height * rate, 0, MidpointRounding.AwayFromZero)
            ImgWidth = Math.Round(Img.Width * rate, 0, MidpointRounding.AwayFromZero)
            '_log.Info("ImgHeight:" + ImgHeight.ToString() + " ImgWidth:" + ImgWidth.ToString())


            'If Width > ImgWidth Or Height > ImgHeight Then '�Ŵ�
            '    If ImgWidth / ImgHeight > Width / Height Then
            '        ImgWidth = Width
            '        ImgHeight = Math.Round(Img.Height * Width / Img.Width, 0, MidpointRounding.AwayFromZero)
            '    Else
            '        ImgHeight = Height
            '        ImgWidth = Math.Round(Img.Width * Height / Img.Height, 0, MidpointRounding.AwayFromZero)
            '    End If
            'Else
            '    If ImgWidth / ImgHeight > Width / Height Then
            '        ImgWidth = Width
            '        ImgHeight = Math.Round(Img.Height * Width / Img.Width, 0, MidpointRounding.AwayFromZero)
            '    Else
            '        ImgHeight = Height
            '        ImgWidth = Math.Round(Img.Width * Height / Img.Height, 0, MidpointRounding.AwayFromZero)
            '    End If
            'End If


            _PictrueType = Img.RawFormat

            Dim _tempimg As New System.Drawing.Bitmap(Width, Height) ', Imaging.PixelFormat.Format32bppRgb
            Dim _graphics As Graphics = Graphics.FromImage(_tempimg)

            Dim _zoom As Single = ImgWidth / Img.Width
            Dim _degrees As Single = 0
            '��ȡ��ǰ���ڵ����ĵ�
            Dim _rect As Rectangle = New Rectangle(0, 0, _tempimg.Width, _tempimg.Height)
            Dim _center As PointF = New PointF(_rect.Width / 2, _rect.Height / 2) '��ǰ�������ĵ�

            '���ź�ͼ��Ŀ��
            Dim _zoomWidth As Single = temp.Width * _zoom
            Dim _zoomHeight As Single = temp.Height * _zoom
            '���ź��ͼ������ĵ�
            Dim _zoomCenter As PointF = New PointF(_zoomWidth / 2, _zoomHeight / 2)
            '����ͼ��ĶԽ���
            Dim _zoomDiagonal As Single = Math.Sqrt(_zoomWidth ^ 2 + _zoomHeight ^ 2)
            '�Խ��߼н�
            Dim _diagonalDegree As Single = Math.Atan(_zoomWidth / _zoomHeight) * 180 / Math.PI
            '��תDegrees�ǶȺ�ͼ�����ĵ�����
            Dim _rotateCenter As PointF = New PointF(Math.Sin((_diagonalDegree - _degrees) * Math.PI / 180) * (_zoomDiagonal / 2), Math.Cos((_diagonalDegree - _degrees) * Math.PI / 180) * (_zoomDiagonal / 2))

            Dim _offsetX As Single = 0
            Dim _offsetY As Single = 0
            _offsetX = _center.X - _rotateCenter.X
            _offsetY = _center.Y - _rotateCenter.Y

            ''����ͼƬ��ʾ����:��ͼƬ�����ĵ��봰�ڵ����ĵ�һ��
            Dim _picRect As RectangleF = New RectangleF(_offsetX / 2, _offsetY / 2, temp.Width * _zoom, temp.Width * _zoom)
            Dim _pCenter As PointF = New PointF(_picRect.X + temp.Width / 2, _picRect.Y + temp.Width / 2)

            _graphics.ResetTransform()

            '����
            _graphics.ScaleTransform(_zoom, _zoom)
            _graphics.TranslateTransform(CSng(_offsetX / _zoom) + CSng(0), CSng(_offsetY / _zoom) + CSng(0))

            '��ͼƽ����ͼƬ�����ĵ���ת
            _graphics.RotateTransform(_degrees)

            Dim imgOutput As New Bitmap(temp, temp.Width, temp.Height)
            Dim _reSetColor As Drawing.Color = imgOutput.GetPixel(1, 1)
            imgOutput.Dispose()
            _graphics.Clear(_reSetColor)
            _graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear
            _graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
            _graphics.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            Dim _attr As New Imaging.ImageAttributes
            _attr.SetColorKey(_reSetColor, _reSetColor)
            _graphics.DrawImage(temp, New Rectangle(0, 0, temp.Width, temp.Height), 0, 0, temp.Width, temp.Height, GraphicsUnit.Pixel, _attr) '
            _graphics.Dispose();
            temp.Dispose()
            Img.Dispose()
            Return _tempimg 'imgOutput ' 'ThumbnailImg 
        End Function
    End Class
End Namespace

