Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Text.RegularExpressions
Namespace Utility
    ''' <summary>
    ''' ×Ö·û¹¤¾ßÀà
    ''' </summary>
    ''' <remarks></remarks>
    Public Class StringUtility
        Private Shared ReadOnly EmailRegulationExpression As String = "^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
        'Private Shared ReadOnly NumericRegulationExpression As String = "^\d*$"
        Private Shared ReadOnly NumericRegulationExpression As String = "^\d{1,7}$"
        Private Shared ReadOnly AalphaCharacters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        Private Shared ReadOnly NumericCharacters As String = "1234567890"
        Private Shared ReadOnly AalphaNumericCharacters As String = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        Private Shared ReadOnly SymbolCharacters As String = "~!@#$%^&*()_+-={}[]\\|;:\"" ',<>.?/"
        Private Shared ReadOnly IPRegulationExpression As String = "^((\d|\d\d|[0-1]\d\d|2[0-4]\d|25[0-5])\.(\d|\d\d|[0-1]\d\d|2[0-4]\d|25[0-5])\.(\d|\d\d|[0-1]\d\d|2[0-4]\d|25[0-5])\.(\d|\d\d|[0-1]\d\d|2[0-4]\d|25[0-5]))$"

        Public Shared Function ValidateEmail(ByVal Email As String) As Boolean
            If Regex.IsMatch(Email, EmailRegulationExpression) Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Shared Function IsEmail(ByVal Email As String) As Boolean
            If String.IsNullOrEmpty(Email) Then
                Return False
            End If
            If Regex.IsMatch(Email, EmailRegulationExpression) Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Shared Function GenerateRandomizeString(ByVal PossibleCharacters As String, ByVal Length As Integer) As String
            Dim _ReturnString As String = ""
            Dim _IRandNum As Integer
            Dim _Rnd As Random = New Random
            For i As Integer = 0 To Length
                _IRandNum = _Rnd.Next(PossibleCharacters.Length)
                _ReturnString += PossibleCharacters(_IRandNum)
            Next
            Return _ReturnString
        End Function


        Public Shared Function ValidateParameter(ByRef Param As String, ByVal CheckForNull As Boolean, ByVal CheckIfEmpty As Boolean, ByVal MaxSize As Integer) As Boolean
            If Param Is Nothing Then
                If CheckForNull Then
                    Return False
                End If
                Return True
            End If
            Param = Param.Trim()
            If CheckIfEmpty And Param.Length < 1 Then
                Return False
            End If
            Return True
        End Function

        Public Shared Function ValidateParameter(ByRef Param As String, ByVal CheckForNumeric As Boolean, ByVal Minimum As Integer, ByVal Maximum As Integer) As Boolean
            If String.IsNullOrEmpty(param) Then
                Return False
            End If

            Try
                Dim _Temp As Integer = Convert.ToInt32(Param)
                If (_Temp >= Minimum) And (_Temp <= Maximum) Then
                    Return True
                Else
                    Return False
                End If
            Catch
                Return False
            End Try

        End Function

        Public Shared Function IsNumber(ByVal NumberString As String) As Boolean
            Dim _NumberRegex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(NumericRegulationExpression)

            If NumberString = String.Empty Then
                Return False
            End If
            Dim _MatchNumber As System.Text.RegularExpressions.Match = _NumberRegex.Match(NumberString)

            If _MatchNumber.Success Then
                Return True
            Else
                Return False
            End If
            'Dim _Flag As Integer = 0
            'Dim _Str() As Char = s.ToCharArray()
            'Dim i As Integer
            'For i = 0 To _Str.Length - 1 Step i + 1
            '    If Char.IsNumber(Str(i)) Then
            '        _Flag = _Flag + 1
            '    Else
            '        _Flag = -1
            '        Exit For
            '    End If
            'Next
            'If _Flag > 0 Then
            '    Return True
            'Else
            '    Return False
            'End If
        End Function

        Public Shared Function IsIP(ByVal IP As String) As Boolean
            If Regex.IsMatch(IP, IPRegulationExpression) Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' ¼ì²âÊäÈëÊÇ·ñÎª×ÖÄ¸ºÍÊý×Ö
        ''' </summary>
        ''' <param name="LetterNumberString">´ý¼ì²â×Ö·û´®</param>
        ''' <returns>·µ»Ø TRUE FALSE</returns>
        ''' <remarks></remarks>
        Public Shared Function IsLetterNumber(ByVal LetterNumberString As String) As Boolean
            Dim _LetterNumberRegex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex(AalphaNumericCharacters)
            Dim _MatchLetterNumber As System.Text.RegularExpressions.Match = _LetterNumberRegex.Match(LetterNumberString)
            If _MatchLetterNumber.Success Then
                Return True
            Else
                Return False
            End If
        End Function
    End Class
End Namespace

