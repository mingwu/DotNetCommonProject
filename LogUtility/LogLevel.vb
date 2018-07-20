Imports System
Imports System.Collections.Generic
Imports System.Text
Namespace LogUtility
    Public Class LogLevel
        Public Shared OFF As Integer = -1
        Public Shared FATAL As Integer = 0
        Public Shared [ERROR] As Integer = 1
        Public Shared WARN As Integer = 2
        Public Shared INFO As Integer = 3
        Public Shared DEBUG As Integer = 4
        Public Shared ALL As Integer = 5

        Public Shared LEVEL_STRING As String() = {"FATAL", "ERROR", "WARN", "INFO", "DEBUG", "ALL"}

        Public Shared Function GetLevelByName(ByVal LevelName As String) As Integer
            Dim _LevelInt As Integer = -1
            If String.IsNullOrEmpty(LevelName) Then
                LevelName = ""
            End If
            For i As Integer = 0 To LEVEL_STRING.Length - 1
                If LevelName.ToUpper.Trim.Equals(LEVEL_STRING(i)) Then
                    _LevelInt = i
                    Exit For
                End If
            Next
            Return _LevelInt
        End Function

        Public Shared Function GetLevelByInt(ByVal LevelInt As Integer) As String
            Return LEVEL_STRING(LevelInt)
        End Function

    End Class
End Namespace

