Imports System.Collections

Public Class Setting
    Private Shared instans As Setting = Nothing

    Public Enum AspectType As Integer
        NONE = 0
        A43 = 1
        A169 = 2
        A11 = 3
        CELL = 9
    End Enum

    Public Enum HVType As Integer
        NONE = 0
        HORIZON = 1
        VERTICAL = 2
    End Enum


    Public Shared Function GetSingleton() As Setting
        If instans Is Nothing Then
            instans = New Setting()
        End If
        Return instans
    End Function

    Public Property Aspect As AspectType

    Public Property Center As HVType

    Public Property Size As HVType

    Public Property Compress As Boolean
End Class
