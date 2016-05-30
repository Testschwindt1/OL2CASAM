Imports System.Runtime.InteropServices
Imports Sagede.OfficeLine.Engine

<Microsoft.VisualBasic.ComClass()> Public Class Connector

    Public Shared _mandant As Mandant

    <System.Runtime.InteropServices.ComVisible(True)> _
    Public Function swiTest() As String
        swiTest = "Test MG"
    End Function

    Public Sub New()
        'MyBase.New()
    End Sub

    Public Sub InitMandant(ByVal Mandant As OLSysIInterop70.Mandant)
        If _mandant Is Nothing Then _mandant = DirectCast(Mandant, Sagede.OfficeLine.Interop70.Mandant).GetRealObject
    End Sub

    Public Sub InitMandantNET(Mandant As Mandant)
        If _mandant Is Nothing Then _mandant = Mandant
    End Sub

    Public Function CheckLizenz(ByVal Version As String) As Boolean
        If ICheckLizenz(Version) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function CheckLizenzNewUpdate(ByVal Version As String) As Boolean
        If ICheckLizenzNewUpdate(Version) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function CheckLizenzOffline() As Boolean
        If ICheckLizenzOffline() Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
