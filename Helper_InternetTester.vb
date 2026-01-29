Imports System.Net.NetworkInformation
Imports TimeShare_Helper
Imports TimeShare_Error

Public Class Helper_InternetTester
    Private WithEvents pinger As New Ping

    Public Event Checkfinished(ByVal sender As Object, ByVal e As InternetCheckerEventArgs)

    ''' <summary>
    ''' Beginnt eine asynchrone Prüfung auf eine bestehende Internetverbindung.
    ''' </summary>
    ''' <param name="url">Die URL dessen Erreichbarkeit geprüft werden 
    '''   soll</param>
    ''' <param name="maxwait">Die Zeit in 1/1000 Sekunden die maximal auf eine 
    '''   Antwort gewartet werden soll</param>
    Public Sub TestAsync(ByVal url As String, ByVal maxwait As Integer)
        If pinger Is Nothing Then pinger =
            New System.Net.NetworkInformation.Ping
        pinger.SendAsync(url, maxwait, Nothing)
    End Sub

    Private Sub pinger_PingCompleted(ByVal sender As Object,
            ByVal e As System.Net.NetworkInformation.PingCompletedEventArgs) _
            Handles pinger.PingCompleted

        RaiseEvent Checkfinished(Me,
            New InternetCheckerEventArgs(e.Reply.Status = IPStatus.Success))
    End Sub

    Public Sub CancelAsyncTest()
        pinger.SendAsyncCancel()
    End Sub

    ''' <summary>
    '''   Überprüft ob eine Internetverbindung besteht.
    ''' </summary>
    ''' <param name="url">Die URL dessen Erreichbarkeit geprüft werden 
    '''   soll</param>
    ''' <param name="maxwait">Die Zeit in 1/1000 Sekunden, die maximal auf eine 
    '''   Antwort gewartet werden soll</param>
    ''' <returns>Im Erfolgsfall wird True zurückgeliefert sonst False</returns>
    ''' <remarks>Diese Funktion kehrt erst bei einer erfolgreichen Antwort oder 
    '''   nach der in maxwait angegebenen Zeit zurück</remarks>
    Public Function Test(ByVal url As String, ByVal maxwait As Integer) As Boolean
        Try
            Dim result As System.Net.NetworkInformation.PingReply =
                pinger.Send(url, maxwait)
            If result.Status = Net.NetworkInformation.IPStatus.Success Then
                Return True
            Else
                Return False
            End If
        Catch
            Return False
        End Try
    End Function
End Class

Public Class InternetCheckerEventArgs
    Inherits EventArgs

    Private _IsAvailable As Boolean

    Public Sub New(ByVal isavailable As Boolean)
        _IsAvailable = isavailable
    End Sub

    Public ReadOnly Property IsAvailable() As Boolean
        Get
            Return _IsAvailable
        End Get
    End Property
End Class
