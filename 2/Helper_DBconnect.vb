Option Strict On
Imports MySql.Data.MySqlClient
Imports System.DateTime
Imports TimeShare_Helper
Imports TimeShare_Error

Friend Class clsDBconnect
    Implements IDisposable

#Region "Definitions"
    Private disposed As Boolean = False
    Private ReadOnly Shared MySqlCommandText As String = ""
    Private ReadOnly Shared Loc_CmdTimeout As Integer = 0
    Private MySqlCon As New MySqlConnection
    Private MySqlConString As String = ""
    Private ReadOnly DbTimeoutSec As Double = 60

    Friend Structure ConnStringDef
        Friend MySqlConString As String
        Friend TestServer As String
        Friend Sub New(ByVal Optional constr As String = "",
                       ByVal Optional TestServer As String = "")
            MySqlConString = constr
            Me.TestServer = TestServer
        End Sub
    End Structure

#End Region
#Region "Properties"
    Private _MySqlCmd As New MySqlCommand
    Friend ReadOnly Property cmd As MySqlCommand
        Get
            Return _MySqlCmd
        End Get
    End Property
    Private _DBIsOpen As Boolean = False
    Friend ReadOnly Property DBIsOpen() As Boolean
        Get
            Return _DBIsOpen
        End Get
    End Property
#End Region

    Friend Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub
    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        Try
            If Not Me.disposed Then
                If disposing Then
                    'managedResource.Dispose()
                    MySqlCon.Close()
                    _MySqlCmd.Dispose()
                    MySqlCon.Dispose()
                    MySqlCon = Nothing
                    _MySqlCmd = Nothing
                    MySqlConString = Nothing
                    _DBIsOpen = False
                    'unmanagedResource = IntPtr.Zero
                End If
            End If
            Me.disposed = True

        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
    End Sub


    Friend Function connect(ByVal Connection As clsDbConnectLocal.SelectDatabase) As Boolean
        If Me.disposed Then
            Throw New ObjectDisposedException(Me.GetType().ToString, "This object has been disposed.")
        End If
        Dim ResVal As New clsDbConnectLocal.ConnStringDef
        Dim DatetimeStart As New DateTime

        Try
            DatetimeStart = Now
            ResVal = clsDbConnectLocal.SelectCaseConnection(Connection)

            MySqlConString = ResVal.MySqlConString
            'TestServer = ResVal.TestServer

            MySqlCon.ConnectionString = MySqlConString
            _MySqlCmd = New MySqlCommand(MySqlCommandText, MySqlCon)
            _MySqlCmd.CommandTimeout = Loc_CmdTimeout
            Dim exept As New Exception
            exept = Nothing

            Do
                For I As Integer = 1 To 10
                    Try
                        MySqlCon.Open()
                        _DBIsOpen = True
                        Exit For
                    Catch ex As Exception
                        'Host4InternetTest = TestServer
                        MySqlCon.Dispose()
                        _DBIsOpen = False
                        If Not ErrorHandling.check4InternetConnect() Then
                            If ex.Message <> "Timeout in IO operation" AndAlso
                               Not ex.Message.StartsWith("Authentication method") AndAlso
                               Not ex.HResult = -2146233080 Then
                                exept = ex
                            End If
                        End If
                    End Try
                    Helper.wait(500)
                Next
                If _DBIsOpen Then
                    Exit Do
                ElseIf exept IsNot Nothing Then
                    Throw (exept)
                ElseIf ErrorHandling.check4InternetConnect() AndAlso Now >= DateAdd(DateInterval.Second, DbTimeoutSec, DatetimeStart) Then
                    Exit Do
                End If
                Helper.wait(500)
            Loop
        Catch ex As MySqlException
            Dim info As String = "Connectionstring: " & MySqlConString & Environment.NewLine
            info = info & "MySqlCommandText: " & MySqlCommandText
            'ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False, info)
            _DBIsOpen = False
        End Try
        'frmMain.IsDBConnected = DBIsOpen_
        Return _DBIsOpen
    End Function
    Friend Sub close()
        Try
            If Me.disposed Then
                Throw New ObjectDisposedException(Me.GetType().ToString, "This object has been disposed.")
            End If
            _MySqlCmd.Dispose()
            MySqlCon.Close()
            MySqlCon.Dispose()
            Me.Dispose()
            'frmMain.IsDBConnected = DBIsOpen_
            _DBIsOpen = False

        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
    End Sub


End Class
