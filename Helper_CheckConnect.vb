Imports System.Drawing
Imports System.Net.NetworkInformation
Imports System.Threading
Imports System.DateTime
Imports DevComponents.DotNetBar
'Imports Microsoft.Win32
''Imports log4net
Imports MySql.Data.MySqlClient
'Imports log4net.Core
Imports System.Windows.Forms

Public Class Helper_CheckConnect
#Region "Properties"
    Private Shared _checkResult As CheckResults
    Public Shared Property CheckResult() As CheckResults
        Get
            Return _checkResult
        End Get
        Set(ByVal value As CheckResults)
            _checkResult = value
        End Set
    End Property
    Private Shared _checkInternet As Boolean
    Public Shared Property CheckInternet() As Boolean
        Get
            Return _checkInternet
        End Get
        Set(ByVal value As Boolean)
            _checkInternet = value
        End Set
    End Property
    Private Shared _checkRegistry As Boolean
    Public Shared Property CheckRegistry() As Boolean
        Get
            Return _checkRegistry
        End Get
        Set(ByVal value As Boolean)
            _checkRegistry = value
        End Set
    End Property
    Private Shared _checkDb As Boolean
    Public Shared Property CheckDb() As Boolean
        Get
            Return _checkDb
        End Get
        Set(ByVal value As Boolean)
            _checkDb = value
        End Set
    End Property
    Private Shared _dbCredential As DbCredentials
    Public Shared Property DbCredential() As DbCredentials
        Get
            Return _dbCredential
        End Get
        Set(ByVal value As DbCredentials)
            _dbCredential = value
        End Set
    End Property

    Private Shared _checkResultInet As StructCheckResultInet
    Public Shared Property CheckResultInet() As StructCheckResultInet
        Get
            Return _checkResultInet
        End Get
        Set(ByVal value As StructCheckResultInet)
            _checkResultInet = value
        End Set
    End Property
    Private Shared _checkResultReg As StructCheckResultReg
    Public Shared Property CheckResultReg() As StructCheckResultReg
        Get
            Return _checkResultReg
        End Get
        Set(ByVal value As StructCheckResultReg)
            _checkResultReg = value
        End Set
    End Property
    Private Shared _checkResultDb As StructCheckResultDb
    Public Shared Property CheckResultDb() As StructCheckResultDb
        Get
            Return _checkResultDb
        End Get
        Set(ByVal value As StructCheckResultDb)
            _checkResultDb = value
        End Set
    End Property
#End Region
#Region "Structures"
    Public Structure DbCredentials
        Public user As String
        Public pass As String
        Public host As String
        Public db As String
        Public port As Integer
    End Structure
    Public Structure CheckResults
        Public InetResult As Boolean
        Public InetMessage As String
        Public InetMessageCaption As String
        Public DbResult As Boolean
        Public DbMessage As String
        Public DbMessageCaption As String
        Public RegResult As Boolean
        Public RegMessage As String
        Public RegMessageCaption As String
    End Structure
    Public Structure StructCheckResultInet
        Public InetResult As Boolean
        Public InetMessage As String
        Public InetMessageCaption As String
    End Structure
    Public Structure StructCheckResultDb
        Public DbResult As Boolean
        Public DbMessage As String
        Public DbMessageCaption As String
    End Structure
    Public Structure StructCheckResultReg
        Public RegResult As Boolean
        Public RegMessage As String
        Public RegMessageCaption As String
    End Structure
    Public Structure DBTestResult
        Public Result As Boolean
        Public Message As String
        Public Symbol As Char
        Public SymbolColor As Color
    End Structure

#End Region

#Region "Perform Cecks"
    Public Shared Sub DoCheck()
        Dim Thread_DoCheckRegistry As Thread = New Threading.Thread(AddressOf DoCheckRegistry)
        Thread_DoCheckRegistry.IsBackground = True
        Thread_DoCheckRegistry.Start()

        Dim Thread_DoCheckInternet As Thread = New Threading.Thread(AddressOf DoCheckInternet)
        Thread_DoCheckInternet.IsBackground = True
        Thread_DoCheckInternet.Start()

        Dim Thread_DoCheckDB As Thread = New Threading.Thread(AddressOf DoCheckDB)
        Thread_DoCheckDB.IsBackground = True
        Thread_DoCheckDB.Start()

        Do While Thread_DoCheckRegistry.IsAlive OrElse Thread_DoCheckInternet.IsAlive OrElse Thread_DoCheckDB.IsAlive
            Application.DoEvents()
        Loop
        Dim CheckRes As New Helper_CheckConnect.CheckResults
        CheckRes.DbMessage = CheckResultDb.DbMessage
        CheckRes.DbMessageCaption = CheckResultDb.DbMessageCaption
        CheckRes.DbResult = CheckResultDb.DbResult
        CheckRes.InetMessage = CheckResultInet.InetMessage
        CheckRes.InetMessageCaption = CheckResultInet.InetMessageCaption
        CheckRes.InetResult = CheckResultInet.InetResult
        CheckRes.RegMessage = CheckResultReg.RegMessage
        CheckRes.RegMessageCaption = CheckResultReg.RegMessageCaption
        CheckRes.RegResult = CheckResultReg.RegResult
        CheckResult = CheckRes
    End Sub
    Private Shared Sub DoCheckInternet()
        Dim cr As New StructCheckResultInet
        If CheckInternet Then
            If Not check4InternetConnect() Then
                cr.InetMessage = "Es besteht keine Internet-Verbindung!" & vbCrLf & vbCrLf & "Informieren Sie bitte Ihren Administrator."
                cr.InetMessageCaption = "Keine Internetverbindung!"
                cr.InetResult = False
            Else
                cr.InetMessage = "OK"
                cr.InetMessageCaption = "OK"
                cr.InetResult = True
            End If
        Else
            cr.InetMessage = "OK"
            cr.InetMessageCaption = "OK"
            cr.InetResult = True
        End If
        CheckResultInet = cr
    End Sub
    Private Shared Sub DoCheckRegistry()
        Dim cr As New StructCheckResultReg
        Dim jetzt As String = Now.ToString
        'Dim LocalRegPath As String = "SOFTWARE\TimeShareIT"
        If CheckRegistry Then
            Dim rb As Boolean = Helper.SelectWriteSettingsTo(Helper.RegistryHiveValue, Helper.RegistryPath, "Test", jetzt)
            Dim jetztread As String = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath, "Test"))
            If jetzt <> jetztread OrElse Not rb Then
                cr.RegMessage = "Es besteht kein Zugriff auf die Windows-Registry!" & vbCrLf & vbCrLf & "Bitte informieren Sie Ihren Administrator."
                cr.RegMessageCaption = "Registry-Zugriff"
                cr.RegResult = False
            Else
                Call Helper.RegistryDeleteKey(Helper.RegistryHiveValue, Helper.RegistryPath, "Test")
                cr.RegMessage = "OK"
                cr.RegMessageCaption = "OK"
                cr.RegResult = True
            End If
        Else
            cr.RegMessage = "OK"
            cr.RegMessageCaption = "OK"
            cr.RegResult = True
        End If
        CheckResultReg = cr
    End Sub
    Private Shared Sub DoCheckDB()
        Dim dbc As DbCredentials = DbCredential
        Dim cr As New StructCheckResultDb
        If CheckDb AndAlso dbc.host <> "" AndAlso dbc.db <> "" AndAlso dbc.user <> "" AndAlso dbc.pass <> "" Then
            If dbc.port = 0 Then
                dbc.port = 3306
                DbCredential = dbc
            End If
            Dim TestResult As DBTestResult = TestDbConnection(dbc.user, dbc.pass, dbc.host, dbc.db)
            If Not TestResult.Result Then
                cr.DbMessage = "Es besteht kein Zugriff auf die Datenbank!" & vbCrLf & vbCrLf & "Bitte informieren Sie Ihren Administrator."
                cr.DbMessageCaption = "Keine Datenbankverbindung!"
                cr.DbResult = False
            Else
                cr.DbMessage = "OK"
                cr.DbMessageCaption = "OK"
                cr.DbResult = True
            End If
        Else
            cr.DbMessage = "OK"
            cr.DbMessageCaption = "OK"
            cr.DbResult = True
        End If
        CheckResultDb = cr
    End Sub

#End Region
#Region "CheckInternet"
    Private Shared Function check4InternetConnect() As Boolean
        Dim reply As PingReply = Nothing
        Try
            Dim success As Boolean = False
            Dim HostList As New List(Of String)
            HostList.Add("google.de")
            HostList.Add("google.com")
            HostList.Add("heise.de")
            HostList.Add("Microsoft.de")


            For Each Host As String In HostList
                For i As Integer = 1 To 10
                    Dim myPing As New Ping()
                    Dim buffer As Byte() = New Byte(31) {}
                    Dim timeout As Integer = 1000
                    Dim pingOptions As New PingOptions()
                    reply = myPing.Send(Host, timeout, buffer, pingOptions)
                    If reply.Status = IPStatus.Success Then
                        success = True
                        Exit For
                    End If
                    Helper.wait(1000)
                Next
                If success Then
                    Exit For
                End If
            Next
            Return (reply.Status = IPStatus.Success)
        Catch generatedExceptionName As Exception
            Return False
        End Try
    End Function

#End Region
#Region "CheckRegistry"
    'Public Shared Function RegistryCreateKey(Hive As RegistryHive, RegPath As String) As RegistryKey
    '    Dim regKey As RegistryKey = Nothing
    '    Select Case Hive
    '        Case RegistryHive.ClassesRoot
    '            If Helper.KeyExist(Hive, RegPath) = False Then
    '                regKey = Registry.ClassesRoot.CreateSubKey(RegPath)
    '            End If
    '        Case RegistryHive.CurrentConfig
    '            If Helper.KeyExist(Hive, RegPath) = False Then
    '                regKey = Registry.CurrentConfig.CreateSubKey(RegPath)
    '            End If
    '        Case RegistryHive.CurrentUser
    '            If Helper.KeyExist(Hive, RegPath) = False Then
    '                regKey = Registry.CurrentUser.CreateSubKey(RegPath)
    '            End If
    '        Case RegistryHive.LocalMachine
    '            If Helper.KeyExist(Hive, RegPath) = False Then
    '                regKey = Registry.LocalMachine.CreateSubKey(RegPath)
    '            End If
    '        Case RegistryHive.PerformanceData
    '            If Helper.KeyExist(Hive, RegPath) = False Then
    '                regKey = Registry.PerformanceData.CreateSubKey(RegPath)
    '            End If
    '        Case RegistryHive.Users
    '            If Helper.KeyExist(Hive, RegPath) = False Then
    '                regKey = Registry.Users.CreateSubKey(RegPath)
    '            End If
    '    End Select
    '    Return regKey
    'End Function
    'Public Shared Function SelectWriteSettingsTo(Hive As RegistryHive, RegPath As String, Key As String, Value As Object) As Boolean
    '    If Value IsNot Nothing Then
    '        Dim regKey As RegistryKey = Nothing
    '        regKey = RegistryCreateKey(Hive, RegPath)
    '        Select Case Hive
    '            Case RegistryHive.ClassesRoot
    '                regKey = Registry.ClassesRoot.OpenSubKey(RegPath, True)
    '            Case RegistryHive.CurrentConfig
    '                regKey = Registry.CurrentConfig.OpenSubKey(RegPath, True)
    '            Case RegistryHive.CurrentUser
    '                regKey = Registry.CurrentUser.OpenSubKey(RegPath, True)
    '            Case RegistryHive.LocalMachine
    '                regKey = Registry.LocalMachine.OpenSubKey(RegPath, True)
    '            Case RegistryHive.PerformanceData
    '                regKey = Registry.PerformanceData.OpenSubKey(RegPath, True)
    '            Case RegistryHive.Users
    '                regKey = Registry.Users.OpenSubKey(RegPath, True)
    '        End Select
    '        Registry.SetValue(regKey.ToString, Key, Value)
    '        regKey.Close()
    '        regKey.Dispose()
    '    End If
    '    Return True
    'End Function
    'Public Shared Function RegistryReadValue(Hive As RegistryHive, RegPath As String, Key As String, Optional VariType As Helper.VarTypeEnum = Helper.VarTypeEnum.Type_String) As Object
    '    Dim regKey As RegistryKey = Nothing
    '    Dim rb As Boolean = False
    '    Dim rd As Double = 0
    '    Dim ri As Integer = 0
    '    Dim rs As String = ""
    '    Dim value As New Object
    '    Select Case Hive
    '        Case RegistryHive.ClassesRoot
    '            regKey = Registry.ClassesRoot.OpenSubKey(RegPath, True)
    '        Case RegistryHive.CurrentConfig
    '            regKey = Registry.CurrentConfig.OpenSubKey(RegPath, True)
    '        Case RegistryHive.CurrentUser
    '            regKey = Registry.CurrentUser.OpenSubKey(RegPath, True)
    '        Case RegistryHive.LocalMachine
    '            regKey = Registry.LocalMachine.OpenSubKey(RegPath, True)
    '        Case RegistryHive.PerformanceData
    '            regKey = Registry.PerformanceData.OpenSubKey(RegPath, True)
    '        Case RegistryHive.Users
    '            regKey = Registry.Users.OpenSubKey(RegPath, True)
    '    End Select
    '    If regKey Is Nothing Then
    '        value = Nothing
    '    Else
    '        value = regKey.GetValue(Key)
    '    End If
    '    If value Is Nothing Then
    '        Select Case VariType
    '            Case Helper.VarTypeEnum.Type_Boolean
    '                Return False
    '            Case Helper.VarTypeEnum.Type_Double
    '                Return 0
    '            Case Helper.VarTypeEnum.Type_Integer
    '                Return 0
    '            Case Helper.VarTypeEnum.Type_String
    '                Return ""
    '            Case Helper.VarTypeEnum.Type_DateAsString
    '                Return ""
    '            Case Helper.VarTypeEnum.Type_Date
    '                Dim DOB As Nullable(Of Date) = CDate(Now)
    '                Dim datestring As String = DOB.Value.ToString("d")
    '                DOB = Nothing
    '                Return datestring
    '            Case Else
    '                Return ""
    '        End Select
    '    Else
    '        Select Case VariType
    '            Case Helper.VarTypeEnum.Type_Boolean
    '                Return Helper_Convert.ConvertToBoolean(value, False)
    '            Case Helper.VarTypeEnum.Type_Double
    '                Return Helper_Convert.ConvertToDouble(value, 0)
    '            Case Helper.VarTypeEnum.Type_Integer
    '                Return Helper_Convert.ConvertToInteger(value, 0)
    '            Case Helper.VarTypeEnum.Type_String
    '                Return Helper_Convert.ConvertToString(value, False, "")
    '            Case Helper.VarTypeEnum.Type_DateAsString
    '                Return Helper_Convert.ConvertDateToString(value, "dd.MM.YYYY HH:mm:ss")
    '            Case Helper.VarTypeEnum.Type_Date
    '                Return Helper_Convert.ConvertToDate(value, Nothing)
    '            Case Else
    '                Return Helper_Convert.ConvertToString(value, False, "")
    '        End Select
    '    End If
    '    regKey.Close()
    '    regKey.Dispose()
    'End Function
    'Public Shared Sub RegistryDeleteKey(ByVal Hive As RegistryHive, ByVal RegPath As String, ByVal SubKey2Delete As String)
    '    RegPath = RegPath.Trim & "\" & SubKey2Delete.Trim
    '    Select Case Hive
    '        Case RegistryHive.ClassesRoot
    '            If Helper.KeyExist(Hive, RegPath) Then
    '                Registry.ClassesRoot.DeleteSubKey(RegPath, False)
    '            End If
    '        Case RegistryHive.CurrentConfig
    '            If Helper.KeyExist(Hive, RegPath) Then
    '                Registry.CurrentConfig.DeleteSubKey(RegPath, False)
    '            End If
    '        Case RegistryHive.CurrentUser
    '            If Helper.KeyExist(Hive, RegPath) Then
    '                Registry.CurrentUser.DeleteSubKey(RegPath, False)
    '            End If
    '        Case RegistryHive.LocalMachine
    '            If Helper.KeyExist(Hive, RegPath) Then
    '                Registry.LocalMachine.DeleteSubKey(RegPath, False)
    '            End If
    '        Case RegistryHive.PerformanceData
    '            If Helper.KeyExist(Hive, RegPath) Then
    '                Registry.PerformanceData.DeleteSubKey(RegPath, False)
    '            End If
    '        Case RegistryHive.Users
    '            If Helper.KeyExist(Hive, RegPath) Then
    '                Registry.Users.DeleteSubKey(RegPath, False)
    '            End If
    '    End Select
    'End Sub

#End Region
#Region "CheckDb"
    Public Shared Function TestDbConnection(user As String, pass As String, host As String, db As String) As DBTestResult
        Dim RetMess As New DBTestResult
        Dim MainFormInfo As New LabelItem

        Try
            'If frmMain.IsHandleCreated Then
            '    MainFormInfo = frmMain.Lbl_Info
            'ElseIf frmMain.IsHandleCreated Then
            '    MainFormInfo = frmMain.Lbl_Info
            'End If
            MainFormInfo.Text = "Prüfe Datenbankverbindung..."
            Using con_test As New MySqlConnection
                con_test.ConnectionString = "Server=" & host & "; UID=" & user & "; Password=" & pass & "; Database=" & db & ";Convert Zero Datetime=True;SslMode=none"
                Dim cmd_test As MySqlCommand = New MySqlCommand("set net_write_timeout=99999; set net_read_timeout=99999", con_test)
                cmd_test.CommandTimeout = clsHelperDbConnectLocal.CmdTimeout
                RetMess.Symbol = ChrW(&HF05D)
                RetMess.SymbolColor = Color.Green
                Try
                    con_test.Open()
                    RetMess.Message = "Connection established." & vbCrLf & "Server: " & host & vbCrLf & "User: " & user
                    If cmd_test.Connection.State = ConnectionState.Open Then
                        RetMess.Result = True
                        RetMess.Message = "Connection established." & vbCrLf & "Server: " & host & vbCrLf & "User: " & user
                        MainFormInfo.Text = "Die Datenbankverbindung wurde erfolgreich hergestellt."
                        'Call Helper_Logger.writelog(Level.Info, "Test-DbConnection wurde erfolgreich hergestellt.")
                        con_test.Close()
                        con_test.Dispose()
                    Else
                        RetMess.Result = False
                        MainFormInfo.Text = "Die Datenbankverbindung konnte nicht hergestellt werden."
                        'Call Helper_Logger.writelog(Level.Info, "Test-DbConnection konnte nicht hergestellt werden. Connection String: " & con_test.ConnectionString)
                        RetMess.Message = "Connection closed." & vbCrLf & "Server: " & host & vbCrLf & "User: " & user
                    End If
                Catch ex As Exception
                    RetMess.Message = ex.Message & vbCrLf & "Error: " & ex.HResult & vbCrLf & "Host: " & host & vbCrLf & "Database: " & db & vbCrLf & "User: " & user & vbCrLf & "Password: " & pass & vbCrLf
                    MainFormInfo.Text = "Die Datenbankverbindung ist mit Fehler " & ex.Message & " fehlgeschlagen."
                    'Call Helper_Logger.writelog(Level.Info, "Test-DbConnection ist fehlgeschlagen. Connection String: " & con_test.ConnectionString)
                    RetMess.Result = False
                    RetMess.Symbol = ChrW(&HF05C)
                    RetMess.SymbolColor = Color.Red
                End Try
            End Using

        Catch ex As Exception
            'Call call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId,  False, False)
            'If Helper.Iside() Then
            'Stop
            'End If
        End Try
        Return RetMess
    End Function

#End Region
#Region "HelperFunctions"
    'Public Shared Sub wait(ByVal milliseconds As Integer)
    '    Dim x As Integer = 0
    '    For x = 1 To milliseconds
    '        Threading.Thread.Sleep(1)
    '        Application.DoEvents()
    '    Next
    'End Sub
    'Public Shared Function Checkvar(ByVal var As Object, ByVal VariType As Helper.VarTypeEnum, Optional ByVal trimed As Boolean = False, Optional ByVal DefaultValue As Object = Nothing) As Object
    '    Dim Var_Type_Boolean As Boolean = False
    '    Dim Var_Type_String As String = ""
    '    Dim Var_Type_Integer As Integer = 0
    '    Dim Var_Type_Double As Double = 0
    '    Dim Var_Type_Single As Single = 0
    '    Dim Var_Type_Long As Single = 0
    '    Dim Var_Type_Date As New Date
    '    Dim var_bool As Boolean = False
    '    Dim var_dbl As Double = 0
    '    Dim var_int As Integer = 0
    '    Dim var_sng As Single = 0
    '    Try
    '        If VariType = Helper.VarTypeEnum.Type_Boolean Then
    '            Return modHelper_Extensions.ToBoolean(var)
    '        ElseIf VariType = Helper.VarTypeEnum.Type_Double Then
    '            Return modHelper_Extensions.ToDouble(var)
    '        ElseIf VariType = Helper.VarTypeEnum.Type_Integer Then
    '            Return modHelper_Extensions.ToInteger(var)
    '        ElseIf VariType = Helper.VarTypeEnum.Type_Long Then
    '            Return modHelper_Extensions.ToLong(var)
    '        ElseIf VariType = Helper.VarTypeEnum.Type_Single Then
    '            Return modHelper_Extensions.ToSingle(var)
    '        ElseIf VariType = Helper.VarTypeEnum.Type_String Then
    '            If var Is Nothing OrElse Convert.IsDBNull(var) OrElse String.IsNullOrWhiteSpace(var.ToString) Then
    '                var = ""
    '            End If
    '            If trimed Then
    '                Var_Type_String = var.ToString.Trim
    '            Else
    '                Var_Type_String = CStr(var)
    '            End If
    '            Return Var_Type_String
    '        ElseIf VariType = Helper.VarTypeEnum.Type_Date Then
    '            If var Is Nothing OrElse Convert.IsDBNull(var) OrElse String.IsNullOrWhiteSpace(var.ToString) Then
    '                Var_Type_Date = CDate("1700-01-01 00:00:00")
    '            ElseIf CStr(var) = "1700-01-01 00:00:00" Then
    '                Var_Type_Date = CDate("1700-01-01 00:00:00")
    '            Else
    '                Var_Type_Date = CDate(var)
    '            End If
    '            Return Var_Type_Date
    '        ElseIf VariType = Helper.VarTypeEnum.Type_DateAsString Then
    '            If var Is Nothing OrElse Convert.IsDBNull(var) OrElse String.IsNullOrWhiteSpace(var.ToString) Then
    '                Var_Type_String = "1700-01-01 00:00:00"
    '            ElseIf CStr(var) = "1700-01-01 00:00:00" Then
    '                Var_Type_String = "1700-01-01 00:00:00"
    '            Else
    '                Var_Type_String = Format(CDate(var), "yyyy-MM-dd HH:mm:ss")
    '            End If
    '            Return Var_Type_String
    '        Else
    '            If var Is Nothing OrElse Convert.IsDBNull(var) OrElse String.IsNullOrWhiteSpace(var.ToString) Then
    '                var = ""
    '            End If
    '            If trimed Then
    '                Var_Type_String = CStr(var)
    '            Else
    '                Var_Type_String = CStr(var)
    '            End If
    '            Return Var_Type_String
    '        End If
    '    Catch ex As Exception
    '        Call call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId,  False, False)
    '        If Helper.IsIDE() Then Stop
    '        Return ""
    '    End Try
    'End Function

#End Region



End Class
