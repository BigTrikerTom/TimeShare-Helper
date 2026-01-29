Option Strict On

Imports Microsoft.VisualBasic.ControlChars
Imports System.Net.NetworkInformation
Imports System.Threading
'Imports CrashReporterDotNET
'Imports log4net.Core
Imports System.Windows.Forms
Imports ShowMessage
Imports System.DateTime
'Imports ShowMessage

Public Class Helper_ErrorHandling
    'Private Shared _reportCrash As New ReportCrash("")
    Private Shared rb As Boolean = False
    Public Enum ErrorLogLevel
        el_Debug
        el_Info
        el_Warn
        el_Error
        el_Fatal
    End Enum

#Region "Properties"
    Private Shared _sendCrashReport As Boolean = False
    Public Shared WriteOnly Property SendCrashReport() As Boolean
        Set(ByVal value As Boolean)
            _sendCrashReport = value
        End Set
    End Property

    Private Shared _sysMailRecipientAddress As String
    Friend Shared Property SysMailRecipientAddress() As String
        Get
            Return _sysMailRecipientAddress
        End Get
        Set(ByVal value As String)
            _sysMailRecipientAddress = value
        End Set
    End Property
    Private Shared _sysMailRecipientName As String
    Friend Shared Property SysMailRecipientName() As String
        Get
            Return _sysMailRecipientName
        End Get
        Set(ByVal value As String)
            _sysMailRecipientName = value
        End Set
    End Property

    Private Shared _sendAlwaysErrorEmail As Boolean = True

    Friend Shared Property SendAlwaysErrorEmail() As Boolean
        Get
            Return _sendAlwaysErrorEmail
        End Get
        Set(ByVal value As Boolean)
            _sendAlwaysErrorEmail = value
        End Set
    End Property


#End Region

    Public Shared Sub GeneralErrorHandler(ByVal sender As Object, ByVal e As System.Threading.ThreadExceptionEventArgs)
        Helper_ErrorHandling.HandleErrorCatch(e.Exception, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
        If Helper.IsIDE() Then
            Stop
        Else
        End If
        'End
        Application.Exit()
    End Sub
    Public Shared Sub UnhandledErrorHandler(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        Helper_ErrorHandling.HandleErrorCatch(CType(e.ExceptionObject, Exception), Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
        If Helper.IsIDE() Then
            Stop
        Else
        End If
        'End
        Application.Exit()
    End Sub

    Public Shared Function HandleErrorCatch(ByVal ex As Exception, ByVal parent_code_element_name As String, ByVal code_element_name As String, ByVal ThreadId As Integer, Optional ByVal ShowAlways As Boolean = False, Optional ByVal LogOnly As Boolean = False, Optional ByVal ZusatzInfo As String = "", Optional ByVal LogLevel As Helper_ErrorHandling.ErrorLogLevel = Helper_ErrorHandling.ErrorLogLevel.el_Error) As Boolean
        Dim st As New StackTrace(True)
        Dim errorInnerException As String = ""
        Dim errorMessage As String = ""
        Dim LogLevelText As String = ""
        st = New StackTrace(ex, True)
        Dim exceptionStackTrace As New System.Diagnostics.StackTrace(ex, True)
        Dim exceptionStackFrames As System.Diagnostics.StackFrame() = exceptionStackTrace.GetFrames()
        Dim Message As String = ""
        'Dim query As String = ""
        'Dim project_name As String = My.Application.Info.ProductName

        'If TypeOf ex Is MySqlException Then
        '    Dim exmysql As MySqlException = CType(ex, MySqlException)
        'End If

        For Each stackFrame As System.Diagnostics.StackFrame In exceptionStackFrames
            Message += [String].Format("at {0} {1} line {2} column {3} " & Environment.NewLine, If(stackFrame.GetFileName() Is Nothing, String.Empty, stackFrame.GetFileName()), stackFrame.GetMethod().ToString(), stackFrame.GetFileLineNumber(), stackFrame.GetFileColumnNumber())
        Next

        'Dim ersource As String = project_name & " -> " & parent_code_element_name & " -> " & code_element_name
        If ex.InnerException IsNot Nothing Then
            errorInnerException = ex.InnerException.ToString
        Else
            errorInnerException = ""
        End If
        If Not ex.Message = Nothing Then
            errorMessage = ex.Message.ToString
        Else
            errorMessage = ""
        End If
        Dim errorText As String = "InnerException: " & errorInnerException & Environment.NewLine &
                                  "===================" & Environment.NewLine &
                                  "Error Message: " & errorMessage & Environment.NewLine & Environment.NewLine &
                                  "===================" & Environment.NewLine &
                                  "Zusatzinformation: " & ZusatzInfo & Environment.NewLine

        Dim errorLine As Integer = st.GetFrame(0).GetFileLineNumber()
        Dim errorSource As String = ex.Source
        Dim pre As String = ""
        Select Case LogLevel
            Case Helper_ErrorHandling.ErrorLogLevel.el_Debug
                pre = "Debug: "
            Case Helper_ErrorHandling.ErrorLogLevel.el_Error
                pre = "Fehler: "
            Case Helper_ErrorHandling.ErrorLogLevel.el_Fatal
                pre = "Fataler Fehler: "
            Case Helper_ErrorHandling.ErrorLogLevel.el_Info
                pre = "Info: "
            Case Helper_ErrorHandling.ErrorLogLevel.el_Warn
                pre = "Warnung: "
            Case Else
                pre = "Fehler: "
        End Select
        Dim sError As String = pre & ex.Message & Environment.NewLine
        sError = sError & "Message: " & Tab & ex.Message & Environment.NewLine
        sError = sError & "Thread-Id: " & Tab & ThreadId.ToString & Environment.NewLine
        sError = sError & "Routine übergeben: " & Tab & code_element_name & Environment.NewLine
        sError = sError & "Routine ermittelt Name: " & Tab & ex.TargetSite.Name & Environment.NewLine
        sError = sError & "Routine ermittelt Modul: " & Tab & ex.TargetSite.Module.Name & Environment.NewLine
        sError = sError & "Quelle: " & Tab & errorSource & Environment.NewLine
        sError = sError & errorText
        If errorLine > 0 Then
            sError = sError & "Zeile: " & Tab & errorLine & Environment.NewLine & Environment.NewLine
        End If
        If Helper_ErrorHandling.check4InternetConnect() Then
            sError = sError & "Internetverbindung: " & Tab & "aufgebaut" & Environment.NewLine
        Else
            sError = sError & "Internetverbindung: " & Tab & "nicht aufgebaut" & Environment.NewLine
        End If
        sError = sError & Environment.NewLine & Environment.NewLine & "==========" & Environment.NewLine & Environment.NewLine & Environment.NewLine
        sError = sError & "Fehlermeldung: " & Tab & errorInnerException & Environment.NewLine
        sError = sError & "--------------------------------------------" & Environment.NewLine & Environment.NewLine & errorMessage
        sError = sError & Environment.NewLine & "============================================" & Environment.NewLine & Environment.NewLine
        sError = sError & "Message: " & Environment.NewLine
        sError = sError & Message
        sError = sError & Environment.NewLine & "============================================" & Environment.NewLine & Environment.NewLine
        sError = sError & "StackTrace: " & Environment.NewLine
        sError = sError & st.ToString
        Select Case LogLevel
            Case Helper_ErrorHandling.ErrorLogLevel.el_Debug
                LogLevelText = "Debug"
            Case Helper_ErrorHandling.ErrorLogLevel.el_Error
                'Call Helper_Logger.writelog(Level.Error, sError, ex)
                LogLevelText = "Allgemeiner Fehler"
            Case Helper_ErrorHandling.ErrorLogLevel.el_Fatal
                'Call Helper_Logger.writelog(Level.Fatal, sError, ex)
                LogLevelText = "Fataler Fehler"
            Case Helper_ErrorHandling.ErrorLogLevel.el_Info
                'Call Helper_Logger.writelog(Level.Info, sError, ex)
                LogLevelText = "Info"
            Case Helper_ErrorHandling.ErrorLogLevel.el_Warn
                'Call Helper_Logger.writelog(Level.Warn, sError, ex)
                LogLevelText = "Warnung"
            Case Else
                'Call Helper_Logger.writelog(Level.Error, sError, ex)
                LogLevelText = "Allgemeiner Fehler"
        End Select
        If Helper.IsIDE() OrElse ShowAlways Then
            If Not LogOnly Then
                ShowMsg(sError, Enumerations.ShowMsgImage.Alert, "Laufzeitfehler...", True, False, Sound.Houston_We_Have_A_Problem)
            End If
        End If

        Call Helper_ErrorHandling.SendErrorEmail(sError, LogLevelText)

        'If _sendCrashReport Then
        '    SendReport(ex, sError)
        'End If

        Return True
    End Function
    Public Shared Function HandleErrorCatch(ByVal errorMessage As String, ByVal parent_code_element_name As String, ByVal code_element_name As String, ByVal ThreadId As Integer, Optional ByVal ShowAlways As Boolean = False, Optional ByVal LogOnly As Boolean = False, Optional ByVal ZusatzInfo As String = "", Optional ByVal LogLevel As Helper_ErrorHandling.ErrorLogLevel = Helper_ErrorHandling.ErrorLogLevel.el_Error) As Boolean
        Dim Message As String = ""
        'Dim query As String = ""
        Dim LogLevelText As String = ""
        'Dim project_name As String = My.Application.Info.ProductName

        'Dim ersource As String = project_name & " -> " & parent_code_element_name & " -> " & code_element_name
        Dim errorText As String = "Error Message: " & errorMessage & Environment.NewLine & Environment.NewLine &
                                  "===================" & Environment.NewLine &
                                  "Zusatzinformation: " & ZusatzInfo & Environment.NewLine
        'Dim errorText As String = errorMessage & Environment.NewLine & Environment.NewLine & "Zusatzinformation: " & Environment.NewLine & "===================" & Environment.NewLine & ZusatzInfo

        Dim sError As String = "Fehler: " & errorMessage & Environment.NewLine
        sError = sError & "Thread-Id: " & Tab & ThreadId.ToString & Environment.NewLine
        sError = sError & "Routine übergeben: " & Tab & code_element_name & Environment.NewLine
        sError = sError & errorText
        If Helper_ErrorHandling.check4InternetConnect() Then
            sError = sError & "Internetverbindung: " & Tab & "aufgebaut" & Environment.NewLine
        Else
            sError = sError & "Internetverbindung: " & Tab & "nicht aufgebaut" & Environment.NewLine
        End If
        sError = sError & Environment.NewLine & Environment.NewLine & "==========" & Environment.NewLine & Environment.NewLine & Environment.NewLine
        sError = sError & Environment.NewLine & "============================================" & Environment.NewLine & Environment.NewLine
        sError = sError & "Message: " & Environment.NewLine
        sError = sError & Message
        sError = sError & Environment.NewLine & "============================================" & Environment.NewLine & Environment.NewLine
        Select Case LogLevel
            Case Helper_ErrorHandling.ErrorLogLevel.el_Debug
                LogLevelText = "Debug"
            Case Helper_ErrorHandling.ErrorLogLevel.el_Error
                'Call Helper_Logger.writelog(Level.Error, sError)
                LogLevelText = "Allgemeiner Fehler"
            Case Helper_ErrorHandling.ErrorLogLevel.el_Fatal
                'Call Helper_Logger.writelog(Level.Fatal, sError)
                LogLevelText = "Fataler Fehler"
            Case Helper_ErrorHandling.ErrorLogLevel.el_Info
                'Call Helper_Logger.writelog(Level.Info, sError)
                LogLevelText = "Info"
            Case Helper_ErrorHandling.ErrorLogLevel.el_Warn
                'Call Helper_Logger.writelog(Level.Warn, sError)
                LogLevelText = "Warnung"
            Case Else
                'Call Helper_Logger.writelog(Level.Error, sError)
                LogLevelText = "Allgemeiner Fehler"
        End Select

        'If _sendCrashReport Then
        '    SendReport(New Exception(errorMessage), sError)
        'End If

        If Helper.IsIDE() OrElse ShowAlways Then
            If Not LogOnly Then
                Try

                Catch ex As Exception
                    Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
                    If Helper.IsIDE() Then Stop
                End Try
                Try

                Catch ex As Exception
                    Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
                    If Helper.IsIDE() Then Stop
                End Try
                Try

                Catch ex As Exception
                    Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
                    If Helper.IsIDE() Then Stop
                End Try
                Try

                Catch ex As Exception
                    Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
                    If Helper.IsIDE() Then Stop
                End Try
                ShowMsg(sError, ShowMsgImage.Alert, "Laufzeitfehler...", True, False, Sound.Houston_We_Have_A_Problem)
            End If
        End If
        Call Helper_ErrorHandling.SendErrorEmail(sError, LogLevelText)

        Return True
    End Function
    Public Shared Function check4InternetConnect() As Boolean
        Dim reply As PingReply = Nothing
        Try
            For i As Integer = 1 To 10
                Dim myPing As New Ping()
                Dim host As [String] = "google.com"
                Dim buffer As Byte() = New Byte(31) {}
                Dim timeout As Integer = 1000
                Dim pingOptions As New PingOptions()
                reply = myPing.Send(host, timeout, buffer, pingOptions)
                If reply.Status = IPStatus.Success Then
                    Exit For
                End If
                Helper.wait(1000)
            Next
            Return (reply.Status = IPStatus.Success)
        Catch generatedExceptionName As Exception
            Return False
        End Try

    End Function
    Private Shared Sub SendErrorEmail(ByVal ErrorMessage As String, ByVal LogLevel As String)
        'Dim cCrypt As cCrypt = New cCrypt
        Dim NachrichtHeader As String = ""
        Dim SMTPCred As New Helper.SMTPCredentials
        Dim SMTP_RecipientAddress As String = SysMailRecipientAddress
        Dim SMTP_RecipientName As String = SysMailRecipientName
        If Not String.IsNullOrWhiteSpace(SMTP_RecipientAddress) AndAlso Not String.IsNullOrWhiteSpace(SMTP_RecipientAddress) AndAlso Not Helper.IsIDE() OrElse SendAlwaysErrorEmail Then
            SMTPCred.SMTP_Password = Helper_cCrypt.DecryptString(Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_Password")))
            SMTPCred.SMTP_User = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_User"))
            SMTPCred.SMTP_Server = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_RelayServer"))
            SMTPCred.SMTP_NoSSL = Helper_VarConvert.ConvertToBoolean(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_NoSSL"))
            SMTPCred.SMTP_SSL = Helper_VarConvert.ConvertToBoolean(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_SSL"))
            SMTPCred.SMTP_STARTTLS = Helper_VarConvert.ConvertToBoolean(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_STARTTLS"))
            SMTPCred.SMTP_SenderAddress = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_SenderAddress"))
            SMTPCred.SMTP_SenderName = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_SenderName"))
            SMTPCred.SMTP_RecipientAddress = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_RecipientAddress"))
            SMTPCred.SMTP_RecipientName = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_RecipientName"))
            SMTPCred.SMTP_CC = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_CC"))
            SMTPCred.SMTP_BCC = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_BCC"))
            SMTPCred.SMTP_NoSSL_Port = Helper_VarConvert.ConvertToInteger(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_NoSSL_Port"))
            SMTPCred.SMTP_SSL_Port = Helper_VarConvert.ConvertToInteger(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_SSL_Port"))
            SMTPCred.SMTP_STARTTLS_Port = Helper_VarConvert.ConvertToInteger(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\email", "SMTP_STARTTLS_Port"))
            SMTPCred.SMTP_RecipientAddress = SysMailRecipientAddress
            SMTPCred.SMTP_RecipientName = SysMailRecipientName

            NachrichtHeader = "Laufzeitfehler der Kategorie: " & LogLevel & " im " & Application.ProductName & ": " & System.Environment.MachineName.ToUpper
            Dim ErrorMessage_l1 As String = "Es ist im " & Application.ProductName & " auf " & System.Environment.MachineName.ToUpper & " am " & Now.ToString("dd-MM-yyyy") & " um " & Now.ToString("HH:mm:ss") & " Uhr ein Fehler aufgetreten!" & Environment.NewLine & Environment.NewLine
            Dim ErrorMessage_l2 As String = "------------------------------------------" & Environment.NewLine & Environment.NewLine
            Dim ErrorMessage1 As String = ErrorMessage_l1 & ErrorMessage_l2 & ErrorMessage.Replace(CrLf, System.Convert.ToChar(10))

            If Not String.IsNullOrEmpty(SMTPCred.SMTP_Password) AndAlso
                                                       Not String.IsNullOrEmpty(SMTPCred.SMTP_User) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_SenderAddress) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_RecipientAddress) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_RecipientAddress) AndAlso
                                                      Not String.IsNullOrEmpty(SMTPCred.SMTP_Server) Then
                rb = Helper_Mail.SendSystemEmailUniversal(Helper.RegistryHiveValue,
                    Helper.RegistryPath,
                    NachrichtHeader,
                    ErrorMessage1,
                    "",
                    New List(Of String),
                    SMTPCred)
            End If

        End If
    End Sub

    Public Shared Sub InitializeCrashReport(ByVal toEmail As String,
                                            ByVal fromEmail As String,
                                            ByVal Optional Silent As Boolean = False,
                                            ByVal Optional ShowScreenshotTab As Boolean = True,
                                            ByVal Optional IncludeScreenshot As Boolean = True,
                                            ByVal Optional CaptureScreen As Boolean = False,
                                            ByVal Optional AnalyzeWithDoctorDump As Boolean = False
                                            )
        toEmail = toEmail.Replace(" ", "-")
        fromEmail = fromEmail.Replace(" ", "-")
        '_reportCrash = New ReportCrash(toEmail) With {
        '    .FromEmail = fromEmail,
        '    .Silent = Silent,
        '    .ShowScreenshotTab = ShowScreenshotTab,
        '    .IncludeScreenshot = IncludeScreenshot,
        '    .CaptureScreen = CaptureScreen,
        '    .AnalyzeWithDoctorDump = AnalyzeWithDoctorDump
        '    }

        '.WebProxy = New WebProxy("Web proxy address, if needed"),
        '.DoctorDumpSettings = New DoctorDumpSettings With {
        '    .ApplicationID = New Guid("Application ID you received from DrDump.com"),
        '    .OpenReportInBrowser = True
        '    }
        '_reportCrash.RetryFailedReports()

    End Sub
    'Private Shared Sub SendReport(ByVal exception As Exception, ByVal Optional developerMessage As String = "")
    '    _reportCrash.DeveloperMessage = developerMessage
    '    _reportCrash.Send(exception)
    'End Sub

End Class
