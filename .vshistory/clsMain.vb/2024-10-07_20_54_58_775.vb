Imports Microsoft.Win32

Public Delegate Sub StatusCallback(ByVal message As CallbackMessage)
Public Structure CallbackMessage
    Public Msg_DateTime As Date
    Public Msg_Message As String
    Public Msg_Status As MessageStatus
End Structure
Public Enum MessageStatus
    status_success
    status_info
    status_warning
    status_error
    status_fatal
    status_other
End Enum

Public Class clsMain
    '? **************** Variablen für Auto-Updater ****************
    Friend Shared ReadOnly host_upd As String = "bayernreich.de"
    Friend Shared ReadOnly db_upd As String = "updater_db"
    Friend Shared ReadOnly user_upd As String = "updater_dbu"
    Friend Shared ReadOnly pass_upd As String = "AIhierb3eb9"
    Friend Shared ReadOnly ftp_user As String = "updater"
    Friend Shared ReadOnly ftp_pass As String = "2ig95FESunFA"
    Friend Shared ReadOnly ftp_host As String = "bayernreich.de"
    Friend Shared ReadOnly ftp_dir As String = "/condrop_server/"
    '? **************** Variablen für Auto-Updater ****************
    Friend Shared IsMultiserverEnabled As Boolean = False
End Class
