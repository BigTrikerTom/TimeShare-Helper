Imports Microsoft.Win32

Friend Class Main
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
    Private Shared _registryHiveValue As RegistryHive
    Friend Shared Property RegistryHiveValue() As RegistryHive
        Get
            Return _registryHiveValue
        End Get
        Set(ByVal value As RegistryHive)
            _registryHiveValue = value
        End Set
    End Property
    Private Shared _registryPath As String
    Friend Shared Property RegistryPath() As String
        Get
            Return _registryPath
        End Get
        Set(ByVal value As String)
            _registryPath = value
        End Set
    End Property

End Class
