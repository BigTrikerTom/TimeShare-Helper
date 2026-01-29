' ######################################################################
' ## Copyright (c) 2021 TimeShareIt GdbR
' ## by Thomas Steger
' ## File creation Date: 2021-1-29 04:37
' ## File update Date: 2021-8-23 18:30
' ## Filename: Helper_DbConnectLocal.vb (F:\ConDrop\ConDrop_Server\Helper_DbConnectLocal.vb)
' ## Project: ConDrop_Server
' ## Last User: stegert
' ######################################################################
'
'

Imports System.Text.RegularExpressions
Imports System.Windows.Forms

'Imports log4net.Core

Imports Microsoft.Win32

Imports MySql.Data.MySqlClient

Public Class clsHelperDbConnectLocal
    Public Shared DBRegFolder_LizenzGenerator As String = ""
    Public Shared MySqlCmd As New MySqlCommand
    Public Shared MySqlCon As New MySqlConnection
    Public Shared MySqlConStr As String = ""
    Public Shared MySqlDBIsOpen As Boolean = False

#Region "Properties"

    Public Shared _user_condrop As String
    Public Shared ReadOnly Property user_condrop() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ConDrop, "db_User")))
            _user_condrop = Helper_cCrypt.DecryptedString
            Return _user_condrop
        End Get
    End Property
    Public Shared _pass_condrop As String
    Public Shared ReadOnly Property pass_condrop() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ConDrop, "db_Password")))
            _pass_condrop = Helper_cCrypt.DecryptedString
            Return _pass_condrop
        End Get
    End Property
    Public Shared _host_condrop As String
    Public Shared ReadOnly Property host_condrop() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ConDrop, "db_Host")))
            _host_condrop = Helper_cCrypt.DecryptedString
            Return _host_condrop
        End Get
    End Property
    Public Shared _db_condrop As String
    Public Shared ReadOnly Property db_condrop() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ConDrop, "db_Databasename")))
            _db_condrop = Helper_cCrypt.DecryptedString
            Return _db_condrop
        End Get
    End Property

    Public Shared _user_progen As String
    Public Shared ReadOnly Property user_progen() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ProGenerator, "db_User")))
            _user_progen = Helper_cCrypt.DecryptedString
            Return _user_progen
        End Get
    End Property
    Public Shared _pass_progen As String
    Public Shared ReadOnly Property pass_progen() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ProGenerator, "db_Password")))
            _pass_progen = Helper_cCrypt.DecryptedString
            Return _pass_progen
        End Get
    End Property
    Public Shared _host_progen As String
    Public Shared ReadOnly Property host_progen() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ProGenerator, "db_Host")))
            _host_progen = Helper_cCrypt.DecryptedString
            Return _host_progen
        End Get
    End Property
    Public Shared _db_progen As String
    Public Shared ReadOnly Property db_progen() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ProGenerator, "db_Databasename")))
            _db_progen = Helper_cCrypt.DecryptedString
            Return _db_progen
        End Get
    End Property

    Public Shared _user_multiserver As String
    Public Shared ReadOnly Property user_multiserver() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_Multiserver, "db_User")))
            _user_multiserver = Helper_cCrypt.DecryptedString
            Return _user_multiserver
        End Get
    End Property
    Public Shared _pass_multiserver As String
    Public Shared ReadOnly Property pass_multiserver() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_Multiserver, "db_Password")))
            _pass_multiserver = Helper_cCrypt.DecryptedString
            Return _pass_multiserver
        End Get
    End Property
    Public Shared _host_multiserver As String
    Public Shared ReadOnly Property host_multiserver() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_Multiserver, "db_Host")))
            _host_multiserver = Helper_cCrypt.DecryptedString
            Return _host_multiserver
        End Get
    End Property
    Public Shared _db_multiserver As String
    Public Shared ReadOnly Property db_multiserver() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_Multiserver, "db_Databasename")))
            _db_multiserver = Helper_cCrypt.DecryptedString
            Return _db_multiserver
        End Get
    End Property


    Public Shared _db_prefix As String
    Public Shared ReadOnly Property db_prefix() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ConDrop, "db_prefix")))
            _db_prefix = Helper_cCrypt.DecryptedString
            Return _db_prefix
        End Get
    End Property
    Public Shared _db_prefix_condrop As String
    Public Shared ReadOnly Property db_prefix_condrop() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ConDrop, "db_prefix")))
            _db_prefix_condrop = Helper_cCrypt.DecryptedString
            Return _db_prefix_condrop
        End Get
    End Property
    Public Shared _db_prefix_progen As String
    Public Shared ReadOnly Property db_prefix_progen() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_ProGenerator, "db_prefix")))
            _db_prefix_progen = Helper_cCrypt.DecryptedString
            If String.IsNullOrWhiteSpace(_db_prefix_progen) Then
                _db_prefix_progen = "pro"
            End If
            Return _db_prefix_progen
        End Get
    End Property
    Public Shared _db_prefix_Multiserver As String
    Public Shared ReadOnly Property db_prefix_Multiserver() As String
        Get
            'Dim cCrypt As cCrypt = New cCrypt
            Helper_cCrypt.Decrypt(256, Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database\" & clsHelperDbConnectLocal.DBRegFolder_Multiserver, "db_prefix")))
            _db_prefix_Multiserver = Helper_cCrypt.DecryptedString
            If String.IsNullOrWhiteSpace(_db_prefix_Multiserver) Then
                _db_prefix_Multiserver = "multiserv"
            End If
            Return _db_prefix_Multiserver
        End Get
    End Property

#End Region

#Region "Definitions"

    Public Structure ConnStringDef
        Public  MySqlConString As String
        Public  TestServer As String
        Sub New(ByVal Optional constr As String = "",
                       ByVal Optional TestServer As String = "")
            MySqlConString = constr
            TestServer = TestServer
        End Sub
    End Structure
    Private Shared ReadOnly LastConnectError As Integer = 0
    Protected  Shared disposed As Boolean = False

    Public Shared DBRegFolder_ConDrop As String = ""
    Public Shared DBRegFolder_ProGenerator As String = ""
    Public Shared DBRegFolder_Multiserver As String = ""

    Public Shared DB_Query As String = ""
    Public Shared DB_CountQuery As String = ""

    Public Shared MySqlCommandText As String = ""
    Public Shared ConDropConnectionString As String = ""
    Public Shared ProgenConnectionString As String = ""
    Public Shared MultiserverConnectionString As String = ""
    Public Shared UpdateConnectionString As String = ""
    Public Shared LicenseConnectionString As String = ""
    Public Shared ProStrukturaConnectionString As String = ""

    Public Shared CmdTimeout As Integer = 0
    Public Shared cmd_test As New MySqlCommand
    Public Shared con_test As New MySqlConnection

    '? Tabellen der ConDrop-DB
    Public Shared db_table_condrop_return_tracking As String = ""
    Public Shared db_table_condrop_Settings As String = ""
    Public Shared db_table_condrop_Absender As String = ""
    Public Shared db_table_condrop_Absender_Shipping As String = ""
    Public Shared db_table_condrop_blocked_alignment As String = ""
    Public Shared db_table_condrop_country As String = ""
    Public Shared db_table_condrop_country_states As String = ""
    Public Shared db_table_condrop_easylog_tracking As String = ""
    Public Shared db_table_condrop_email As String = ""
    Public Shared db_table_condrop_email_attachements As String = ""
    Public Shared db_table_condrop_email_folder As String = ""
    Public Shared db_table_condrop_email_placeholder As String = ""
    Public Shared db_table_condrop_email_rules As String = ""
    Public Shared db_table_condrop_email_templates As String = ""
    Public Shared db_table_condrop_history As String = ""
    Public Shared db_table_condrop_Logs As String = ""
    Public Shared db_table_condrop_msg As String = ""
    Public Shared db_table_condrop_order_incomming As String = ""
    Public Shared db_table_condrop_Origin As String = ""
    Public Shared db_table_condrop_Platforms As String = ""
    Public Shared db_table_condrop_retoure As String = ""
    Public Shared db_table_condrop_return_classification As String = ""
    Public Shared db_table_condrop_saved_addresses As String = ""
    Public Shared db_table_condrop_shipping_label As String = ""
    Public Shared db_table_condrop_shipping_scans As String = ""
    Public Shared db_table_condrop_order_shipping As String = ""
    Public Shared db_table_condrop_statistics As String = ""
    Public Shared db_table_condrop_stock_article As String = ""
    Public Shared db_table_condrop_ebay_buyer As String = ""
    Public Shared db_table_condrop_plentysalesorderreferrer As String = ""
    Public Shared db_table_condrop_import_prodws As String = ""
    Public Shared db_table_condrop_prodws As String = ""
    Public Shared db_table_condrop_prodws_active As String = ""
    Public Shared db_table_condrop_Absender_Internetmarke As String = ""
    Public Shared db_table_condrop_email_placeholder_exclude As String = ""
    Public Shared db_table_condrop_storage_def As String = ""
    Public Shared db_table_condrop_order_primelabel As String = ""
    Public Shared db_table_condrop_AmazonMws As String = ""
    Public Shared db_table_condrop_lastupdated_cache As String = ""

    '? Tabellen der Progen-DB
    Public Shared db_table_progen_Absender As String = ""
    Public Shared db_table_progen_colors As String = ""
    Public Shared db_table_progen_groups As String = ""
    Public Shared db_table_progen_products_groups As String = ""
    Public Shared db_table_progen_Prices As String = ""
    Public Shared db_table_progen_products As String = ""
    Public Shared db_table_progen_user_groups As String = ""
    Public Shared db_table_progen_user As String = ""

    '? Tabellen der Progen/ConDrop-DB
    Public Shared db_table_progen_condrop_amazon_flatfiles As String = ""
    Public Shared db_table_progen_condrop_amazon_flatfiles_default As String = ""
    Public Shared db_table_progen_condrop_amazon_subgroups_flatfiles As String = ""
    Public Shared db_table_progen_condrop_amazon_subgroups_flatfiles_values As String = ""
    Public Shared db_table_progen_condrop_amazon_validvalues As String = ""
    Public Shared db_table_progen_condrop_csv_flatfiles As String = ""
    Public Shared db_table_progen_condrop_csv_flatfiles_default As String = ""
    Public Shared db_table_progen_condrop_ebay_categories As String = ""
    Public Shared db_table_progen_condrop_ebay_flatfiles As String = ""
    Public Shared db_table_progen_condrop_ebay_flatfiles_default As String = ""
    Public Shared db_table_progen_condrop_motive_manufacturer As String = ""
    Public Shared db_table_progen_condrop_plenty_flatfiles As String = ""
    Public Shared db_table_progen_condrop_plenty_flatfiles_default As String = ""
    Public Shared db_table_progen_condrop_products_groups As String = ""
    Public Shared db_table_progen_condrop_products_master As String = ""
    Public Shared db_table_progen_condrop_products_master_groups As String = ""
    Public Shared db_table_progen_condrop_products_master_subgroups As String = ""
    Public Shared db_table_progen_condrop_products_prices As String = ""
    Public Shared db_table_progen_condrop_products_subgroups As String = ""
    Public Shared db_table_progen_condrop_products_subgroups_prices As String = ""
    Public Shared db_table_progen_condrop_size_groups As String = ""
    Public Shared db_table_ean As String = ""
    Public Shared db_table_export As String = ""

    '? Tabellen der Multiserver-DB
    Public Shared db_table_multiserver_known_server As String = ""
    Public Shared db_table_multiserver_order_incomming As String = ""
    Public Shared db_table_multiserver_order_transfer As String = ""
    Public Shared db_table_multiserver_products_base As String = ""

    '? Tabellen der Updater-DB
    Public Shared db_table_updates As String = ""
    Public Shared db_table_updater_logs As String = ""
    Public Shared db_table_updater_changelog As String = ""
    Public Shared db_table_checksums As String = ""
#End Region

    Public Enum SelectDatabase
        ConDrop
        Progenerator
        Multiserver
        Updater
        License
        Prostruktura
    End Enum


    Public Shared Sub GetDBValues()
        Dim rbrwv As Boolean = False

        Try
            DBRegFolder_ConDrop = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database", "DBRegFolder_ConDrop"))
            DBRegFolder_ProGenerator = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database", "DBRegFolder_ProGenerator"))
            DBRegFolder_Multiserver = Helper_VarConvert.ConvertToString(helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database", "DBRegFolder_Multiserver"))
            If IsNothing(DBRegFolder_ConDrop) OrElse DBRegFolder_ConDrop = "" Then
                DBRegFolder_ConDrop = "ConDrop"
                rbrwv = Helper.RegistryWriteValue(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database", "DBRegFolder_ConDrop", DBRegFolder_ConDrop)
            End If
            If IsNothing(DBRegFolder_ProGenerator) OrElse DBRegFolder_ProGenerator = "" Then
                DBRegFolder_ProGenerator = "Progenerator"
                rbrwv = Helper.RegistryWriteValue(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database", "DBRegFolder_ProGenerator", DBRegFolder_ProGenerator)
            End If
            If String.IsNullOrEmpty(DBRegFolder_Multiserver) Then
                DBRegFolder_Multiserver = "Multiserver"
                rbrwv = Helper.RegistryWriteValue(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database", "DBRegFolder_Multiserver", DBRegFolder_Multiserver)
            End If

            MySqlCommandText = "set net_write_timeout=31536000; set net_read_timeout=31536000; UseCompression=True;"
            'MySqlCommandText = "set net_write_timeout=31536000; set net_read_timeout=31536000; UseCompression=True;SslMode=none;"
            Dim ConnectStringAddOn As String = ""
            ConnectStringAddOn = "ConvertZeroDatetime=True; UseCompression=True; default command timeout=60; Allow User Variables=true; pooling=true; persistsecurityinfo=True;"
            'ConnectStringAddOn = "ConvertZeroDatetime=True; UseCompression=True; default command timeout=60; Allow User Variables=true; pooling=true; persistsecurityinfo=True;SslMode=none;"
            ConDropConnectionString = "Server=" & host_condrop & "; UID=" & user_condrop & "; Password=" & pass_condrop & "; Database=" & db_condrop & ";" & ConnectStringAddOn
            ProgenConnectionString = "Server=" & host_progen & "; UID=" & user_progen & "; Password=" & pass_progen & "; Database=" & db_progen & ";" & ConnectStringAddOn
            MultiserverConnectionString = "Server=" & host_multiserver & "; UID=" & user_multiserver & "; Password=" & pass_multiserver & "; Database=" & db_multiserver & ";" & ConnectStringAddOn
            'UpdateConnectionString = "Server=" & clsMain.host_upd & "; UID=" & clsMain.user_upd & "; Password=" & clsMain.pass_upd & "; Database=" & clsMain.db_upd & ";" & ConnectStringAddOn

            Dim DBQueryTimeOut As String = helper.SelectReadSettingsFrom(Helper.RegistryHiveValue, Helper.RegistryPath & "\Database", "DBQueryTimeOut").ToString
            If String.IsNullOrEmpty(DBQueryTimeOut) Then
                CmdTimeout = 86400
            Else
                CmdTimeout = CInt(Mid(DBQueryTimeOut, 1, 2).ToString.Trim)
            End If

            Dim rb1 As Boolean = False
            Dim rb2 As Boolean = False
            Dim rb3 As Boolean = False
            If user_condrop = "" OrElse pass_condrop = "" OrElse host_condrop = "" OrElse db_condrop = "" Then
                rb1 = False
            Else
                rb1 = ConnectionCheck(user_condrop, pass_condrop, host_condrop, db_condrop)
            End If
            If user_progen = "" OrElse pass_progen = "" OrElse host_progen = "" OrElse db_progen = "" Then
                rb2 = False
            Else
                rb2 = ConnectionCheck(user_progen, pass_progen, host_progen, db_progen)
            End If
            If user_multiserver = "" OrElse pass_multiserver = "" OrElse host_multiserver = "" OrElse db_multiserver = "" Then
                rb3 = False
            Else
                rb3 = ConnectionCheck(user_multiserver, pass_multiserver, host_multiserver, db_multiserver)
            End If
            clsMain.IsMultiserverEnabled = rb3

            '? Tabellen der ConDrop-DB
            db_table_condrop_return_tracking = db_prefix_condrop & "_return_tracking"
            db_table_condrop_Settings = db_prefix_condrop & "_settings"
            db_table_condrop_Absender = db_prefix_condrop & "_absender"
            db_table_condrop_Absender_Shipping = db_prefix_condrop & "_absender_shipping"
            db_table_condrop_blocked_alignment = db_prefix_condrop & "_blocked_alignment"
            db_table_condrop_country = db_prefix_condrop & "_country"
            db_table_condrop_country_states = db_prefix_condrop & "_country_states"
            db_table_condrop_easylog_tracking = db_prefix_condrop & "_easylog_tracking"
            db_table_condrop_email = db_prefix_condrop & "_email"
            db_table_condrop_email_attachements = db_prefix_condrop & "_email_attachements"
            db_table_condrop_email_folder = db_prefix_condrop & "_email_folder"
            db_table_condrop_email_placeholder = db_prefix_condrop & "_email_placeholder"
            db_table_condrop_email_rules = db_prefix_condrop & "_email_rules"
            db_table_condrop_email_templates = db_prefix_condrop & "_email_templates"
            db_table_condrop_history = db_prefix_condrop & "_order_history"
            db_table_condrop_Logs = db_prefix_condrop & "_logs"
            db_table_condrop_msg = db_prefix_condrop & "_msg"
            db_table_condrop_order_incomming = db_prefix_condrop & "_order_incomming"
            db_table_condrop_Origin = db_prefix_condrop & "_origin"
            db_table_condrop_Platforms = db_prefix_condrop & "_platforms"
            db_table_condrop_retoure = db_prefix_condrop & "_retoure"
            db_table_condrop_return_classification = db_prefix_condrop & "_return_classification"
            db_table_condrop_saved_addresses = db_prefix_condrop & "_saved_addresses"
            db_table_condrop_shipping_label = db_prefix_condrop & "_shipping_label"
            db_table_condrop_shipping_scans = db_prefix_condrop & "_shipping_scans"
            db_table_condrop_order_shipping = db_prefix_condrop & "_order_shipping"
            db_table_condrop_statistics = db_prefix_condrop & "_statistics"
            db_table_condrop_stock_article = db_prefix_condrop & "_stock_article"
            db_table_condrop_ebay_buyer = db_prefix_condrop & "_eBayBuyer"
            db_table_condrop_plentysalesorderreferrer = db_prefix_condrop & "_plentysalesorderreferrer"
            db_table_condrop_import_prodws = db_prefix_condrop & "_prodws_import"
            db_table_condrop_prodws = db_prefix_condrop & "_prodws"
            db_table_condrop_prodws_active = db_prefix_condrop & "_prodws_active"
            db_table_condrop_Absender_Internetmarke = db_prefix_condrop & "_absender_internetmarke"
            db_table_condrop_email_placeholder_exclude = db_prefix_condrop & "_email_placeholder_exclude"
            db_table_condrop_storage_def = db_prefix_condrop & "_storage_def"
            db_table_condrop_order_primelabel = db_prefix_condrop & "_order_primelabel"
            db_table_condrop_AmazonMws = db_prefix_condrop & "_AmazonMWS"
            db_table_condrop_lastupdated_cache = db_prefix_condrop & "_lastupdated_cache"

            '? Tabellen der Progen-DB
            db_table_ean = db_prefix_progen & "_ean"
            db_table_export = db_prefix_progen & "_export"
            db_table_progen_Absender = db_prefix_progen & "_absender"
            db_table_progen_colors = db_prefix_progen & "_colors"
            db_table_progen_groups = db_prefix_progen & "_groups"
            db_table_progen_products_groups = db_prefix_progen & "_products_groups"
            db_table_progen_Prices = db_prefix_progen & "_products_prices"
            db_table_progen_products = db_prefix_progen & "_products"
            db_table_progen_user_groups = db_prefix_progen & "_groups"
            db_table_progen_user = db_prefix_progen & "_user"

            '? Tabellen der Progen/ConDrop-DB
            db_table_progen_condrop_amazon_flatfiles = db_prefix_progen & "_condrop_amazon_flatfiles"
            db_table_progen_condrop_amazon_flatfiles_default = db_prefix_progen & "_condrop_amazon_flatfiles_default"
            db_table_progen_condrop_amazon_subgroups_flatfiles = db_prefix_progen & "_condrop_amazon_subgroups_flatfiles"
            db_table_progen_condrop_amazon_subgroups_flatfiles_values = db_prefix_progen & "_condrop_amazon_subgroups_flatfiles_values"
            db_table_progen_condrop_amazon_validvalues = db_prefix_progen & "_condrop_amazon_valid_values"
            db_table_progen_condrop_csv_flatfiles = db_prefix_progen & "_condrop_csv_flatfiles"
            db_table_progen_condrop_csv_flatfiles_default = db_prefix_progen & "_condrop_csv_flatfiles_default"
            db_table_progen_condrop_ebay_categories = db_prefix_progen & "_condrop_ebay_categories"
            db_table_progen_condrop_ebay_flatfiles = db_prefix_progen & "_condrop_ebay_flatfiles"
            db_table_progen_condrop_ebay_flatfiles_default = db_prefix_progen & "_condrop_ebay_flatfiles_default"
            db_table_progen_condrop_motive_manufacturer = db_prefix_progen & "_condrop_motive_manufacturer"
            db_table_progen_condrop_plenty_flatfiles = db_prefix_progen & "_condrop_plenty_flatfiles"
            db_table_progen_condrop_plenty_flatfiles_default = db_prefix_progen & "_condrop_plenty_flatfiles_default"
            db_table_progen_condrop_products_groups = db_prefix_progen & "_condrop_products_groups"
            db_table_progen_condrop_products_master = db_prefix_progen & "_condrop_products_master"
            db_table_progen_condrop_products_master_groups = db_prefix_progen & "_condrop_products_master_groups"
            db_table_progen_condrop_products_master_subgroups = db_prefix_progen & "_condrop_products_master_subgroups"
            db_table_progen_condrop_products_prices = db_prefix_progen & "_condrop_products_prices"
            db_table_progen_condrop_products_subgroups = db_prefix_progen & "_condrop_products_subgroups"
            db_table_progen_condrop_products_subgroups_prices = db_prefix_progen & "_condrop_products_subgroups_prices"
            db_table_progen_condrop_size_groups = db_prefix_progen & "_condrop_size_groups"

            '? Tabellen der Multiserver-DB
            db_table_multiserver_known_server = db_prefix_Multiserver & "_known_server"
            db_table_multiserver_order_incomming = db_prefix_Multiserver & "_order_incomming"
            db_table_multiserver_order_transfer = db_prefix_Multiserver & "_order_transfer"
            db_table_multiserver_products_base = db_prefix_Multiserver & "_products_"

            '? Tabellen der Updater-DB
            db_table_updates = "updates"
            db_table_updater_logs = "update_logs"
            db_table_updater_changelog = "changelog"
            db_table_checksums = "checksums"


        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False)
            If Helper.IsIDE() Then Stop
        End Try
    End Sub

    Public Shared Function GetConnectionString(ByVal ConnectionString As String) As ConnStringDef
        Dim RetVal As New ConnStringDef
        Try
            RetVal.MySqlConString = ConnectionString
            Dim reg3 As Regex = New Regex("(Server=)(.*)(; UID.*)",
                                          RegexOptions.IgnoreCase Or RegexOptions.Singleline)
            Dim res3 As Match = reg3.Match(ConnectionString)
            If (res3.Success) Then
                RetVal.TestServer = res3.Groups(2).ToString
            End If

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return RetVal
    End Function
    Public Shared Function ConnectionCheck(ByVal user As String,
                                       ByVal pass As String,
                                       ByVal host As String,
                                       ByVal db As String,
                                       ByVal Optional port As Integer = 3306,
                                       ByVal Optional ShowMsgBox As Boolean = False) As Boolean
        Dim CheckResult As New Helper_CheckConnect.CheckResults
        Dim DbCredential As New Helper_CheckConnect.DbCredentials

        Try
            Helper_CheckConnect.CheckInternet = True
            Helper_CheckConnect.CheckRegistry = True
            Helper_CheckConnect.CheckDb = True

            DbCredential.user = user
            DbCredential.pass = pass
            DbCredential.host = host
            DbCredential.db = db
            DbCredential.port = port

            Call Helper_CheckConnect.DoCheck()
            CheckResult = Helper_CheckConnect.CheckResult
            If Not CheckResult.InetResult Then
                If ShowMsgBox Then
                    MessageBox.Show(CheckResult.InetMessage, CheckResult.InetMessageCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                End If
                'Call Helper_Logger.writelog(Level.Fatal, "Internetanbindung ist fehlgeschlagen." & CheckResult.InetMessage)
                Return False
                'Application.Exit()
            End If
            If Not CheckResult.DbResult Then
                If ShowMsgBox Then
                    MessageBox.Show(CheckResult.DbMessage, CheckResult.DbMessageCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                End If
                'Call Helper_Logger.writelog(Level.Fatal, "Datenbankanbindung ist fehlgeschlagen." & CheckResult.DbMessage)
                Return False
                'Application.Exit()
            End If
            If Not CheckResult.RegResult Then
                If ShowMsgBox Then
                    MessageBox.Show(CheckResult.RegMessage, CheckResult.RegMessageCaption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                End If
                'Call Helper_Logger.writelog(Level.Fatal, "Registry ist nicht lesbar." & CheckResult.RegMessage)
                Return False
                'Application.Exit()
            End If
            Return True

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
            Return False
        End Try
    End Function
    Public Shared Function TestDbConnection(user As String, pass As String, host As String, db As String) As Boolean
        Dim result As Boolean = False
        Try
            If Not IsNothing(cmd_test.Connection) Then
                If cmd_test.Connection.State = ConnectionState.Open Then
                    con_test.Close()
                End If
            End If

            con_test.ConnectionString = "Server=" & host & "; UID=" & user & "; Password=" & pass & "; Database=" & db & ";Convert Zero Datetime=True"
            cmd_test = New MySqlCommand("set net_write_timeout=99999; set net_read_timeout=99999", con_test)
            '        cmd_test.Connection = con_test
            cmd_test.CommandTimeout = CmdTimeout
            con_test.Open()

            If cmd_test.Connection.State = ConnectionState.Open Then
              'Call Helper_Logger.writelog(Level.Info, "Datenbankverbindung zu " & db & " wurde erfolgreich aufgebaut.")
                result = True
                con_test.Close()
            Else
              'Call Helper_Logger.writelog(Level.Info, "Datenbankverbindung zu " & db & " konnte nicht aufgebaut werden.")
                result = False
            End If

        Catch exm As MySqlException
            result = False
        Catch ex As Exception
            result = False
        End Try
        Return result
    End Function

    Public Shared Function SelectCaseConnection(ByVal Connection As clsHelperDbConnectLocal.SelectDatabase) As clsHelperDbConnectLocal.ConnStringDef
        Dim ResVal As New clsHelperDbConnectLocal.ConnStringDef
        Try
            Select Case Connection
                Case clsHelperDbConnectLocal.SelectDatabase.ConDrop
                    ResVal = clsHelperDbConnectLocal.GetConnectionString(ConDropConnectionString)
                Case clsHelperDbConnectLocal.SelectDatabase.Progenerator
                    ResVal = clsHelperDbConnectLocal.GetConnectionString(ProgenConnectionString)
                Case clsHelperDbConnectLocal.SelectDatabase.Multiserver
                    ResVal = clsHelperDbConnectLocal.GetConnectionString(MultiserverConnectionString)
                Case clsHelperDbConnectLocal.SelectDatabase.Updater
                    ResVal = clsHelperDbConnectLocal.GetConnectionString(UpdateConnectionString)
                Case clsHelperDbConnectLocal.SelectDatabase.License
                    ResVal = clsHelperDbConnectLocal.GetConnectionString(LicenseConnectionString)
            End Select

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return ResVal
    End Function
End Class
