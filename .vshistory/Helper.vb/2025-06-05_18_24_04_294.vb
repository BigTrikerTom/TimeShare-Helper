' ######################################################################
' ## Copyright (c) 2021 TimeShareIt GdbR
' ## by Thomas Steger
' ## File creation Date: 2020-12-15 17:18
' ## File update Date: 2021-8-23 18:12
' ## Filename: Helper.vb (F:\++++ Code Share\classes\Helper.vb)
' ## Project: ConDrop_Server
' ## Last User: stegert
' ######################################################################
'
'

Imports DevComponents.AdvTree
Imports DevComponents.DotNetBar
Imports Microsoft.Win32
Imports System.IO.Compression
Imports System
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.Security.Principal
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.DateTime
Imports DevComponents.Editors.DateTimeAdv

Imports System.Drawing
Imports Limilabs.Mail
Imports Limilabs.Mail.Headers
Imports Limilabs.Mail.MIME
Imports System.Drawing.Imaging
Imports Limilabs.Client.SMTP
Imports System.Security.Cryptography
Imports System.Text
Imports Newtonsoft.Json.Linq
'Imports log4net.Core
Imports System.Configuration
Imports System.Drawing.Drawing2D
Imports System.Reflection
Imports System.Xml.Serialization
Imports DevComponents.DotNetBar.Controls
'Imports Spire.Pdf.General.Render.Font.OpenTypeFile
Imports System.DirectoryServices
'Imports Microsoft.VisualBasic.CompilerServices
Imports Microsoft.VisualBasic.ControlChars
Imports DirectoryEntry = System.DirectoryServices.DirectoryEntry
Imports System.Security
'Imports FeierTage
Imports System.Windows.Forms
Imports System.Xml
Imports WindowsFirewallHelper
Imports Google.Protobuf.Reflection.SourceCodeInfo.Types
Imports Ini.Net
Imports TimeShare_Error

Public Class Helper
#Region "Enum"
    Public Enum VarTypeEnu
        Type_Boolean
        Type_String
        Type_Integer
        Type_Long
        Type_Double
        Type_Single
        Type_Date
        Type_DateTime
        Type_DateAsString
        Type_Object
    End Enum
    Public Enum ToDates
        DateOnly
        TimeOnly
        DateTime
    End Enum

    Public Enum Platform
        Afterbuy = 1
        Plenty = 2
        CSV = 3
        eBay = 4
        Amazon = 5
        Manuell = 6
        Shopware = 7
        Otto = 8
    End Enum

#End Region
#Region "Structures"
    Public Structure multiScreenBounds
        Friend outerLeft As Integer
        Friend outerTop As Integer
        Friend outerRight As Integer
        Friend outerBottom As Integer
    End Structure
    Public Structure WorkdayDef
        Public isWorkday As Boolean
        Public Datum As Date
        Public Country As String
        Public Feiertag As String
    End Structure
    Public Structure WindowLocation
        Public fLocation As Point
        Public fWidth As Integer
        Public fHeight As Integer
    End Structure
    Public Structure ReadSettingReturn
        Public Value As String
        Public Value_1 As String
        Public Value_2 As String
    End Structure
    Public Structure RegFindResult
        Public Hive As RegistryHive
        Public RegPath As String
        Public Key As String
        Public Name As String
        Public VarType As VarTypeEnu
        Public Value As Object
    End Structure
    Public Structure TimeAddReturn
        Public Hours As String
        Public Minutes As Integer
    End Structure
    Public Structure SMTPCredentials
        Public SMTP_Password As String
        Public SMTP_User As String
        Public SMTP_Server As String
        Public SMTP_NoSSL As Boolean
        Public SMTP_SSL As Boolean
        Public SMTP_STARTTLS As Boolean
        Public SMTP_SenderAddress As String
        Public SMTP_SenderName As String
        Public SMTP_RecipientAddress As String
        Public SMTP_RecipientName As String
        Public SMTP_CC As String
        Public SMTP_BCC As String
        Public SMTP_NoSSL_Port As Integer
        Public SMTP_SSL_Port As Integer
        Public SMTP_STARTTLS_Port As Integer
    End Structure

#End Region

#Region "Properties"
    Private Shared _registryHiveValue As RegistryHive
    Public Shared Property RegistryHiveValue() As RegistryHive
        Get
            Return _registryHiveValue
        End Get
        Set(ByVal value As RegistryHive)
            _registryHiveValue = value
        End Set
    End Property
    Private Shared _RegistryPath As String
    Public Shared Property RegistryPath() As String
        Get
            Return _RegistryPath
        End Get
        Set(ByVal value As String)
            _RegistryPath = value
        End Set
    End Property

    Private Shared _alwaysOnTop As Boolean = False
    Public Shared Property AlwaysOnTop() As Boolean
        Get
            Return _alwaysOnTop
        End Get
        Set(ByVal value As Boolean)
            _alwaysOnTop = value
        End Set
    End Property

    'Private Shared _registryHiveValue As RegistryHive

    'Public Shared Property RegistryHiveValue() As RegistryHive
    '    Get
    '        Return _registryHiveValue
    '    End Get
    '    Set(ByVal value As RegistryHive)
    '        _registryHiveValue = value
    '    End Set
    'End Property

    'Private Shared _RegPath As String

    'Public Shared Property RegistryPath() As String
    '    Get
    '        Return _registryPath
    '    End Get
    '    Set(ByVal value As String)
    '        _registryPath = value
    '    End Set
    'End Property

#End Region

#Region "AdvTree"
    Public Shared Function FindNodeByTag(ByRef AdvTree As AdvTree, ByVal Tag As String) As Node
        Dim retnode As New Node

        Try
            If AdvTree.Nodes.Count > 0 Then
                For Each node As Node In AdvTree.Nodes
                    Application.DoEvents()
                    Dim StrVal As String = TryCast(node.Tag, String)
                    If String.IsNullOrEmpty(StrVal) Then StrVal = ""
                    If StrVal = Tag Then
                        Return node
                    End If
                    retnode = SearchRecursiveTag(AdvTree, node, Tag)
                    If retnode IsNot Nothing Then
                        Return retnode
                    End If
                Next
            Else
                retnode = Nothing
                Return retnode
            End If
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retnode
    End Function
    Private Shared Function SearchRecursiveTag(ByRef AdvTree As AdvTree, ByVal node As Node, ByVal Tag As String) As Node
        Dim retnode As Node = Nothing
        Try
            For Each aNode As Node In node.Nodes
                Application.DoEvents()
                Dim StrVal As String = TryCast(aNode.Tag, String)
                If String.IsNullOrEmpty(StrVal) Then StrVal = ""
                If StrVal = Tag Then
                    Return aNode
                End If
                retnode = SearchRecursiveTag(AdvTree, aNode, Tag)
                If retnode IsNot Nothing Then
                    Return retnode
                End If
            Next
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retnode
    End Function
    Public Shared Function FindNodeByName(ByRef AdvTree As AdvTree, ByVal NameVal As String) As Node
        Dim FoundNode As Node = Nothing
        Dim retnode As Node = Nothing
        Try
            For Each node As Node In AdvTree.Nodes
                Application.DoEvents()
                If node.Text = NameVal Then
                    Return node
                End If
                retnode = SearchRecursiveName(AdvTree, node, NameVal)
                If retnode IsNot Nothing Then
                    Return retnode
                End If
            Next
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retnode
    End Function
    Private Shared Function SearchRecursiveName(ByRef AdvTree As AdvTree, ByVal node As Node, ByVal NameVal As String) As Node
        Dim retnode As Node = Nothing
        Try
            For Each aNode As Node In node.Nodes
                Application.DoEvents()
                If aNode.Text = NameVal Then
                    Return aNode
                End If
                retnode = SearchRecursiveName(AdvTree, aNode, NameVal)
                If retnode IsNot Nothing Then
                    Return retnode
                End If
            Next
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retnode
    End Function

    Public Shared Sub DisplayTreeView(ByVal TreeView As AdvTree, ByVal root As JToken, ByVal rootName As String)
        TreeView.BeginUpdate()

        Try
            TreeView.Nodes.Clear()
            Dim tNode As Node = CType(TreeView.Nodes(TreeView.Nodes.Add(New Node(rootName))), Node)
            tNode.Tag = root
            tNode.Expanded = True
            tNode.Selectable = False
            AddNode(root, tNode, True)
        Finally
            TreeView.EndUpdate()
            TreeView.ExpandAll()
        End Try
    End Sub
    Private Shared Sub AddNode(ByVal token As JToken, ByVal inNode As Node, ByVal NextSelectable As Boolean)
        If token Is Nothing Then Return
        If TypeOf token Is JValue Then
            Dim childNode As Node = CType(inNode.Nodes(inNode.Nodes.Add(New Node(token.ToString()))), Node)
            childNode.Tag = token
            childNode.Expanded = True
            childNode.Selectable = NextSelectable
        ElseIf TypeOf token Is JObject Then
            Dim obj As JObject = CType(token, JObject)
            For Each propertyValue As JProperty In obj.Properties()
                Dim childNode As Node = CType(inNode.Nodes(inNode.Nodes.Add(New Node(propertyValue.Name))), Node)
                childNode.Tag = propertyValue
                childNode.Expanded = True
                If childNode.Text = "resources" Then
                    NextSelectable = False
                    childNode.Expanded = True
                Else
                    NextSelectable = True
                    childNode.Expanded = True
                End If
                childNode.Selectable = NextSelectable
                AddNode(propertyValue.Value, childNode, NextSelectable)
            Next
        ElseIf TypeOf token Is JArray Then
            Dim array As JArray = CType(token, JArray)
            For i As Integer = 0 To array.Count - 1
                Dim childNode As Node = CType(inNode.Nodes(inNode.Nodes.Add(New Node(inNode.Text & " " & i.ToString()))), Node)
                childNode.Tag = array(i)
                childNode.Expanded = True
                If childNode.Text = "resources" Then
                    NextSelectable = False
                    childNode.Expanded = True
                Else
                    NextSelectable = True
                    childNode.Expanded = True
                End If
                childNode.Selectable = NextSelectable
                AddNode(array(i), childNode, NextSelectable)
            Next
        Else
            Debug.WriteLine(String.Format("{0} not implemented", token.Type))
        End If
    End Sub



#End Region
#Region "Berechnungen"
    Public Shared Function BytesToKilobytes(ByVal intBytes As Integer) As Integer
        Return intBytes \ 1024
    End Function
    Public Shared Function BytesToMegabytes(ByVal intBytes As Integer) As Integer
        Return intBytes \ 1024 \ 1024
    End Function
    Public Shared Function FormatBytes(ByVal BytesCaller As ULong) As String
        Dim DoubleBytes As Double
        Try
            Select Case BytesCaller
                Case Is >= 1099511627776
                    DoubleBytes = Helper_VarConvert.ConvertToDouble(BytesCaller / 1099511627776) 'TB
                    Return FormatNumber(DoubleBytes, 0) & " TB"
                Case 1073741824 To 1099511627775
                    DoubleBytes = Helper_VarConvert.ConvertToDouble(BytesCaller / 1073741824) 'GB
                    Return FormatNumber(DoubleBytes, 0) & " GB"
                Case 1048576 To 1073741823
                    DoubleBytes = Helper_VarConvert.ConvertToDouble(BytesCaller / 1048576) 'MB
                    Return FormatNumber(DoubleBytes, 0) & " MB"
                Case 1024 To 1048575
                    DoubleBytes = Helper_VarConvert.ConvertToDouble(BytesCaller / 1024) 'KB
                    Return FormatNumber(DoubleBytes, 0) & " KB"
                Case 0 To 1023
                    DoubleBytes = BytesCaller ' bytes
                    Return FormatNumber(DoubleBytes, 0) & " bytes"
                Case Else
                    Return ""
            End Select
        Catch
            Return ""
        End Try
    End Function
    Public Shared Function IsInt(ByVal value As Object) As Boolean
        Dim i As Integer
        Return Integer.TryParse(System.Convert.ToString(value), i)
    End Function
    Public Shared Function KonvertZahl(zahlenstring As String, Optional ReturnAsString As Boolean = False) As String
        Dim ReturnStr As String
        Dim po_p As Integer = 0
        Dim po_k As Integer = 0
        Dim prdummy As String = ""
        po_p = zahlenstring.IndexOf(".")
        po_k = zahlenstring.IndexOf(".")
        If po_p > 0 AndAlso po_k > 0 Then
            If po_p < po_k Then
                prdummy = Replace(zahlenstring, ".", "$")
                prdummy = Replace(zahlenstring, ",", ".")
                prdummy = Replace(zahlenstring, "$", "")
            Else
                prdummy = Replace(zahlenstring, ",", "")
            End If
        Else
            prdummy = Replace(zahlenstring, ",", ".")
        End If
        If prdummy Is Nothing OrElse prdummy = "" Then
            prdummy = zahlenstring
        End If
        If ReturnAsString Then
            ReturnStr = prdummy.ToString
        Else
            ReturnStr = CStr(prdummy)
        End If
        Return ReturnStr
    End Function
#End Region
#Region "Color"
    Public Shared Function ARGBColorToHex(color As Color) As String
        Return "#" & color.A.ToString("X2") & color.R.ToString("X2") & color.G.ToString("X2") & color.B.ToString("X2")
    End Function
    Public Shared Function RGBColorToHex(color As Color) As String
        Return "#" & color.R.ToString("X2") & color.G.ToString("X2") & color.B.ToString("X2")
    End Function
    Public Shared Function GetComplementaryColor(ByVal colorvalue As Color) As Color
        Dim farbeinv As String = ColorTranslator.ToHtml(colorvalue).Replace("#", "")
        Dim hexVal As Integer = System.Convert.ToInt32(farbeinv, 16) Xor &HFFFFFF
        farbeinv = "#" & System.Convert.ToString(hexVal, 16).PadLeft(6, "0"c)
        Dim colorreturn As Color = ColorTranslator.FromHtml(farbeinv)
        Return colorreturn
    End Function
    Public Shared Function GetComplementaryColor(ByVal colorvalue As String) As Color
        Dim farbeinv As String = colorvalue.Replace("#", "")
        Dim hexVal As Integer = System.Convert.ToInt32(farbeinv, 16) Xor &HFFFFFF
        farbeinv = "#" & System.Convert.ToString(hexVal, 16).PadLeft(6, "0"c)
        Dim colorreturn As Color = ColorTranslator.FromHtml(farbeinv)
        Return colorreturn
    End Function
#End Region
#Region "Controls"

    Public Structure ModuleDef
        Public ModulName As String
        Public Component As String
        Public LicKey As String
        Public ValidTo As Date
        Public LicQuantity As Integer
        Public Control1 As Object
        Public Control2 As Object
        Public Control3 As Object
    End Structure
    Public Shared Sub ActivateControlsByLicense(ByVal ModulesList As List(Of ModuleDef))
        'Dim cCrypt As cCrypt = New cCrypt
        Dim ModuleEntry1 As String = ""

        Try
            For Each MnuModule As ModuleDef In ModulesList
                Call SetControlVisible(MnuModule.Control1, False)
                Call SetControlVisible(MnuModule.Control2, False)
            Next

            For Each sp As SettingsProperty In My.Settings.Properties
                'Debug.Print(sp.Name & ": " & sp.DefaultValue.ToString)
                If sp.Name.ToUpper.Contains("ALLOW") Then
                    For Each MnuModule As ModuleDef In ModulesList
                        Helper_cCrypt.Encrypt(256, MnuModule.ModulName)
                        ModuleEntry1 = Helper_cCrypt.EncryptedString
                        If sp.DefaultValue.ToString = ModuleEntry1 Then
                            Call SetControlVisible(MnuModule.Control1, True)
                            Call SetControlVisible(MnuModule.Control2, True)
                        End If
                    Next
                End If
            Next

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If IsIDE() Then Stop
        End Try
    End Sub
    Public Shared Sub SetControlVisible(ByVal Control As Object, ByVal visible As Boolean)

        Dim contr0 As ButtonX = TryCast(Control, ButtonX)
        If contr0 IsNot Nothing Then
            contr0.Visible = visible
        End If
        Dim contr1 As ButtonItem = TryCast(Control, ButtonItem)
        If contr1 IsNot Nothing Then
            contr1.Visible = visible
        End If
        Dim contr2 As TableLayoutPanel = TryCast(Control, TableLayoutPanel)
        If contr2 IsNot Nothing Then
            contr2.Visible = visible
        End If
        Dim contr3 As Panel = TryCast(Control, Panel)
        If contr3 IsNot Nothing Then
            contr3.Visible = visible
        End If
    End Sub
    Public Shared Function FindControlStartsWith(root As Control, name As String, recursive As Boolean, comparison As StringComparison) As Control()
        If root Is Nothing Then
            Throw New ArgumentNullException("root")
        End If
        Dim controls As New List(Of Control)
        Dim stack As New Stack(Of Control)()
        stack.Push(root)

        While stack.Count > 0
            Dim c As Control = stack.Pop()
            If c.Name.StartsWith(name, comparison) Then
                controls.Add(c)
            End If
            If recursive Then
                For Each child As Control In root.Controls
                    stack.Push(child)
                Next
            End If
        End While
        Return controls.ToArray()
    End Function
    Public Shared Function FindControlStartsWith(root As SuperTabControlPanel, name As String, recursive As Boolean, comparison As StringComparison) As Control()
        If root Is Nothing Then
            Throw New ArgumentNullException("root")
        End If
        Dim controls As New List(Of Control)
        Dim stack As New Stack(Of Control)()
        stack.Push(root)

        While stack.Count > 0
            Dim c As Control = stack.Pop()
            If c.Name.StartsWith(name, comparison) Then
                controls.Add(c)
            End If
            If recursive Then
                For Each child As Control In root.Controls
                    stack.Push(child)
                Next
            End If
        End While
        Return controls.ToArray()
    End Function
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, Optional ByVal ctrlType As System.Type = Nothing) As List(Of Control)
        If parent Is Nothing Then Return list
        If ctrlType Is Nothing Then
            list.Add(parent)
        Else
            If parent.GetType Is ctrlType Then
                list.Add(parent)
            End If
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type, ByVal StartsWithPattern As String) As List(Of Control)
        If parent Is Nothing Then Return list
        If parent.GetType Is ctrlType Then
            If parent.Name.StartsWith(StartsWithPattern) Then
                list.Add(parent)
            End If
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType, StartsWithPattern)
        Next
        Return list
    End Function

#End Region
#Region "Dateien & Pfade"
    Public Shared Sub CollectFilesAndFolders(ByVal rootDirs As IEnumerable(Of DirectoryInfo),
                                         ByVal folderCollector As ICollection(Of DirectoryInfo),
                                         ByVal fileCollector As ICollection(Of FileInfo),
                                         Optional ByVal pattern As String = "*.*",
                                         Optional ByVal DaysBack As Integer = -10)
        'rekursive anonyme Methode
        Dim recurse As Action(Of IEnumerable(Of DirectoryInfo)) = Sub(dirs As IEnumerable(Of DirectoryInfo))
                                                                      For Each dirinf As DirectoryInfo In dirs
                                                                          Dim files As IEnumerable(Of FileInfo)
                                                                          Try
                                                                              files = dirinf.EnumerateFiles(pattern)
                                                                          Catch ex As UnauthorizedAccessException
                                                                              'für manche Directories hat das Prog keine Rechte
                                                                              Continue For
                                                                          End Try
                                                                          For Each fileInf As FileInfo In files
                                                                              Dim DayBefore As Date = CDate(Format(Now.AddDays(Math.Abs(DaysBack) * -1), "yyyy-MM-dd") & " 00:00:00")
                                                                              If CDate(fileInf.LastWriteTime) >= DayBefore Then
                                                                                  fileCollector.Add(fileInf)
                                                                              End If
                                                                          Next
                                                                          folderCollector.Add(dirinf)
                                                                          'selbst-aufruf
                                                                          recurse(dirinf.EnumerateDirectories)
                                                                      Next
                                                                  End Sub
        'anonyme Methode aufrufen
        recurse(rootDirs)
    End Sub
    Public Shared Function GetAppPath() As String
        Dim path As String = Assembly.GetExecutingAssembly().Location
        path = path.Substring(0, path.LastIndexOf(" \ ") + 1)
        Return path
    End Function
    Public Shared Function GetFileType(extension As String) As String
        Dim contentType As String
        Using rgk As RegistryKey = Registry.ClassesRoot.OpenSubKey(System.Convert.ToString("\") & extension)
            contentType = rgk.GetValue("", String.Empty).ToString()
        End Using
        Return contentType
    End Function
    Public Shared Function IsFileInUse(ByVal fullFilePath As String) As Boolean
        Dim ff As Integer = FileSystem.FreeFile()
        If File.Exists(fullFilePath) Then
            Try
                FileOpen(ff, fullFilePath, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
            Catch
                Return True
            Finally
                FileClose(ff)
            End Try
        End If
        Return False
    End Function
    Public Shared Function FileInUse(ByVal sFile As String) As Boolean
        Dim thisFileInUse As Boolean = False
        If File.Exists(sFile) Then
            Try
                Using f As New FileStream(sFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                    ' thisFileInUse = False
                End Using
            Catch
                thisFileInUse = True
            End Try
        End If
        Return thisFileInUse
    End Function
    Public Shared Function IsIDE(ByVal Optional AlwaysIsIde As Boolean = False) As Boolean
        If Debugger.IsAttached AndAlso Not AlwaysIsIde Then
            Return True
        ElseIf Not Debugger.IsAttached AndAlso Not AlwaysIsIde Then
            Return False
        ElseIf Debugger.IsAttached AndAlso AlwaysIsIde Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Function FileLineCount(ByVal Filename As String) As Integer
        Dim Lines As Integer = 0

        Using oStream As New IO.FileStream(Filename, IO.FileMode.Open, IO.FileAccess.Read)
            ' Blockgröße
            Dim BlockSize As Integer = 16384

            Dim Bytes(BlockSize) As Byte
            Dim ReadSize As Integer
            Dim StartIndex As Integer

            ' Datei blockweise auslesen, bis Dateiende erreicht
            Do
                ReadSize = oStream.Read(Bytes, 0, BlockSize)

                ' Anzahl Zeilenumbruchzeichen ermitteln
                StartIndex = -1
                Do
                    StartIndex = Array.IndexOf(Bytes, CByte(13), StartIndex + 1, ReadSize - StartIndex)
                    If StartIndex >= 0 Then Lines += 1
                Loop Until StartIndex < 0

            Loop Until ReadSize = 0
            oStream.Close()
        End Using

        Return Lines
    End Function
    Public Shared Function PathExists(Pfad As String, Optional ReturnBool As Boolean = False) As Object
        Try
            Pfad = Pfad.Trim
            If Pfad <> "" AndAlso Path.IsPathRooted(Pfad) AndAlso Not Directory.Exists(Pfad) Then
                Directory.CreateDirectory(Pfad)
            End If
            If Right(Pfad, 1) <> "\" Then
                Pfad = Pfad & "\"
            End If
            If ReturnBool Then
                Return True
            Else
                Return CStr(Pfad)
            End If
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False)
            If Helper.IsIDE() Then Stop
            If ReturnBool Then
                Return False
            Else
                Return ""
            End If
        End Try
    End Function
    Public Shared Function ValPath(Path As String, Optional ByVal CreatePath As Boolean = True) As String
        Path = Path.Trim
        If Right(Path, 1) <> "\" Then
            Path = Path & "\"
        End If
        If CreatePath AndAlso Not Directory.Exists(Path) Then
            Directory.CreateDirectory(Path)
        End If
        Return (Path)
    End Function
    Public Shared Function PreventDoubleFiles(ByVal fullPath As String) As String
        Dim count As Integer = 1
        Dim fileNameOnly As String = Path.GetFileNameWithoutExtension(fullPath)
        Dim extension As String = Path.GetExtension(fullPath)
        Dim filepath As String = Path.GetDirectoryName(fullPath)
        Dim newFullPath As String = fullPath

        While File.Exists(newFullPath)
            Dim tempFileName As String = String.Format("{0}({1})", fileNameOnly, Math.Min(Interlocked.Increment(count), count - 1))
            newFullPath = Path.Combine(filepath, tempFileName & extension)
        End While
        Return newFullPath
    End Function
    Public Shared Sub WriteJsonToFile(ByVal jsonString As String,
                                  ByVal FilePath As String,
                                  ByVal FileName As String,
                                  Optional ByVal SelectEncoding As Encoding = Nothing)
        Dim JsonFormatted As String = ""
        Try
            If IsNothing(SelectEncoding) Then
                SelectEncoding = Encoding.UTF8
            End If
            JsonFormatted = Newtonsoft.Json.JsonConvert.SerializeObject(jsonString, Newtonsoft.Json.Formatting.Indented)
            File.WriteAllText(Path.Combine(Helper.ValPath(FilePath), FileName), JsonFormatted, SelectEncoding)
            'Debug.Print(JsonFormatted)

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
    End Sub

    Public Shared Function GetTempPath(Optional ByVal TempSource As String = "windows") As String
        Dim ProgTempPath As String = ValPath(Application.StartupPath & "\temp", True)
        Dim WindowsTempPath As String = ValPath(Path.GetTempPath(), True)
        WindowsTempPath = ValPath(WindowsTempPath & "ProStruktura", True)
        If TempSource = "windows" Then
            Return WindowsTempPath
        Else
            Return ProgTempPath
        End If
    End Function
    Public Shared Sub MoveFilesOlderThan(ByVal RootPath As String, ByVal OldPath As String, ByVal DaysKeepNew As Integer)
        OldPath = ValPath(OldPath, True)
        Dim FSearch As New Helper_FileSearch
        Dim Filter As New Helper_FileSearch.SearchFilter
        Dim FileArray As New List(Of String)
        With Filter
            .Listing = Helper_FileSearch.SearchFilter.LO.FILES_ONLY
            .FileTypes = "xml"
            .MaxDate = DateAdd(DateInterval.Day, DaysKeepNew * -1, Now)
            .NoSubFolders = True
            .DateType = Helper_FileSearch.SearchFilter.DT.LAST_WRITE_TIME
        End With
        FSearch.Search(RootPath, "List", Filter, FileArray)
        For Each filename As String In FileArray
            Application.DoEvents()
            Dim fname As String = Path.GetFileName(filename)
            File.Move(filename, OldPath & fname)
        Next
    End Sub
    Public Shared Function DeleteFilesOlderThanDays(ByVal Path As String, ByVal FileAge As Integer, Optional ByVal Pattern As String = " * .* ") As Integer
        Try
            Dim FileCount As Integer = 0
            If Directory.Exists(Path) AndAlso FileAge > 0 Then
                For Each file As FileInfo In New DirectoryInfo(Path).GetFiles(Pattern)
                    If (Now - file.CreationTime).Days > FileAge Then
                        file.Delete()
                        FileCount = FileCount + 1
                    End If
                Next
            End If
            If FileCount > 0 Then
                'Call Helper_Logger.writelog(Level.Info, "Es wurden " & FileCount & " alte Dateien mit dem Pattern " & Pattern & " gelöscht.")
            End If
            Return FileCount
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False)
            If Helper.IsIDE() Then Stop
            Return 0
        End Try
    End Function
    Public Shared Sub DeleteFilesOlderThan(ByVal RootPath As String, ByVal DaysKeepOld As Integer)
        Dim FSearch As New Helper_FileSearch
        Dim Filter As New Helper_FileSearch.SearchFilter
        Dim FileArray As New List(Of String)
        With Filter
            .Listing = Helper_FileSearch.SearchFilter.LO.FILES_ONLY
            .FileTypes = "xml"
            .MaxDate = DateAdd(DateInterval.Day, DaysKeepOld * -1, Now)
            .NoSubFolders = True
            .DateType = Helper_FileSearch.SearchFilter.DT.LAST_WRITE_TIME
        End With
        FSearch.Search(RootPath, "List", Filter, FileArray)
        For Each filename As String In FileArray
            Application.DoEvents()
            Dim fname As String = Path.GetFileName(filename)
            File.Delete(filename)
        Next
    End Sub
    Public Shared Sub CleanUpBackupArchiv(ByVal SavePathNew As String,
                                      ByVal SavePathOld As String,
                                      ByVal DaysKeepNew As Integer,
                                      ByVal DaysKeepOld As Integer,
                                      ByVal DoMove As Boolean,
                                      ByVal DoDelete As Boolean)
        Try
            If DoMove AndAlso DaysKeepNew > 0 Then
                Call MoveFilesOlderThan(SavePathNew, SavePathOld, DaysKeepNew)
            End If
            If DoDelete AndAlso DaysKeepOld > 0 Then
                Call DeleteFilesOlderThan(SavePathOld, DaysKeepOld)
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
    End Sub
    Public Shared Function CleanPathname(ByVal Pathname As String) As String
        Dim NewPathname As String = ""
        Try
            If String.IsNullOrWhiteSpace(Pathname) Then
                Pathname = ""
            End If
            NewPathname = Regex.Replace(Pathname, "<|>|\?|""|:|\||\\|\/|\*|&", "-", RegexOptions.IgnoreCase)
            NewPathname = Regex.Replace(NewPathname, " CON | PRN | AUX | NUL | COM1 | COM2 | COM3 | COM4 | COM5 | COM6 | COM7 | COM8 | COM9 | LPT1 | LPT2 | LPT3 | LPT4 | LPT5 | LPT6 | LPT7 | LPT8 | LPT9 ", "-", RegexOptions.IgnoreCase)

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
            Return NewPathname
        End Try
        Return NewPathname
    End Function
    Public Shared Sub CheckPath(ctrl As TextBox)
        Dim dn As String = ""
        If ctrl.Text.Trim <> "" Then
            dn = Path.GetDirectoryName(ctrl.Text)
            If Directory.Exists(ctrl.Text) Then
                ctrl.BackColor = Color.LightGreen
            Else
                ctrl.BackColor = Color.Red
            End If
        End If
    End Sub
    Public Shared Function IsDriveReady(ByVal sDrive As String) As Boolean
        ' Prüft, ob das angegebene Laufwerk existiert und ob darauf zugegriffen werden kann
        Try
            Dim oDrive As New DriveInfo(sDrive)
            Return oDrive.IsReady
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Shared Function FormatXML(ByVal xml As String) As String
        Dim result As String = ""
        Dim formattedXml As String = ""
        Dim document As XmlDocument = New XmlDocument()

        Using mStream As MemoryStream = New MemoryStream()
            Using writer As XmlTextWriter = New XmlTextWriter(mStream, Encoding.Unicode)

                Try
                    document.LoadXml(xml)
                    writer.Formatting = System.Xml.Formatting.Indented
                    document.WriteContentTo(writer)
                    writer.Flush()
                    mStream.Flush()
                    mStream.Position = 0
                    Dim sReader As StreamReader = New StreamReader(mStream)
                    formattedXml = sReader.ReadToEnd()
                    result = formattedXml
                Catch unusedXmlException As XmlException
                    Call Helper_ErrorHandling.HandleErrorCatch(unusedXmlException, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
                    If Helper.IsIDE() Then Stop
                Catch ex As Exception
                    Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
                    If Helper.IsIDE() Then Stop
                End Try

            End Using
        End Using
        Return result
    End Function

    Public Shared Sub WriteStringToFile(ByVal Text2Write As List(Of String),
                                    ByVal PathName As String,
                                    ByVal FileName As String)
        Try
            Dim TextOut As String = ""
            TextOut = String.Join(Environment.NewLine, Text2Write)
            Call WriteStringToFile(TextOut, PathName, FileName)

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
    End Sub
    Public Shared Sub WriteStringToFile(ByVal Text2Write As String,
                                    ByVal PathName As String,
                                    ByVal FileName As String,
                                    ByVal Optional Overwrite As Boolean = True)
        Try
            FileName = Path.Combine(ValPath(PathName, True), FileName)
            If File.Exists(FileName) AndAlso Overwrite Then
                File.Delete(FileName)
            End If
            File.WriteAllText(FileName, Text2Write)

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
    End Sub

#End Region
#Region "Datum & Zeit"
    Public Shared Function GetLastUpdateDate() As DateTime
        Dim LastUpdateDate As New DateTime
        Dim LastUpdateDateLocal As New DateTime

        Try
            Dim VersionRevision As String = My.Application.Info.Version.Revision.ToString
            Dim pad As Char = "0"c
            VersionRevision = VersionRevision.PadLeft(4, pad)

            Dim TimeHour As Integer = Helper_VarConvert.ConvertToInteger(VersionRevision.Substring(1, 2), 0)
            Dim TimeMinute As Integer = Helper_VarConvert.ConvertToInteger(VersionRevision.Substring(3, 2), 0)
            Dim AppInfoVersionMinor As Integer = My.Application.Info.Version.Minor
            Dim AppInfoVersionBuild As Integer = My.Application.Info.Version.Build
            LastUpdateDate = New DateTime(AppInfoVersionMinor, 1, 1, TimeHour, TimeMinute, 0).AddDays(AppInfoVersionBuild - 1)
            LastUpdateDateLocal = LastUpdateDate.ToLocalTime()

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return LastUpdateDateLocal

    End Function
    Public Shared Function AddWorkingDays(ByVal DateIn As DateTime, ByVal ShiftDate As Integer) As DateTime
        Dim datDate As DateTime = DateIn.AddDays(ShiftDate)
        Dim Feiertag As New List(Of Helper_FeierTage.TypeFeiertag)
        Dim IsFeiertag As Boolean = False
        Feiertag = GetAlleFeiertage()
        While Weekday(datDate) = 1 OrElse Weekday(datDate) = 7 OrElse IsFeiertag
            datDate = datDate.AddDays(Helper_VarConvert.ConvertToDouble(IIf(ShiftDate < 0, -1, 1)))
            Dim Response As New Helper_FeierTage.TypeFeiertag
            Response = Feiertag.Find(Function(p) p.DatumA.ToString("yyyy-MM-dd") = datDate.ToString("yyyy-MM-dd"))
            If Not Response.NameA = Nothing Then
                IsFeiertag = True
            Else
                IsFeiertag = False
            End If
        End While
        Return datDate
    End Function
    Public Shared Function DateRange(Start As DateTime, Thru As DateTime) As IEnumerable(Of Date)
        Return Enumerable.Range(0, (Thru.Date - Start.Date).Days + 1).Select(Function(i) Start.AddDays(i))
    End Function
    Public Shared Function Day_between_days(ByVal StartDate As Date, ByVal EndDate As Date, ByVal CompareDate As Date, ByVal StartEnd_involved As Boolean) As Boolean
        Dim Temp1 As Integer
        Dim Temp2 As Integer
        Temp1 = Date.Compare(StartDate, CompareDate)
        Temp2 = Date.Compare(EndDate, CompareDate)
        If Date.Compare(StartDate, EndDate) < 0 Then
            If Temp2 = 0 AndAlso StartEnd_involved OrElse Temp1 = 0 AndAlso StartEnd_involved Then
                Return True
            End If
            If Temp1 < 0 AndAlso Temp2 > 0 Then
                Return True
            Else
                Return False
            End If
        Else
            'MsgBox("The start date is after the end date", MsgBoxStyle.Critical, "Error !")
            Return False
        End If
    End Function
    Public Shared Function FormatElapsedTime(elapsedTime As TimeSpan, Optional ShowAll As Boolean = True) As String
        Dim StrElapsedTime As String = ""
        Dim StrElapsedTimeA(3) As String
        If elapsedTime.Hours = 0 AndAlso Not ShowAll Then
            StrElapsedTimeA(0) = ""
        Else
            StrElapsedTimeA(0) = Format(elapsedTime.Hours, "00") & " Std."
        End If
        If elapsedTime.Minutes = 0 AndAlso Not ShowAll Then
            StrElapsedTimeA(1) = ""
        Else
            StrElapsedTimeA(1) = Format(elapsedTime.Minutes, "00") & " Min."
        End If
        If elapsedTime.Seconds = 0 AndAlso Not ShowAll Then
            StrElapsedTimeA(2) = ""
        Else
            StrElapsedTimeA(2) = Format(elapsedTime.Seconds, "00") & " Sek."
        End If
        If elapsedTime.Milliseconds = 0 AndAlso Not ShowAll Then
            StrElapsedTimeA(3) = ""
        Else
            StrElapsedTimeA(3) = Format(elapsedTime.Milliseconds, "000") & " MSek."
        End If
        StrElapsedTime = Join(StrElapsedTimeA, " ")
        Return StrElapsedTime
    End Function
    Public Shared Function GetAlleFeiertageAsDate(Optional ByVal OnlyDayOff As Boolean = False) As List(Of Date)
        Dim dl As New List(Of Date)
        Dim df As New List(Of Helper_FeierTage.TypeFeiertag)
        df = GetAlleFeiertage()
        If OnlyDayOff Then
            Dim Result As New List(Of Helper_FeierTage.TypeFeiertag)
            Result = df.FindAll(Function(p1)
                                    Return p1.arbeitsfrei
                                End Function)
            df = Result
        End If
        df.Sort(Function(x, y) x.DatumA.CompareTo(y.DatumA))

        For Each da As Helper_FeierTage.TypeFeiertag In df
            dl.Add(da.DatumA)
        Next
        dl.Sort()
        Return dl
    End Function
    Public Shared Function GetAlleFeiertage() As List(Of Helper_FeierTage.TypeFeiertag)
        'Debug.Print(LogManager.GetRepository().Threshold.DisplayName)

        Dim Feiertag As New List(Of Helper_FeierTage.TypeFeiertag)
        Dim FeiertagA(65) As Helper_FeierTage.TypeFeiertag
        Dim a, b, c, d, e, s As Long
        Dim M, N As Long
        Dim OsternA As Date
        Dim JahrA As Integer = Helper_VarConvert.ConvertToInteger(Now.ToString("yyyy"), 0)
        Try
            FeiertagA(0).NameA = "4. Advent"
            FeiertagA(0).DatumA = GetViertenAdvent(JahrA)
            FeiertagA(0).arbeitsfrei = False
            FeiertagA(1).NameA = "3. Advent"
            FeiertagA(1).DatumA = DateAdd("d", -7, GetViertenAdvent(JahrA))
            FeiertagA(1).arbeitsfrei = False
            FeiertagA(2).NameA = "2. Advent"
            FeiertagA(2).DatumA = DateAdd("d", -14, GetViertenAdvent(JahrA))
            FeiertagA(2).arbeitsfrei = False
            FeiertagA(3).NameA = "1. Advent"
            FeiertagA(3).DatumA = DateAdd("d", -21, GetViertenAdvent(JahrA))
            FeiertagA(3).arbeitsfrei = False
            FeiertagA(4).NameA = "Totensonntag"
            FeiertagA(4).DatumA = DateAdd("d", -28, GetViertenAdvent(JahrA))
            FeiertagA(4).arbeitsfrei = False
            FeiertagA(5).NameA = "Volkstrauertag"
            FeiertagA(5).DatumA = DateAdd("d", -35, GetViertenAdvent(JahrA))
            FeiertagA(5).arbeitsfrei = False
            FeiertagA(6).NameA = "Buß- und Bettag"
            FeiertagA(6).DatumA = DateAdd("d", -11, FeiertagA(3).DatumA)
            FeiertagA(6).arbeitsfrei = False
            FeiertagA(7).NameA = "Muttertag"
            FeiertagA(7).DatumA = GetMuttertag(JahrA)
            FeiertagA(7).arbeitsfrei = False
            FeiertagA(8).NameA = "Erntedankfest"
            FeiertagA(8).DatumA = GetErntedankfest(JahrA)
            FeiertagA(8).arbeitsfrei = False
            FeiertagA(9).NameA = "Neujahr"
            FeiertagA(9).DatumA = CDate("01.01." & JahrA)
            FeiertagA(9).arbeitsfrei = True
            FeiertagA(10).NameA = "Heiligen drei Könige"
            FeiertagA(10).DatumA = CDate("06.01." & JahrA)
            FeiertagA(10).arbeitsfrei = False
            FeiertagA(11).NameA = "Maifeiertag"
            FeiertagA(11).DatumA = CDate("01.05." & JahrA)
            FeiertagA(11).arbeitsfrei = True
            FeiertagA(12).NameA = "Tag der Deutschen Einheit"
            FeiertagA(12).DatumA = CDate("03.10." & JahrA)
            FeiertagA(12).arbeitsfrei = True
            FeiertagA(13).NameA = "Heiligabend"
            FeiertagA(13).DatumA = CDate("24.12." & JahrA)
            FeiertagA(13).arbeitsfrei = False
            FeiertagA(14).NameA = "1. Weihnachtstag"
            FeiertagA(14).DatumA = CDate("25.12." & JahrA)
            FeiertagA(14).arbeitsfrei = True
            FeiertagA(15).NameA = "2. Weihnachtstag"
            FeiertagA(15).DatumA = CDate("26.12." & JahrA)
            FeiertagA(15).arbeitsfrei = True
            FeiertagA(16).NameA = "Silvester"
            FeiertagA(16).DatumA = CDate("31.12." & JahrA)
            FeiertagA(16).arbeitsfrei = False
            FeiertagA(17).NameA = "Reformationstag"
            FeiertagA(17).DatumA = CDate("31.10." & JahrA)
            FeiertagA(17).arbeitsfrei = False
            FeiertagA(18).NameA = "Allerheiligen"
            FeiertagA(18).DatumA = CDate("01.11." & JahrA)
            FeiertagA(18).arbeitsfrei = False
            FeiertagA(19).NameA = "Maria Himmelfahrt"
            FeiertagA(19).DatumA = CDate("15.08." & JahrA)
            FeiertagA(19).arbeitsfrei = False
            FeiertagA(20).NameA = "Valentinstag"
            FeiertagA(20).DatumA = CDate("14.02." & JahrA)
            FeiertagA(20).arbeitsfrei = False
            FeiertagA(21).NameA = "Tag der Arbeit"
            FeiertagA(21).DatumA = CDate("01.05." & JahrA)
            FeiertagA(21).arbeitsfrei = True
            FeiertagA(22).NameA = "Friedensfest"
            FeiertagA(22).DatumA = CDate("08.08." & JahrA)
            FeiertagA(22).arbeitsfrei = False
            FeiertagA(23).NameA = "Nikolaus"
            FeiertagA(23).DatumA = CDate("06.12." & JahrA)
            FeiertagA(23).arbeitsfrei = False
            a = JahrA Mod 19
            b = JahrA Mod 4
            c = JahrA Mod 7
            M = CLng(((8 * Val(JahrA / 100) + 13) / 25) - 2)
            s = CLng(JahrA / 100) - CLng(JahrA / 400) - 2
            M = (15 + s - M) Mod 30
            N = (6 + s) Mod 7
            d = (M + 19 * a) Mod 30
            If d = 29 Then
                d = 28
            ElseIf d = 28 Then
                If (JahrA Mod 19) > 10 Then d = 27
            End If
            e = (2 * b + 4 * c + 6 * d + N) Mod 7
            OsternA = CDate(Format(DateAdd("d", (d + e + 1), JahrA & "-03-21"), "yyyy-MM-dd"))
            FeiertagA(24).NameA = "Ostersonntag"
            FeiertagA(24).DatumA = OsternA
            FeiertagA(24).arbeitsfrei = True
            FeiertagA(25).NameA = "Ostermontag"
            FeiertagA(25).DatumA = DateAdd("d", 1, OsternA)
            FeiertagA(25).arbeitsfrei = True
            FeiertagA(26).NameA = "Karfreitag"
            FeiertagA(26).DatumA = DateAdd("d", -2, OsternA)
            FeiertagA(26).arbeitsfrei = True
            FeiertagA(27).NameA = "Pfingstsonntag"
            FeiertagA(27).DatumA = DateAdd("d", 49, OsternA)
            FeiertagA(27).arbeitsfrei = True
            FeiertagA(28).NameA = "Pfingstmontag"
            FeiertagA(28).DatumA = DateAdd("d", 50, OsternA)
            FeiertagA(28).arbeitsfrei = True
            FeiertagA(29).NameA = "Christi Himmelfahrt"
            FeiertagA(29).DatumA = DateAdd("d", 39, OsternA)
            FeiertagA(29).arbeitsfrei = True
            FeiertagA(30).NameA = "Aschermittwoch"
            FeiertagA(30).DatumA = DateAdd("d", -46, OsternA)
            FeiertagA(30).arbeitsfrei = False
            FeiertagA(31).NameA = "Fronleichnam"
            FeiertagA(31).DatumA = DateAdd("d", 60, OsternA)
            FeiertagA(31).arbeitsfrei = True
            FeiertagA(32).NameA = "Herz-Jesu-Freitag"
            FeiertagA(32).DatumA = DateAdd("d", 68, OsternA)
            FeiertagA(32).arbeitsfrei = False
            FeiertagA(33).NameA = "Rosenmontag"
            FeiertagA(33).DatumA = DateAdd("d", -48, OsternA)
            FeiertagA(33).arbeitsfrei = False
            FeiertagA(34).NameA = "Eisheiligen: Mamertus"
            FeiertagA(34).DatumA = CDate("11.05." & JahrA)
            FeiertagA(34).arbeitsfrei = False
            FeiertagA(35).NameA = "Eisheiligen: Pankratius"
            FeiertagA(35).DatumA = CDate("12.05." & JahrA)
            FeiertagA(35).arbeitsfrei = False
            FeiertagA(36).NameA = "Eisheiligen: Servatius"
            FeiertagA(36).DatumA = CDate("13.05." & JahrA)
            FeiertagA(36).arbeitsfrei = False
            FeiertagA(37).NameA = "Eisheiligen: Bonifatius"
            FeiertagA(37).DatumA = CDate("14.05." & JahrA)
            FeiertagA(37).arbeitsfrei = False
            FeiertagA(38).NameA = "Eisheiligen: kalte Sophie"
            FeiertagA(38).DatumA = CDate("15.05." & JahrA)
            FeiertagA(38).arbeitsfrei = False
            FeiertagA(39).NameA = "Weiber Fastnacht"
            FeiertagA(39).DatumA = DateAdd("d", -52, OsternA)
            FeiertagA(39).arbeitsfrei = False
            FeiertagA(40).NameA = "Fastnacht"
            FeiertagA(40).DatumA = DateAdd("d", -47, OsternA)
            FeiertagA(40).arbeitsfrei = False
            FeiertagA(41).NameA = "Beginn der Sommerzeit"
            FeiertagA(41).DatumA = GetBeginnSommerzeit(JahrA)
            FeiertagA(41).arbeitsfrei = False
            FeiertagA(42).NameA = "Beginn der Winterzeit"
            FeiertagA(42).DatumA = GetBeginnWinterzeit(JahrA)
            FeiertagA(42).arbeitsfrei = False
            FeiertagA(43).NameA = "Winteranfang"
            FeiertagA(43).DatumA = CDate("21.12." & JahrA)
            FeiertagA(43).arbeitsfrei = False
            FeiertagA(44).NameA = "Frühlingsanfang"
            FeiertagA(44).DatumA = CDate("20.03." & JahrA)
            FeiertagA(44).arbeitsfrei = False
            FeiertagA(45).NameA = "Sommeranfang"
            FeiertagA(45).DatumA = CDate("21.06." & JahrA)
            FeiertagA(45).arbeitsfrei = False
            FeiertagA(46).NameA = "Herbstanfang"
            FeiertagA(46).DatumA = CDate("22.09." & JahrA)
            FeiertagA(46).arbeitsfrei = False
            FeiertagA(47).NameA = "Tag des Gedenkens an die Opfer des Nationalsozialismus"
            FeiertagA(47).DatumA = CDate("27.01." & JahrA)
            FeiertagA(47).arbeitsfrei = False
            FeiertagA(48).NameA = "Unbefleckte Empfängnis Mariens"
            FeiertagA(48).DatumA = CDate("08.12." & JahrA)
            FeiertagA(48).arbeitsfrei = False
            FeiertagA(49).NameA = "Darstellung des Herrn"
            FeiertagA(49).DatumA = CDate("02.02." & JahrA)
            FeiertagA(49).arbeitsfrei = False
            FeiertagA(50).NameA = "Joseftag"
            FeiertagA(50).DatumA = CDate("19.03." & JahrA)
            FeiertagA(50).arbeitsfrei = False
            FeiertagA(51).NameA = "Verkündigung des Herrn"
            FeiertagA(51).DatumA = CDate("25.03." & JahrA)
            FeiertagA(51).arbeitsfrei = False
            FeiertagA(52).NameA = "Geburt Johannes des Täufers"
            FeiertagA(52).DatumA = CDate("24.06." & JahrA)
            FeiertagA(52).arbeitsfrei = False
            FeiertagA(53).NameA = "Fest der Apostel Petrus und Paulus"
            FeiertagA(53).DatumA = CDate("29.06." & JahrA)
            FeiertagA(53).arbeitsfrei = False
            FeiertagA(54).NameA = "Mariä Heimsuchung"
            FeiertagA(54).DatumA = CDate("02.07." & JahrA)
            FeiertagA(54).arbeitsfrei = False
            FeiertagA(55).NameA = "Verklärung des Herrn"
            FeiertagA(55).DatumA = CDate("06.08." & JahrA)
            FeiertagA(55).arbeitsfrei = False
            FeiertagA(56).NameA = "Kreuzerhöhung"
            FeiertagA(56).DatumA = CDate("14.09." & JahrA)
            FeiertagA(56).arbeitsfrei = False
            FeiertagA(57).NameA = "Fest der Erzengel Michael, Gabriel und Raphael"
            FeiertagA(57).DatumA = CDate("29.09." & JahrA)
            FeiertagA(57).arbeitsfrei = False
            FeiertagA(58).NameA = "Allerseelen"
            FeiertagA(58).DatumA = CDate("2.11." & JahrA)
            FeiertagA(58).arbeitsfrei = False
            FeiertagA(59).NameA = "Weißer Sonntag"
            FeiertagA(59).DatumA = DateAdd("d", 7, OsternA)
            FeiertagA(59).arbeitsfrei = False
            FeiertagA(60).NameA = "Johannistag"
            FeiertagA(60).DatumA = CDate("24.06." & JahrA)
            FeiertagA(60).arbeitsfrei = False
            FeiertagA(61).NameA = "Siebenschläfertag"
            FeiertagA(61).DatumA = CDate("27.06." & JahrA)
            FeiertagA(61).arbeitsfrei = False
            FeiertagA(62).NameA = "Gründonnerstag"
            FeiertagA(62).DatumA = DateAdd("d", -3, OsternA)
            FeiertagA(62).arbeitsfrei = False
            FeiertagA(63).NameA = "Karsamstag"
            FeiertagA(63).DatumA = DateAdd("d", -1, OsternA)
            FeiertagA(63).arbeitsfrei = False
            FeiertagA(64).NameA = "Europatag"
            FeiertagA(64).DatumA = CDate("05.05." & JahrA)
            FeiertagA(64).arbeitsfrei = False
            FeiertagA(65).NameA = "Tag der Deutschen Einheit"
            FeiertagA(65).DatumA = CDate("03.10." & JahrA)
            FeiertagA(65).arbeitsfrei = False
            For i As Integer = 0 To 65
                Dim ft As New Helper_FeierTage.TypeFeiertag
                ft.DatumA = FeiertagA(i).DatumA
                ft.NameA = FeiertagA(i).NameA
                ft.arbeitsfrei = FeiertagA(i).arbeitsfrei
                Feiertag.Add(ft)
            Next
            Return Feiertag
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False)
            If Helper.IsIDE() Then Stop
            Return Feiertag
        End Try
    End Function
    Public Shared Function GetBeginnSommerzeit(JahrA As Long) As Date
        Dim i As Long
        Dim gbs As Date
        For i = 31 To 20 Step -1
            If Format(CDate(JahrA & "-03-" & i.ToString), "ddd") = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedDayName(System.DayOfWeek.Sunday) Then
                gbs = CDate(JahrA & "-03-" & i.ToString)
                Exit For
            End If
        Next i
        Return gbs
    End Function
    Public Shared Function GetBeginnWinterzeit(JahrA As Long) As Date
        Dim i As Long
        Dim gbw As Date
        For i = 31 To 20 Step -1
            If Format(CDate(JahrA & "-10-" & i.ToString), "ddd") = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedDayName(System.DayOfWeek.Sunday) Then
                gbw = CDate(JahrA & "-10-" & i.ToString)
                Exit For
            End If
        Next i
        Return gbw
    End Function
    Public Shared Function GetErntedankfest(JahrA As Long) As Date
        Dim i As Long
        Dim gef As Date
        For i = 1 To 16
            If Format(CDate(JahrA & "-10-" & i.ToString), "ddd") = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedDayName(System.DayOfWeek.Sunday) Then
                gef = CDate(JahrA & "-10-" & i.ToString)
                Exit For
            End If
        Next i
        Return gef
    End Function
    Public Shared Function GetMuttertag(JahrA As Long) As Date
        Dim W1 As Boolean
        Dim i As Long
        Dim gmt As Date
        W1 = False
        For i = 1 To 31
            If Format(CDate(JahrA & "-05-" & i.ToString), "ddd") = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedDayName(System.DayOfWeek.Sunday) Then
                If W1 Then
                    gmt = CDate(JahrA & "-05-" & i.ToString)
                    Exit For
                Else
                    W1 = True
                End If
            End If
        Next i
        Return gmt
    End Function
    Public Shared Function GetViertenAdvent(JahrA As Long) As Date
        Dim gva As Date
        Dim Date4Advent As New Date
        For i As Integer = 24 To 1 Step -1
            Date4Advent = CDate(JahrA & "-12-" & i.ToString)
            Dim DayString As String = Format(Date4Advent, "ddd")
            If DayString = CultureInfo.CurrentCulture.DateTimeFormat.GetAbbreviatedDayName(System.DayOfWeek.Sunday) Then
                gva = CDate(JahrA & "-12-" & i.ToString)
                Return gva
            End If
        Next
        Return gva
    End Function
    Friend Shared Function IsWorkDay(ByVal Bundesland As String) As Boolean
        Dim IsWorkDayRet As Boolean = True
        'If Helper_Convert.ConvertToBoolean(SelectReadSettingsFrom(clsMain.RegistryHiveValue, clsMain.RegistryPath & "\Printing", "OnlyPrintOnWeekday")) Then
        'Dim Bundesland As String = Helper_Convert.ConvertToString(SelectReadSettingsFrom(clsMain.RegistryHiveValue, clsMain.RegistryPath & "\Printing", "HolidaysOfCountry"))
        If Bundesland = "" Then
            Bundesland = "Bayern"
        End If
        Dim DayOfWeekInt As Integer = Weekday(Now, FirstDayOfWeek.Monday)
        Dim today As DateTime = Now
        Dim isfeiertag As Boolean = Helper_FeierTage.IsFeiertag(today, CType(Bundesland, Helper_FeierTage.Land))
        If DayOfWeekInt > 5 Then
            IsWorkDayRet = False
        ElseIf isfeiertag Then
            IsWorkDayRet = False
        End If
        'End If
        Return IsWorkDayRet
    End Function
    Public Shared Function IsWorkDayAsWorkdayDef(ByVal Bundesland As String) As WorkdayDef 'HolidayDef
        'Dim aktCulture As CultureInfo = CultureInfo.CurrentCulture
        'Dim ci As CultureInfo = New CultureInfo("de-DE")
        'Thread.CurrentThread.CurrentCulture = ci

        Dim IsWorkDayRet As Boolean = True
        Dim WorkDay As New WorkdayDef
        Dim today As DateTime = Now
        WorkDay.Country = ""
        WorkDay.Datum = today
        WorkDay.isWorkday = True
        WorkDay.Feiertag = ""
        'If Helper_Convert.ConvertToBoolean(SelectReadSettingsFrom(clsMain.RegistryHiveValue, clsMain.RegistryPath & "\Printing", "OnlyPrintOnWeekday"), False) Then
        'Dim Bundesland As String = Helper_Convert.ConvertToString(SelectReadSettingsFrom(clsMain.RegistryHiveValue, clsMain.RegistryPath & "\Printing", "HolidaysOfCountry"), False, "")
        If Bundesland = "" Then
            Bundesland = "Bayern"
        End If
        'Dim ftag As New FeierTage()
        Dim DayOfWeekInt As Integer = Weekday(today, FirstDayOfWeek.Monday)
        Dim holiday As New Helper_FeierTage.HolidayDef
        WorkDay.Country = Bundesland
        WorkDay.Datum = today
        WorkDay.isWorkday = True
        WorkDay.Feiertag = ""

        'Dim nFeierTageLand As Helper_FeierTage.Land
        'nFeierTageLand = CType(System.Enum.Parse(nFeierTageLand.GetType(), Bundesland), Helper_FeierTage.land)
        'holiday.isHoliday = Helper_FeierTage.IsFeiertag(today,nFeierTageLand)

        holiday.isHoliday = Helper_FeierTage.IsFeiertag(today, Bundesland)
        If DayOfWeekInt > 5 AndAlso Not holiday.isHoliday Then
            IsWorkDayRet = False
            WorkDay.Country = Bundesland
            WorkDay.Datum = today
            WorkDay.isWorkday = False
            WorkDay.Feiertag = today.ToString("dddd")
        ElseIf holiday.isHoliday Then
            IsWorkDayRet = False
            WorkDay.Country = Bundesland
            WorkDay.Datum = holiday.Datum
            WorkDay.isWorkday = False
            WorkDay.Feiertag = holiday.Feiertag
        End If
        'End If
        Return WorkDay  'IsWorkDayRet
        'Thread.CurrentThread.CurrentCulture = aktCulture
    End Function
    Public Shared Function SecToTime(ByVal Seconds As Integer, Optional ByRef rHour As Integer = 0, Optional ByRef rMinute As Integer = 0, Optional ByRef rSecond As Integer = 0) As String
        rHour = (Seconds \ 3600)
        rMinute = (Seconds - (rHour * 3600)) \ 60
        rSecond = (Seconds - (rHour * 3600) - (rMinute * 60))
        Return Format(rHour, "00") & ":" & Format(rMinute, "00") & ":" & Format(rSecond, "00")
    End Function
    Public Shared Function TimeToUnix(ByVal dteDate As Date) As Long
        If dteDate.IsDaylightSavingTime Then
            dteDate = DateAdd(DateInterval.Hour, -1, dteDate)
        End If
        Return DateDiff(DateInterval.Second, #1/1/1970#, dteDate)
    End Function
    Public Shared Function UnixToTime(ByVal strUnixTime As String) As Date
        Dim UnixToTimeVar As Date = DateAdd(DateInterval.Second, Val(strUnixTime), #1/1/1970#)
        If UnixToTimeVar.IsDaylightSavingTime Then
            UnixToTimeVar = DateAdd(DateInterval.Hour, 1, UnixToTimeVar)
        End If
        Return UnixToTimeVar
    End Function
    Public Shared Function valDate(datum As String, Optional Dates As ToDates = ToDates.DateTime) As String
        Dim dDate As String
        datum = datum.Trim
        If datum <> "" AndAlso Not System.Convert.IsDBNull(datum) AndAlso datum <> "1700-01-01 00:00:00" AndAlso IsDate(datum) Then
            Select Case Dates
                Case ToDates.DateOnly
                    dDate = Format(CDate(datum), "yyyy-MM-dd")
                Case ToDates.TimeOnly
                    dDate = Format(CDate(datum), "HH:mm:ss")
                Case ToDates.DateTime
                    dDate = Format(CDate(datum), "yyyy-MM-dd HH:mm:ss")
                Case Else
                    dDate = Format(CDate(datum), "yyyy-MM-dd HH:mm:ss")
            End Select
        Else
            dDate = "1700-01-01 00:00:00"
        End If
        Return dDate
    End Function
    Public Shared Function valDate(datum As Date, Optional Dates As ToDates = ToDates.DateTime) As Date
        Dim dDate As New Date
        If Not String.IsNullOrEmpty(datum.ToString) AndAlso Not System.Convert.IsDBNull(datum) AndAlso datum <> CDate("1700-01-01 00:00:00") AndAlso IsDate(datum) Then
            Select Case Dates
                Case ToDates.DateOnly
                    dDate = CDate(Format(CDate(datum), "yyyy-MM-dd"))
                Case ToDates.TimeOnly
                    dDate = CDate(Format(CDate(datum), "HH:mm:ss"))
                Case ToDates.DateTime
                    dDate = CDate(Format(CDate(datum), "yyyy-MM-dd HH:mm:ss"))
                Case Else
                    dDate = CDate(Format(CDate(datum), "yyyy-MM-dd HH:mm:ss"))
            End Select
        End If
        Return dDate
    End Function
    Public Shared Sub wait(ByVal milliseconds As Integer)
        For x As Integer = 1 To milliseconds
            Thread.Sleep(1)
            Application.DoEvents()
        Next
    End Sub
    Public Shared Function CalcHours(ByVal tvalue1 As ComboBoxEx, ByVal tvalue2 As ComboBoxEx, ByVal tvalue3 As ComboBoxEx, ByVal tvalue4 As ComboBoxEx) As String
        Try
            'a = tvalue1.Name
            'a = tvalue1.Text

            Dim regex As Regex = New Regex("^([0-9]|0[0-9]|1[0-9]|2[0-3]):[0-5][0-9]$")
            Dim match As Match = regex.Match(tvalue1.Text)
            If Not match.Success Then
                tvalue1.Text = "00:00"
            End If
            match = regex.Match(tvalue2.Text)
            If Not match.Success Then
                tvalue2.Text = "00:00"
            End If
            match = regex.Match(tvalue3.Text)
            If Not match.Success Then
                tvalue3.Text = "00:00"
            End If
            match = regex.Match(tvalue4.Text)
            If Not match.Success Then
                tvalue4.Text = "00:00"
            End If


            Dim tText1 As String = String2Time(tvalue1.Text)
            Dim tText2 As String = String2Time(tvalue2.Text)
            Dim tText3 As String = String2Time(tvalue3.Text)
            Dim tText4 As String = String2Time(tvalue4.Text)
            tvalue1.Text = tText1
            tvalue2.Text = tText2
            tvalue3.Text = tText3
            tvalue4.Text = tText4

            Dim tvaluestr1 As String() = tText1.Split(CType(":", Char))
            If tText1 = "00:00" Then tText2 = "00:00"
            Dim tvaluestr2 As String() = tText2.Split(CType(":", Char))
            Dim tvaluestr3 As String() = tText3.Split(CType(":", Char))
            If tText3 = "00:00" Then tText4 = "00:00"
            Dim tvaluestr4 As String() = tText4.Split(CType(":", Char))

            Dim minv1 As Integer = Helper_VarConvert.ConvertToInteger(tvaluestr1(0), 0) * 60 + Helper_VarConvert.ConvertToInteger(tvaluestr1(1), 0) 'convert2Minutes(tvalue1.Value)
            Dim minv2 As Integer = Helper_VarConvert.ConvertToInteger(tvaluestr2(0), 0) * 60 + Helper_VarConvert.ConvertToInteger(tvaluestr2(1), 0) 'convert2Minutes(tvalue2.Value)
            Dim minn1 As Integer = Helper_VarConvert.ConvertToInteger(tvaluestr3(0), 0) * 60 + Helper_VarConvert.ConvertToInteger(tvaluestr3(1), 0) 'convert2Minutes(tvalue3.Value)
            Dim minn2 As Integer = Helper_VarConvert.ConvertToInteger(tvaluestr4(0), 0) * 60 + Helper_VarConvert.ConvertToInteger(tvaluestr4(1), 0) 'convert2Minutes(tvalue4.Value)

            Dim Minuten As Integer = (minv2 - minv1) + (minn2 - minn1)
            Dim hours As Integer = Fix(Minuten \ 60)
            Dim minutes As Integer = Minuten - (hours * 60)
            Dim timeElapsed As String = Format(hours, "00") & ":" & Format(minutes, "00")
            'If timeElapsed = "23:59" Then
            'Stop
            'End If
            Return timeElapsed

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False, "")
            If Helper.IsIDE() Then Stop
            Return "00:00"
        End Try
    End Function
    Public Shared Function CalcHours(ByVal tvalue1 As DateTimeInput, ByVal tvalue2 As DateTimeInput, ByVal tvalue3 As DateTimeInput, ByVal tvalue4 As DateTimeInput) As String
        'a = tvalue2.Name
        'Dim b1 As String = tvalue1.Value.ToShortTimeString
        'Dim b2 As String = tvalue2.Value.ToShortTimeString
        'Dim b3 As String = tvalue3.Value.ToShortTimeString
        'Dim b4 As String = tvalue4.Value.ToShortTimeString

        Dim tvaluestr1() As String = tvalue1.Value.ToShortTimeString.Split(CType(":", Char))
        If tvalue1.Value.ToShortTimeString = "00:00" Then tvalue2.Text = "00:00"
        Dim tvaluestr2() As String = tvalue2.Value.ToShortTimeString.Split(CType(":", Char))
        Dim tvaluestr3() As String = tvalue3.Value.ToShortTimeString.Split(CType(":", Char))
        If tvalue3.Value.ToShortTimeString = "00:00" Then tvalue4.Text = "00:00"
        Dim tvaluestr4() As String = tvalue4.Value.ToShortTimeString.Split(CType(":", Char))

        Dim minv1 As Integer = Helper_VarConvert.ConvertToInteger(tvaluestr1(0)) * 60 + Helper_VarConvert.ConvertToInteger(tvaluestr1(1)) 'convert2Minutes(tvalue1.Value)
        Dim minv2 As Integer = Helper_VarConvert.ConvertToInteger(tvaluestr2(0)) * 60 + Helper_VarConvert.ConvertToInteger(tvaluestr2(1)) 'convert2Minutes(tvalue2.Value)
        Dim minn1 As Integer = Helper_VarConvert.ConvertToInteger(tvaluestr3(0)) * 60 + Helper_VarConvert.ConvertToInteger(tvaluestr3(1)) 'convert2Minutes(tvalue3.Value)
        Dim minn2 As Integer = Helper_VarConvert.ConvertToInteger(tvaluestr4(0)) * 60 + Helper_VarConvert.ConvertToInteger(tvaluestr4(1)) 'convert2Minutes(tvalue4.Value)

        Dim Minuten As Integer = (minv2 - minv1) + (minn2 - minn1)
        Dim hours As Integer = Fix(Minuten \ 60)
        Dim minutes As Integer = Minuten - (hours * 60)
        Dim timeElapsed As String = Format(hours, "00") & ":" & Format(minutes, "00")
        'If timeElapsed = "23:59" Then
        'Stop
        'End If
        Return timeElapsed
    End Function
    Public Shared Function convert2Minutes(ByVal TimeStr As Date) As Integer
        Dim tn1 As String = ValidateTime(Format(TimeStr, "HH:mm"))
        Dim dtn As DateTime
        If DateTime.TryParse(tn1, dtn) Then
            tn1 = dtn.ToString("HH:mm")
        End If
        Dim min As Integer = Helper_VarConvert.ConvertToInteger(Split(tn1, ":")(0), 0) * 60 + Helper_VarConvert.ConvertToInteger(Split(tn1, ":")(1), 0)
        Return min
    End Function
    Public Shared Function convert2Minutes(ByVal TimeStr As String) As Integer
        Dim tn1 As String = ValidateTime(TimeStr)
        Dim dtn As DateTime
        If DateTime.TryParse(tn1, dtn) Then
            tn1 = dtn.ToString("HH:mm")
        End If
        Dim min As Integer = Helper_VarConvert.ConvertToInteger(Split(tn1, ":")(0), 0) * 60 + Helper_VarConvert.ConvertToInteger(Split(tn1, ":")(1), 0)
        Return min
    End Function
    Public Shared Function ValidateTime(ByVal TimeString As String) As String
        Dim culture As CultureInfo
        Dim styles As DateTimeStyles
        Dim dateResult As DateTime
        culture = CultureInfo.CreateSpecificCulture("de-DE")
        styles = DateTimeStyles.None
        If DateTime.TryParse(TimeString, culture, styles, dateResult) Then
            Dim ts() As String = TimeString.Split(CType(":", Char))
            If ts(0) Is Nothing OrElse String.IsNullOrWhiteSpace(ts(0)) Then
                ts(0) = "00"
            End If
            If ts(1) Is Nothing OrElse String.IsNullOrWhiteSpace(ts(1)) Then
                ts(1) = "00"
            End If
            Return ts(0) & ":" & ts(1)
        Else
            Return "00:00"
        End If
    End Function


    Public Shared Function GetFirstDayOfMonth(ByVal Datum As Date) As Date
        Dim FirstDay As New Date
        FirstDay = CDate(Datum.ToString("yyyy-MM") & "-01")
        Return FirstDay
    End Function
    Public Shared Function GetLastDayOfMonth(ByVal Datum As Date) As Date
        Dim DaysInMonth As Integer = Date.DaysInMonth(Datum.Year, Datum.Month)
        Dim LastDay As Date = New Date(Datum.Year, Datum.Month, DaysInMonth)
        Return LastDay
    End Function
    Public Shared Function getFirstDayOfWeek(ByVal SearchDate As Date) As Date
        Dim dayIndex As Integer = SearchDate.DayOfWeek
        If dayIndex < System.DayOfWeek.Monday Then
            dayIndex += 7
        End If
        Dim dayDiff As Integer = dayIndex - System.DayOfWeek.Monday
        Dim monday As Date = SearchDate.AddDays(-dayDiff)
        Return monday
    End Function
    Public Shared Function getLastDayOfWeek(ByVal SearchDate As Date) As Date
        Dim sunday As Date = DateAdd(DateInterval.Day, 6, getFirstDayOfWeek(SearchDate))
        Return sunday
    End Function

    Public Shared Function GetWeekStartDate1(weekNumber As Integer, year As Integer) As Date
        Dim startDate As New DateTime(year, 1, 1)
        Dim weekDate As DateTime = DateAdd(DateInterval.WeekOfYear, weekNumber - 1, startDate)
        Return DateAdd(DateInterval.Day, (-weekDate.DayOfWeek) + 1, weekDate)
    End Function
    Public Shared Function GetWeekStartDate(ByVal Week As Integer, ByVal Year As Integer) As Date
        Dim FirstDayOfWeek As System.DayOfWeek = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek  ' System.DayOfWeek.Monday
        Dim RetVal As New Date
        Dim dt As Date = New Date(Year, 1, 1)
        If dt.DayOfWeek > 4 Then dt = dt.AddDays(7 - dt.DayOfWeek) Else dt = dt.AddDays(-dt.DayOfWeek)
        dt = dt.AddDays(FirstDayOfWeek)
        RetVal = dt.AddDays(7 * (Week - 1))
        Return RetVal
    End Function
    Public Shared Function GetMondayFromDate(ByVal Datum As Date) As Date
        Dim dayIndex As Integer = Datum.DayOfWeek
        If dayIndex < System.DayOfWeek.Monday Then
            dayIndex += 7 'Monday is first day of week, no day of week should have a smaller index
        End If
        Dim dayDiff As Integer = dayIndex - System.DayOfWeek.Monday
        Dim monday As Date = Datum.AddDays(-dayDiff)
        Return monday
    End Function
    Public Shared Function FirstDateOfWeekISO8601(ByVal weekOfYear As Integer, ByVal year As Integer) As DateTime
        Dim jan1 As DateTime = New DateTime(year, 1, 1)
        Dim daysOffset As Integer = System.DayOfWeek.Thursday - jan1.DayOfWeek
        Dim firstThursday As DateTime = jan1.AddDays(daysOffset)
        Dim cal As Calendar = CultureInfo.CurrentCulture.Calendar
        Dim firstWeek As Integer = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, System.DayOfWeek.Monday)
        Dim weekNum As Integer = weekOfYear
        If firstWeek <= 1 Then
            weekNum -= 1
        End If

        Dim result As Date = firstThursday.AddDays(weekNum * 7)
        Return result.AddDays(-3)
    End Function
    Public Shared Function getWeekNumber_old(ByVal SearchDate As Date) As Integer
        Dim dfi As DateTimeFormatInfo = DateTimeFormatInfo.CurrentInfo
        Dim cal As Calendar = dfi.Calendar
        dfi.CalendarWeekRule = CalendarWeekRule.FirstDay
        cal.GetWeekOfYear(SearchDate, dfi.CalendarWeekRule, dfi.FirstDayOfWeek)
        Dim weekOfyear As Integer = cal.GetWeekOfYear(SearchDate, dfi.CalendarWeekRule, System.DayOfWeek.Monday)
        Return weekOfyear

    End Function
    Public Shared Function CalendarWeek(ByVal nWeek As Integer, ByVal nYear As Integer) As Date
        ' Wochentag des 4. Januar des Jahres ermitteln
        Dim dStart As New Date(nYear, 1, 4)
        Dim nDay As Integer = (dStart.DayOfWeek + 6) Mod 7 + 1
        ' Beginn der 1. KW des Jahres
        Dim dFirst As Date = dStart.AddDays(1 - nDay)
        ' Gesuchte KW ermitteln
        Return dFirst.AddDays((nWeek - 1) * 7)
    End Function
    Public Shared Function getWeekNumber(ByVal dDate As Date) As Integer
        Dim WeekNr As Integer = 0
        ' Startdatum der ersten Kalenderwoche des Jahres und Folgejahres berechnen
        Dim dThisYear As Date = GetWeekStartDate(1, dDate.Year)
        Dim dNextYear As Date = GetWeekStartDate(1, dDate.Year + 1)

        ' Prüfen, ob Datum zur ersten Woche des Folgejahres gehört
        If dDate >= dNextYear Then
            ' Rückgabe: KW 1 des Folgejahres
            WeekNr = 1
        Else
            ' KW = Differenz zum ersten Tag der ersten Woche
            WeekNr = dDate.Subtract(dThisYear).Days \ 7 + 1
        End If

        Return WeekNr
    End Function


    Public Shared Function TimeDiff(ByVal sTime1 As String, ByVal sTime2 As String, ByVal sTime3 As String, ByVal sTime4 As String) As TimeAddReturn
        Dim time1 As New TimeSpan
        Dim time2 As New TimeSpan
        Dim time3 As New TimeSpan
        Dim time4 As New TimeSpan
        Dim nMin As Integer = 0
        Dim nMin1 As Integer = 0
        Dim nMin2 As Integer = 0
        Dim RetVal As New TimeAddReturn
        Try
            If sTime1 = "" Then sTime1 = "00:00:00"
            If sTime2 = "" Then sTime2 = "00:00:00"
            If sTime3 = "" Then sTime3 = "00:00:00"
            If sTime4 = "" Then sTime4 = "00:00:00"
            time1 = TimeSpan.Parse(sTime1)
            time2 = TimeSpan.Parse(sTime2)
            time3 = TimeSpan.Parse(sTime3)
            time4 = TimeSpan.Parse(sTime4)

            If Date.TryParse(sTime1, New Date) AndAlso Date.TryParse(sTime2, New Date) Then
                With time2.Subtract(time1)
                    nMin1 = .Hours * 60 + .Minutes
                    If nMin1 < 0 Then nMin1 += 24 * 60
                End With
            End If
            If Date.TryParse(sTime3, New Date) AndAlso Date.TryParse(sTime4, New Date) Then
                With time4.Subtract(time3)
                    nMin2 = .Hours * 60 + .Minutes
                    If nMin2 < 0 Then nMin2 += 24 * 60
                End With
            End If
            nMin = nMin1 + nMin2

            Dim sTime As String = New TimeSpan(0, nMin, 0).ToString.Substring(0, 5)

            RetVal.Hours = sTime
            RetVal.Minutes = nMin
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
            RetVal.Hours = "00: 00"
            RetVal.Minutes = 0
        End Try
        Return RetVal
    End Function
    Public Shared Function AddHours(ByVal TotalHours As String, ByVal Hours2Add As String) As String
        If TotalHours = "" Then
            TotalHours = "00:00"
        End If
        Dim tmpTotal() As String = TotalHours.Split(CType(":", Char))
        Dim int_TotalHours As Integer = Helper_VarConvert.ConvertToInteger(tmpTotal(0), 0) * 60 + Helper_VarConvert.ConvertToInteger(tmpTotal(1), 0)
        Dim tmpHours() As String = Hours2Add.Split(CType(":", Char))
        Dim int_Hours2Add As Integer = Helper_VarConvert.ConvertToInteger(tmpHours(0), 0) * 60 + Helper_VarConvert.ConvertToInteger(tmpHours(1), 0)
        Dim sum As Integer = int_TotalHours + int_Hours2Add
        Dim sumHours As Integer = Helper_VarConvert.ConvertToInteger(Fix(sum / 60), 0)
        Dim sumMinutes As Integer = sum - (sumHours * 60)
        Return Helper_VarConvert.ConvertToString(sumHours, False, "") & ":" & Helper_VarConvert.ConvertToString(sumMinutes, False, "")
    End Function
    Public Shared Function HHmm2Minutes(ByVal TimeString As String) As Double
        Dim hours As Double = 0
        If TimeString.Contains(":") Then
            Dim ts() As String = Split(TimeString, ":")
            hours = Math.Round(Helper_VarConvert.ConvertToDouble(ts(0)) + (Helper_VarConvert.ConvertToDouble(ts(1)) / 60), 2)
        End If
        Return hours

    End Function
    Public Shared Function mmss2Seconds(ByVal TimeString As String) As Double
        Dim minutes As Double = 0
        If TimeString.Contains(":") Then
            Dim ts() As String = Split(TimeString, ":")
            minutes = Math.Round(Helper_VarConvert.ConvertToDouble(ts(0)) + (Helper_VarConvert.ConvertToDouble(ts(1)) / 60), 2)
        End If
        Return minutes

    End Function
    Public Shared Function Number2MonthName(ByVal MonthNumber As Integer, ByVal Optional ShortNames As Boolean = False) As String
        Dim RetVal As String = ""
        Select Case MonthNumber
            Case 1
                RetVal = "Januar"
            Case 2
                RetVal = "Februar"
            Case 3
                RetVal = "März"
            Case 4
                RetVal = "April"
            Case 5
                RetVal = "Mai"
            Case 6
                RetVal = "Juni"
            Case 7
                RetVal = "Juli"
            Case 8
                RetVal = "August"
            Case 9
                RetVal = "September"
            Case 10
                RetVal = "Oktober"
            Case 11
                RetVal = "November"
            Case 12
                RetVal = "Dezember"
        End Select
        Return RetVal
    End Function
    Public Shared Function FormatTimeDez(ByVal nMin As Double) As Double
        Dim RetVal As Double = 0
        Dim nStd As Double = 0
        Dim nMin1 As Double = 0
        ' Ganze Stunden ermitteln
        nStd = Math.Truncate(nMin / 60)
        ' Rest-Minuten
        nMin1 = nMin - (nStd * 60)
        ' formatierter Rückgabewert
        ' Minuten in dezimaler Schreibweise
        If Application.CurrentCulture.Name = "de-DE" Then
            RetVal = Helper_VarConvert.ConvertToDouble((nStd) & "," & (nMin1 / 60 * 100).ToString("00"))
        Else
            RetVal = Helper_VarConvert.ConvertToDouble((nStd) & "." & (nMin1 / 60 * 100).ToString("00"))
        End If
        Return RetVal
    End Function
    Public Shared Function String2Time(ByVal TimeString As String) As String
        Dim timestr As String = TimeString
        Dim hour As Integer = 0
        Dim min As Integer = 0
        If Not TimeString.Contains(":") AndAlso TimeString.Length >= 2 Then
            If timestr.Contains(":") Then
                Dim ri As Integer = timestr.IndexOf(":")
                hour = Helper_VarConvert.ConvertToInteger(Mid(timestr, 1, ri - 1), 0)
                min = Helper_VarConvert.ConvertToInteger(Mid(timestr, ri + 1, 1))
            ElseIf timestr.Length = 4 Then
                If IsNumeric(Mid(timestr, 1, 2)) AndAlso Helper_VarConvert.ConvertToInteger(Mid(timestr, 1, 2)) <= 23 Then
                    hour = Helper_VarConvert.ConvertToInteger(Mid(timestr, 1, 2))
                    min = Helper_VarConvert.ConvertToInteger(Mid(timestr, 3, 2))
                ElseIf IsNumeric(Mid(timestr, 1, 1)) AndAlso IsNumeric(Mid(timestr, 2, 2)) Then
                    hour = Helper_VarConvert.ConvertToInteger(Mid(timestr, 1, 1))
                    min = Helper_VarConvert.ConvertToInteger(Mid(timestr, 2, 2))
                End If
            ElseIf timestr.Length = 3 Then
                If IsNumeric(Mid(timestr, 1, 2)) AndAlso Helper_VarConvert.ConvertToInteger(Mid(timestr, 1, 2)) <= 23 Then
                    hour = Helper_VarConvert.ConvertToInteger(Mid(timestr, 1, 2))
                    min = Helper_VarConvert.ConvertToInteger(Mid(timestr, 3, 1))
                ElseIf IsNumeric(Mid(timestr, 1, 1)) AndAlso IsNumeric(Mid(timestr, 2, 2)) Then
                    hour = Helper_VarConvert.ConvertToInteger(Mid(timestr, 1, 1))
                    min = Helper_VarConvert.ConvertToInteger(Mid(timestr, 2, 2))
                End If
            ElseIf timestr.Length = 2 Then
                If IsNumeric(timestr) AndAlso Helper_VarConvert.ConvertToInteger(timestr) <= 23 Then
                    hour = Helper_VarConvert.ConvertToInteger(timestr)
                    min = 0
                ElseIf IsNumeric(Mid(timestr, 1, 1)) AndAlso IsNumeric(Mid(timestr, 2, 1)) Then
                    hour = Helper_VarConvert.ConvertToInteger(Mid(timestr, 1, 1))
                    min = Helper_VarConvert.ConvertToInteger(Mid(timestr, 2, 1))
                End If
            End If
            timestr = String.Format("{0:00}", hour) & ":" & String.Format("{0:00}", min)
        End If
        Return timestr
    End Function
    Public Shared Function MonthFromInteger(ByVal value As Integer) As String
        If value = 1 Then
            Return "Januar"
        ElseIf value = 2 Then
            Return "Februar"
        ElseIf value = 3 Then
            Return "März"
        ElseIf value = 4 Then
            Return "April"
        ElseIf value = 5 Then
            Return "Mai"
        ElseIf value = 6 Then
            Return "Juni"
        ElseIf value = 7 Then
            Return "Juli"
        ElseIf value = 8 Then
            Return "August"
        ElseIf value = 9 Then
            Return "September"
        ElseIf value = 10 Then
            Return "Oktober"
        ElseIf value = 11 Then
            Return "November"
        ElseIf value = 12 Then
            Return "Dezember"
        Else
            Return ""
        End If
    End Function
    Private Shared Function GetTimeInterval(ByVal nSeks As Long) As String
        Dim h As Long, m As Long
        Dim sInterv As String
        h = nSeks \ 3600
        nSeks = nSeks Mod 3600
        m = nSeks \ 60
        nSeks = nSeks Mod 60
        sInterv = Format(h, "00") & ":" & Format(m, "00") & ":" & Format(nSeks, "00")
        Return sInterv
    End Function
    Public Shared Function FindDateTimePickerStartsWith(root As Control, name As String) As List(Of DateTimePicker)
        Dim controls As New List(Of DateTimePicker)
        Try
            If root IsNot Nothing Then
                For Each ctl As Control In root.Controls
                    If ctl.Name.StartsWith(name) Then
                        controls.Add(CType(ctl, DateTimePicker))
                    End If
                Next
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return controls
    End Function
    Public Shared Function Time2Double(ByVal TimeString As String) As Double
        Dim hours As Double = 0
        If TimeString.Contains(":") Then
            Dim ts() As String = Split(TimeString, ":")
            hours = Math.Round(CDbl(ts(0)) + (CDbl(ts(1)) / 60), 2)
        End If
        Return hours

    End Function

#End Region
#Region "Email"
    Public Shared Function isEmail(ByVal email As String) As Boolean
        Dim emailRegex As New Regex("([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)", RegexOptions.IgnoreCase)
        Return emailRegex.IsMatch(email)
    End Function
    Public Shared Function isEmail(ByVal email As String, ByVal ReturnEmail As Boolean) As String
        Dim emailRegex As New Regex("([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)", RegexOptions.IgnoreCase)
        Dim ret As Match = emailRegex.Match(email)
        Return ret.Value
    End Function
    Public Shared Function SendSystemEmailUniversal(ByVal RegistryHiveValue As RegistryHive,
                                                ByVal RegPath As String, ByVal Header As String,
                                                ByVal MessageText As String,
                                                ByVal Optional MessageHtml As String = "",
                                                ByVal Optional PathAttachment As List(Of String) = Nothing,
                                                ByVal Optional SMTPCred As SMTPCredentials = Nothing,
                                                ByVal Optional Bitmaps As List(Of Bitmap) = Nothing) As Boolean
        Dim ReturnVal As Boolean = False
        Dim builder As New MailBuilder()
        Dim attachment As MimeData = Nothing

        Try
            If SMTPCred.SMTP_SenderAddress Is Nothing Then
                SMTPCred.SMTP_Password = DecryptString(Helper_VarConvert.ConvertToString(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_Password")))
                SMTPCred.SMTP_User = Helper_VarConvert.ConvertToString(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_User"))
                SMTPCred.SMTP_Server = Helper_VarConvert.ConvertToString(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_RelayServer"))
                SMTPCred.SMTP_NoSSL = Helper_VarConvert.ConvertToBoolean(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_NoSSL"))
                SMTPCred.SMTP_SSL = Helper_VarConvert.ConvertToBoolean(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_SSL"))
                SMTPCred.SMTP_STARTTLS = Helper_VarConvert.ConvertToBoolean(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_STARTTLS"))
                SMTPCred.SMTP_SenderAddress = Helper_VarConvert.ConvertToString(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_SenderAddress"))
                SMTPCred.SMTP_SenderName = Helper_VarConvert.ConvertToString(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_SenderName"))
                SMTPCred.SMTP_RecipientAddress = Helper_VarConvert.ConvertToString(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_RecipientAddress"))
                SMTPCred.SMTP_RecipientName = Helper_VarConvert.ConvertToString(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_RecipientName"))
                SMTPCred.SMTP_CC = Helper_VarConvert.ConvertToString(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_CC"))
                SMTPCred.SMTP_BCC = Helper_VarConvert.ConvertToString(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_BCC"))
                SMTPCred.SMTP_NoSSL_Port = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_NoSSL_Port"))
                SMTPCred.SMTP_SSL_Port = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_SSL_Port"))
                SMTPCred.SMTP_STARTTLS_Port = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(RegistryHiveValue, RegistryPath & "\email", "SMTP_STARTTLS_Port"))

            End If
            If SMTPCred.SMTP_SenderAddress <> "" Then
                builder.From.Add(New MailBox(SMTPCred.SMTP_SenderAddress, SMTPCred.SMTP_SenderName))
                builder.To.Add(New MailBox(SMTPCred.SMTP_RecipientAddress, SMTPCred.SMTP_RecipientName))
                If SMTPCred.SMTP_CC <> "" Then
                    builder.Cc.Add(New MailBox(SMTPCred.SMTP_CC))
                End If
                If SMTPCred.SMTP_BCC <> "" Then
                    builder.Bcc.Add(New MailBox(SMTPCred.SMTP_BCC))
                End If
                builder.Subject = Header
                builder.Text = MessageText
                If MessageHtml <> "" Then
                    builder.Html = MessageHtml
                End If

                If PathAttachment IsNot Nothing Then
                    For Each Anhang As String In PathAttachment
                        If Anhang <> "" AndAlso File.Exists(Anhang) Then
                            attachment = builder.AddAttachment(Anhang)
                        End If
                    Next
                End If
            End If

            Dim c As Integer = 0
            If Bitmaps IsNot Nothing AndAlso Bitmaps.Count > 0 Then

                Dim temp As String = GetTempPath()

                For Each bMap As Bitmap In Bitmaps
                    c += 1
                    Dim filename As String = temp & "Screenshot " & c.ToString & ".jpg"
                    bMap.Save(filename, ImageFormat.Jpeg)
                    attachment = builder.AddAttachment(filename)
                    File.Delete(filename)
                Next
            End If

            If Not String.IsNullOrEmpty(SMTPCred.SMTP_Password) AndAlso
                                                   Not String.IsNullOrEmpty(SMTPCred.SMTP_User) AndAlso
                                                  Not String.IsNullOrEmpty(SMTPCred.SMTP_SenderAddress) AndAlso
                                                  Not String.IsNullOrEmpty(SMTPCred.SMTP_RecipientAddress) AndAlso
                                                  Not String.IsNullOrEmpty(SMTPCred.SMTP_RecipientAddress) AndAlso
                                                  Not String.IsNullOrEmpty(SMTPCred.SMTP_Server) Then
                Dim email As IMail = builder.Create()
                Using Smtp As New Smtp()
                    Try

                        If SMTPCred.SMTP_NoSSL Then
                            Smtp.Connect(SMTPCred.SMTP_Server, SMTPCred.SMTP_NoSSL_Port, False)
                            Smtp.UseBestLogin(SMTPCred.SMTP_User, SMTPCred.SMTP_Password)
                        ElseIf SMTPCred.SMTP_SSL Then
                            Smtp.ConnectSSL(SMTPCred.SMTP_Server, SMTPCred.SMTP_SSL_Port)
                            Smtp.UseBestLogin(SMTPCred.SMTP_User, SMTPCred.SMTP_Password)
                        ElseIf SMTPCred.SMTP_STARTTLS Then
                            Smtp.Connect(SMTPCred.SMTP_Server, SMTPCred.SMTP_STARTTLS_Port)
                            Smtp.StartTLS()
                            Smtp.UseBestLogin(SMTPCred.SMTP_User, SMTPCred.SMTP_Password)
                        End If
                        Dim result As ISendMessageResult = Smtp.SendMessage(email)
                        If result.Status = SendMessageStatus.Success Then
                            ReturnVal = True
                        Else
                            ReturnVal = False
                        End If

                    Catch ex As Exception
                        'Call Helper_Logger.writelog(Level.Error, "Fehler beim Versenden der Email: " & ex.Message, ex)
                        If Helper.IsIDE() Then
                            MessageBoxEx.Show("Es ist ein Fehler beim Versenden einer System-Email aufgetreten:" & Environment.NewLine & "" & Environment.NewLine & ex.Message, "Mailversand Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        ReturnVal = False
                    End Try
                End Using
            End If
            Return ReturnVal

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
            Return False
        End Try
    End Function
#End Region
#Region "Windows & Forms"
    Public Shared Sub CenterFormOnParent(ByVal ParentWindow As Form, ByVal Window As Form, ByVal RegistryHiveValue As RegistryHive, ByVal RegPath As String)
        'Dim ParentWindowName As String = ParentWindow.Name
        Dim lp As New Point
        Dim Window_Location As New WindowLocation

        Try
            Window_Location = Helper.ReadWindowPosition(RegistryHiveValue, RegistryPath, ParentWindow, False, False, False)

            Dim pX As Integer = Window_Location.fLocation.X
            Dim pY As Integer = Window_Location.fLocation.Y
            Dim pWiHeight As Integer = Window_Location.fHeight
            Dim pWiWidth As Integer = Window_Location.fWidth
            Dim WiHeight As Integer = Window.Size.Height
            Dim WiWidth As Integer = Window.Size.Width
            lp.X = System.Convert.ToInt32(pX + (pWiWidth - WiWidth) / 2)
            lp.Y = System.Convert.ToInt32(pY + (pWiHeight - WiHeight) / 2)
            Window.Location = lp

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
    End Sub
    ' ---------- Invoke_grafik_CopyFromScreen --------------------------------------------------
    Private Shared Function Invoke_grafik_CopyFromScreen(ByVal Target As Form) As Bitmap
        Dim ScreenShot As New Bitmap(Target.Width, Target.Height)
        If Target.InvokeRequired Then
            Target.BeginInvoke(Function()
                                   Using grafik As Graphics = Graphics.FromImage(ScreenShot)
                                       grafik.CopyFromScreen(Target.PointToScreen(New Point(0, 0)), New Point(0, 0), New Size(Target.Width, Target.Height))
                                   End Using
                                   Return ScreenShot
                               End Function)
            Return ScreenShot
        Else
            Using grafik As Graphics = Graphics.FromImage(ScreenShot)
                grafik.CopyFromScreen(Target.PointToScreen(New Point(0, 0)), New Point(0, 0), New Size(Target.Width, Target.Height))
            End Using
            Return ScreenShot
        End If
    End Function
    Public Shared Function IsFormLoaded(ByVal sName As String) As Boolean
        Dim bResult As Boolean = False
        'Dim Hwnd As IntPtr
        For Each oForm As Form In Application.OpenForms
            If oForm.Name.ToLower = sName.ToLower Then
                'hwnd = oForm.Handle
                bResult = True : Exit For
            End If
        Next
        Return (bResult)
    End Function
    Public Shared Function MakeScreenShot(ByVal Form As Form) As Bitmap
        Dim ScreenShot As New Bitmap(Form.Width, Form.Height)
        'Using grafik As Graphics = Graphics.FromImage(ScreenShot)
        '    grafik.CopyFromScreen(Form.PointToScreen(New Point(0, 0)), New Point(0, 0), New Size(Form.Width, Form.Height))
        'End Using
        ScreenShot = Invoke_grafik_CopyFromScreen(Form)
        Return ScreenShot
    End Function
    Public Shared Function ReadWindowPosition(ByVal RegHive As RegistryHive,
                                          ByVal RegPath As String,
                                          ByVal Window As Form,
                                          ByVal Optional SetWindowLocation As Boolean = True,
                                          ByVal Optional SetWindowHeight As Boolean = True,
                                          ByVal Optional SetWindowWidth As Boolean = True,
                                          ByVal Optional MinWidth As Integer = 50,
                                          ByVal Optional MinHeight As Integer = 50,
                                          ByVal Optional WindowAlwaysOnTop As Boolean = False) As WindowLocation
        Dim WindowName As String = Window.Name
        Dim lp As New Point
        Dim Window_X As Integer = 0
        Dim Window_Y As Integer = 0
        Dim Window_Height As Integer = 0
        Dim Window_Width As Integer = 0
        Dim Window_Location As New WindowLocation

        Try
            Dim mscb As multiScreenBounds = GetScreenDimensions()

            Window_X = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(RegHive, RegistryPath & "\Parameter", "Location_" & WindowName & "_X"), 0)
            Window_Y = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(RegHive, RegistryPath & "\Parameter", "Location_" & WindowName & "_Y"), 0)
            lp = New Point(Window_X, Window_Y)
            Window_Height = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(RegHive, RegistryPath & "\Parameter", "Location_" & WindowName & "_Height", Helper.VarTypeEnu.Type_Integer), 0)
            Window_Width = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(RegHive, RegistryPath & "\Parameter", "Location_" & WindowName & "_Width", Helper.VarTypeEnu.Type_Integer), 0)
            Window_Location.fLocation = lp
            Window_Location.fHeight = Window_Height
            Window_Location.fWidth = Window_Width

            If SetWindowLocation Then
                If Window_X < mscb.outerRight OrElse Window_X > mscb.outerLeft Then
                    Window_X = 10
                End If
                If Window_Y < mscb.outerBottom OrElse Window_Y > mscb.outerTop Then
                    Window_Y = 10
                End If
                Window.Location = lp
            End If
            If SetWindowHeight AndAlso (Window.FormBorderStyle = FormBorderStyle.Sizable OrElse Window.FormBorderStyle = FormBorderStyle.SizableToolWindow) Then
                If Window_Height <> Window.Size.Height Then
                    Window.Height = Window_Height
                Else
                    Window.Height = Window.Size.Height
                End If
                If Window.Size.Height < MinHeight Then
                    Window.Height = MinHeight
                End If
            End If
            If SetWindowWidth AndAlso (Window.FormBorderStyle = FormBorderStyle.Sizable OrElse Window.FormBorderStyle = FormBorderStyle.SizableToolWindow) Then
                If Window_Width <> Window.Size.Width Then
                    Window.Width = Window_Width
                Else
                    Window.Width = Window.Size.Width
                End If
                If Window.Size.Width < MinWidth Then
                    Window.Width = MinWidth
                End If
            End If
            If WindowAlwaysOnTop Then
                Window.BringToFront()
                Window.TopMost = True
            End If

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return Window_Location
    End Function
    Public Shared Function fReadWindowPosition(ByVal ReadRegistryHiveValue As RegistryHive,
                                           ByVal RegPath As String,
                                           ByVal Window As Form,
                                           Optional ByVal WindowLocation As Boolean = True,
                                           Optional ByVal WindowHeight As Boolean = True,
                                           Optional ByVal WindowWidth As Boolean = True) As WindowLocation
        Dim WindowName As String = Window.Name
        Dim RetVal As New WindowLocation

        Dim mscb As multiScreenBounds = GetScreenDimensions()

        If WindowLocation Then
            Dim X As Integer = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(ReadRegistryHiveValue, RegistryPath & "\Parameter", "Location_" & WindowName & "_X"))
            Dim Y As Integer = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(ReadRegistryHiveValue, RegistryPath & "\Parameter", "Location_" & WindowName & "_Y"))
            If X < mscb.outerRight OrElse X > mscb.outerLeft Then
                X = 10
            End If
            If Y < mscb.outerBottom OrElse Y > mscb.outerTop Then
                Y = 10
            End If
            Dim lp As New Point(X, Y)
            RetVal.fLocation = lp
        End If
        If WindowHeight Then
            Dim WiHeight As Integer = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(ReadRegistryHiveValue, RegistryPath & "\Parameter", "Location_" & WindowName & "_Height", Helper.VarTypeEnu.Type_Integer))
            If WiHeight > Window.Size.Height Then
                RetVal.fHeight = WiHeight
            Else
                RetVal.fHeight = Window.Size.Height
            End If
        End If
        If WindowWidth Then
            Dim WiWidth As Integer = Helper_VarConvert.ConvertToInteger(SelectReadSettingsFrom(ReadRegistryHiveValue, RegistryPath & "\Parameter", "Location_" & WindowName & "_Width", Helper.VarTypeEnu.Type_Integer))
            If WiWidth > Window.Size.Width Then
                RetVal.fWidth = WiWidth
            Else
                RetVal.fWidth = Window.Size.Width
            End If
        End If
        Return RetVal
    End Function
    Public Shared Function GetScreenDimensions() As multiScreenBounds
        Dim mscb As multiScreenBounds = getMultiScreenBounds()
        For Each scr As Screen In Screen.AllScreens
            If scr.WorkingArea.Left < mscb.outerRight Then
                mscb.outerRight = scr.WorkingArea.Left
            End If
            If scr.WorkingArea.Top < mscb.outerBottom Then
                mscb.outerBottom = scr.WorkingArea.Top
            End If
            If scr.WorkingArea.Left + scr.WorkingArea.Width > mscb.outerLeft Then
                mscb.outerLeft = scr.WorkingArea.Left + scr.WorkingArea.Width
            End If
            If scr.WorkingArea.Top + scr.WorkingArea.Height > mscb.outerTop Then
                mscb.outerTop = scr.WorkingArea.Top + scr.WorkingArea.Height
            End If

        Next

        Return mscb
    End Function
    Public Shared Sub WriteWindowPosition(ByVal RegHive As RegistryHive,
                                      ByVal RegPath As String,
                                      ByVal Window As Form,
                                      Optional ByVal WindowLocation As Boolean = True,
                                      Optional ByVal WindowHeight As Boolean = True,
                                      Optional ByVal WindowWidth As Boolean = True,
                                      Optional ByVal overwrite As Boolean = False,
                                      Optional ByVal StartupComplete As Boolean = False)
        Dim Window_X As Integer = 0
        Dim Window_Y As Integer = 0
        Dim lp As New Point
        Dim mscb As multiScreenBounds = GetScreenDimensions()

        If StartupComplete OrElse overwrite Then
            If Window.WindowState = FormWindowState.Minimized OrElse
                Window.Location.X = 0 OrElse
                Window.Location.Y = 0 Then Exit Sub

            Window_X = Window.Location.X
            Window_Y = Window.Location.Y
            lp = New Point(Window_X, Window_Y)

            If Window.Location.X >= mscb.outerLeft Then
                Window_X = mscb.outerLeft - 5
            End If
            If Window.Location.X <= mscb.outerRight - Window.Size.Width Then
                Window_X = mscb.outerRight - Window.Size.Width - 5
            End If

            If Window.Location.Y >= mscb.outerTop Then
                Window_Y = mscb.outerTop - 5
            End If
            If Window.Location.Y <= mscb.outerBottom - Window.Size.Height Then
                Window_Y = mscb.outerBottom - Window.Size.Height - 5
            End If
            Window.Location = lp

            Dim WindowName As String = Window.Name
            Dim rb As Boolean = False
            If WindowLocation Then
                rb = Helper.RegistryWriteValue(RegHive, RegistryPath & "\Parameter", "Location_" & WindowName & "_X", Window.Location.X)
                rb = Helper.RegistryWriteValue(RegHive, RegistryPath & "\Parameter", "Location_" & WindowName & "_Y", Window.Location.Y)
            End If
            If WindowHeight Then
                rb = Helper.RegistryWriteValue(RegHive, RegistryPath & "\Parameter", "Location_" & WindowName & "_Height", Window.Height)
            End If
            If WindowWidth Then
                rb = Helper.RegistryWriteValue(RegHive, RegistryPath & "\Parameter", "Location_" & WindowName & "_Width", Window.Width)
            End If
        End If
    End Sub
    Private Shared Function getMultiScreenBounds() As multiScreenBounds
        Dim mscb As New multiScreenBounds
        Dim anzScr As Integer = Screen.AllScreens.Length - 1
        Dim bs(anzScr) As Rectangle
        For i As Integer = 0 To anzScr
            bs(i) = Screen.AllScreens(i).Bounds
            If mscb.outerLeft > bs(i).Left Then
                mscb.outerLeft = bs(i).Left
            End If
            If mscb.outerTop > bs(i).Top Then
                mscb.outerTop = bs(i).Top
            End If
            If mscb.outerRight < bs(i).Right Then
                mscb.outerRight = bs(i).Right
            End If
            If mscb.outerBottom < bs(i).Bottom Then
                mscb.outerBottom = bs(i).Bottom
            End If
        Next
        Return mscb
    End Function
    Public Shared Sub SetCtrlBackground(ByVal Frm As Form, CtrlType As System.Type)
        Dim allctl As New List(Of Control)
        For Each ctl As Control In FindControlRecursive(allctl, Frm, CtrlType)
            Application.DoEvents()
            If ctl.BackColor = Color.Gainsboro Then
                ctl.BackColor = Color.Transparent
            End If
        Next
    End Sub
    Public Shared Sub SetCtrlBackground(ByVal Frm As UserControl, CtrlType As System.Type)
        Dim allctl As New List(Of Control)
        For Each ctl As Control In FindControlRecursive(allctl, Frm, CtrlType)
            Application.DoEvents()
            If ctl.BackColor = Color.Gainsboro Then
                ctl.BackColor = Color.Transparent
            End If
        Next
    End Sub
    Public Shared Function FindByTag(ByVal ctrl As Control, ByVal tag As String) As Control
        If ctrl Is Nothing Then
            Return Nothing
        End If

        If TypeOf ctrl.Tag Is String AndAlso CStr(ctrl.Tag) = tag Then
            Return ctrl
        End If

        Return (From control In ctrl.Controls Select FindByTag(CType(control, Control), tag)).FirstOrDefault(Function(c) c IsNot Nothing)
    End Function
    Public Shared Sub CloseOtherMdiForms(ByVal FormNotToClose As String, ByVal MdiParent As Form)
        For Each child As Form In MdiParent.MdiChildren
            'a = child.Name
            If child.Name <> FormNotToClose Then
                child.Close()
                child.Dispose()
            End If
        Next

    End Sub
#End Region
#Region "Images"
    Public Shared Function BytesToImage(ByVal ByteArray() As Byte) As Image
        Dim ResImage As Image = Nothing
        Try
            ' Konvertiert ein Byte-Array in ein Image-Objekt 
            Dim IC As New ImageConverter
            ' Mithilfe des ImageConverter-Objekts ein Byte-Array 
            ' in ein Image überführen 
            ResImage = CType(IC.ConvertFrom(ByteArray), Image)

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return ResImage
    End Function
    Public Shared Function ImageToBytes(ByVal Img As Image) As Byte()
        Dim ResByte As Byte() = Nothing
        Try
            ' Konvertiert ein Image-Objekt in ein Byte-Array 
            Dim IC As New ImageConverter
            ' Mithilfe des ImageConverter-Objekts ein Image 
            ' in ein Byte-Array überführen 
            ResByte = CType(IC.ConvertTo(Img, GetType(Byte())), Byte())

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return ResByte
    End Function
    Public Shared Function byteArrayToImage(ByVal byteArrayIn As Byte()) As Image
        Using mStream As New MemoryStream(byteArrayIn)
            Return Image.FromStream(mStream)
        End Using
    End Function
    Public Shared Function SafeImageFromFile(path As String) As Image
        Dim img As Image = Nothing
        If File.Exists(path) Then
            Using fs As New FileStream(path, FileMode.Open, FileAccess.Read)
                img = Image.FromStream(fs)
            End Using
        End If
        'Call Helper_Logger.writelog(Level.Debug, "Die Grafik wurde als " & path & " gespeichert.")
        Return img
    End Function
    Public Shared Function ScaleImage(ByVal OldImage As Image, ByVal TargetHeight As Integer, ByVal TargetWidth As Integer) As Image
        Dim NewHeight As Double = 0
        Dim NewWidth As Double = 0
        NewHeight = TargetHeight
        NewWidth = NewHeight / OldImage.Height * OldImage.Width
        If NewWidth > TargetWidth Then
            NewWidth = TargetWidth
            NewHeight = NewWidth / OldImage.Width * OldImage.Height
        End If
        Return New Bitmap(OldImage, Helper_VarConvert.ConvertToInteger(NewWidth), Helper_VarConvert.ConvertToInteger(NewHeight))
    End Function
    Public Shared Function ResizeImage1(ByVal InputImage As Image, ByVal NewWidth As Integer, ByVal NewHeight As Integer) As Image
        Return New Bitmap(InputImage, New Size(NewWidth, NewHeight))
    End Function
    Public Shared Function GetImageFromByteArray(ByVal InputStream As FileStream) As Image
        Dim image As Image = Nothing
        Dim imageConverter As New ImageConverter
        Dim fileData As Byte() = Nothing

        Try
            Using BinaryReader As BinaryReader = New BinaryReader(InputStream)
                fileData = BinaryReader.ReadBytes(Helper_VarConvert.ConvertToInteger(InputStream.Length))
            End Using
            imageConverter = New System.Drawing.ImageConverter()
            image = TryCast(imageConverter.ConvertFrom(fileData), System.Drawing.Image)

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return image
    End Function
    Public Shared Function GetImageFromByteArray(ByVal InputStream As MemoryStream) As Image
        Dim image As Image = Nothing
        Dim imageConverter As New ImageConverter
        Dim fileData As Byte() = Nothing

        Try
            Using BinaryReader As BinaryReader = New BinaryReader(InputStream)
                fileData = BinaryReader.ReadBytes(Helper_VarConvert.ConvertToInteger(InputStream.Length))
            End Using
            imageConverter = New System.Drawing.ImageConverter()
            image = TryCast(imageConverter.ConvertFrom(fileData), System.Drawing.Image)

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return image
    End Function

#End Region
#Region "ResizeImage"
    Public Overloads Shared Function ResizeImage(SourceImage As Image, TargetWidth As Int32, TargetHeight As Int32) As Bitmap
        Dim bmSource As Bitmap = New Bitmap(SourceImage)
        Return ResizeImage(bmSource, TargetWidth, TargetHeight)
    End Function
    Public Overloads Shared Function ResizeImage(bmSource As Bitmap, TargetWidth As Int32, TargetHeight As Int32) As Bitmap
        Dim bmDest As New Bitmap(TargetWidth, TargetHeight, PixelFormat.Format32bppArgb)

        Dim nSourceAspectRatio As Double = bmSource.Width / bmSource.Height
        Dim nDestAspectRatio As Double = bmDest.Width / bmDest.Height

        Dim NewX As Double = 0
        Dim NewY As Double = 0
        Dim NewWidth As Double = bmDest.Width
        Dim NewHeight As Double = bmDest.Height

        If nDestAspectRatio = nSourceAspectRatio Then
            'same ratio
        ElseIf nDestAspectRatio > nSourceAspectRatio Then
            'Source is taller
            NewWidth = System.Convert.ToInt32(Math.Floor(nSourceAspectRatio * NewHeight))
            NewX = System.Convert.ToInt32(Math.Floor((bmDest.Width - NewWidth) / 2))
        Else
            'Source is wider
            NewHeight = System.Convert.ToInt32(Math.Floor((1 / nSourceAspectRatio) * NewWidth))
            NewY = System.Convert.ToInt32(Math.Floor((bmDest.Height - NewHeight) / 2))
        End If

        Using grDest As Graphics = Graphics.FromImage(bmDest)
            With grDest
                .CompositingQuality = CompositingQuality.HighQuality
                .InterpolationMode = InterpolationMode.HighQualityBicubic
                .PixelOffsetMode = PixelOffsetMode.HighQuality
                .SmoothingMode = SmoothingMode.AntiAlias
                .CompositingMode = CompositingMode.SourceOver

                .DrawImage(CType(bmSource, Image), Helper_VarConvert.ConvertToInteger(NewX), Helper_VarConvert.ConvertToInteger(NewY), Helper_VarConvert.ConvertToInteger(NewWidth), Helper_VarConvert.ConvertToInteger(NewHeight))
            End With
        End Using

        Return bmDest
    End Function
#End Region
#Region "Json"
    Public Shared Function checkJasonValue(Value As JToken) As Object
        Dim RetVal As New Object
        Try
            'a = Value.Type
            If Value Is Nothing Then
                RetVal = ""
            Else
                Select Case Value.Type
                    Case JTokenType.Boolean
                        RetVal = CBool(Value)
                    Case JTokenType.Date
                        RetVal = CDate(Value)
                    Case JTokenType.Float
                        RetVal = (Value)
                    Case JTokenType.Integer
                        RetVal = Helper_VarConvert.ConvertToInteger(Value)
                    Case JTokenType.String
                        RetVal = CStr(Value)
                    Case JTokenType.None
                        RetVal = ""
                    Case JTokenType.Null
                        RetVal = ""
                    Case Else
                        RetVal = ""
                End Select
            End If
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
            RetVal = ""
        End Try
        Return RetVal
    End Function
#End Region
#Region "Konvertierungen"
    Public Shared Function GetDoubleFromString(ByVal Value As String) As Double
        Dim match As Match = Regex.Match(Value, "\d+(?:[.,]\d+)*")
        If match.Success Then
            Return Helper_VarConvert.ConvertToDouble(match.Value)
        Else
            Return 0
        End If
    End Function
#End Region
#Region "Netzwerk"
    Public Shared Function DnsLookup(ByVal hostNameOrAddress As String) As List(Of IPAddress)
        Dim IpList As New List(Of IPAddress)
        Try
            Dim ipAddresses As IPAddress() = Dns.GetHostAddresses(hostNameOrAddress)
            'Dim hostEntry As IPHostEntry = Dns.GetHostEntry(hostNameOrAddress)
            'Dim ipAddresses As IPAddress() = hostEntry.AddressList
            For Each ip As IPAddress In ipAddresses
                IpList.Add(ip)
            Next
        Catch ex As Exception
            If ex.Message = "Der angegebene Host ist unbekannt" Then
                If Helper.IsIDE() Then Stop
            Else
                Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
                If Helper.IsIDE() Then Stop
            End If
        End Try
        Return IpList
    End Function
    Public Shared Function PingHost(ByVal HostName As String) As Boolean
        Dim IpList As New List(Of IPAddress)
        Dim myPing As New Ping()
        Dim buffer As Byte() = New Byte(31) {}
        Dim timeout As Integer = 1000
        Dim pingOptions As New PingOptions()
        Dim reply As PingReply
        Try
            IpList = DnsLookup(HostName)
            For Each ip As IPAddress In IpList
                For i As Integer = 1 To 10
                    reply = myPing.Send(ip, timeout, buffer, pingOptions)
                    If reply.Status = IPStatus.Success Then
                        Return True
                    End If
                    wait(500)
                Next
            Next

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return False
    End Function
    Public Shared Function check4InternetConnect() As Boolean
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
                    wait(1000)
                Next
                If success Then
                    Exit For
                Else
                    Return False
                End If
            Next
            If reply.Status = IPStatus.Success Then
                Return True
            Else
                Return False
            End If
        Catch generatedExceptionName As Exception
            Return False
        End Try
    End Function
    Public Shared Sub FtpUploadFile(ByVal _FileName As String, ByVal _UploadPath As String, ByVal _FTPUser As String, ByVal _FTPPass As String)

        Try
            Dim _FileInfo As New FileInfo(_FileName)
            Dim _FtpWebRequest As FtpWebRequest = CType(FtpWebRequest.Create(New Uri(_UploadPath)), FtpWebRequest)
            _FtpWebRequest.Credentials = New NetworkCredential(_FTPUser, _FTPPass)
            _FtpWebRequest.KeepAlive = False
            _FtpWebRequest.Timeout = 20000
            _FtpWebRequest.Method = WebRequestMethods.Ftp.UploadFile
            _FtpWebRequest.UseBinary = True
            _FtpWebRequest.ContentLength = _FileInfo.Length
            Dim buffLength As Integer = 2048
            Dim buff(buffLength - 1) As Byte
            Using _FileStream As FileStream = _FileInfo.OpenRead()
                Using _Stream As Stream = _FtpWebRequest.GetRequestStream()
                    Dim contentLen As Integer = _FileStream.Read(buff, 0, buffLength)
                    Do While contentLen <> 0
                        _Stream.Write(buff, 0, contentLen)
                        contentLen = _FileStream.Read(buff, 0, buffLength)
                    Loop
                End Using
            End Using
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False)
            If Helper.IsIDE() Then Stop
        End Try
    End Sub
    Function GetIPAddress(ByVal CompName As String) As String

        Try
            Dim i1 As Integer
            Dim strIPAddress As String = Nothing
            Dim strIPAddress1 As String = ""
            For i1 = 0 To 1
                strIPAddress1 = ""
                strIPAddress1 = Dns.GetHostEntry(CompName).AddressList(i1).ToString()
                If Not IsNothing(strIPAddress1) Then
                    If IsValidIP(strIPAddress1) Then
                        strIPAddress = strIPAddress1
                        Exit For
                    End If
                End If
            Next i1
            If strIPAddress Is Nothing Then
                Return ""
            Else
                Return strIPAddress
            End If
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False)
            If Helper.IsIDE() Then Stop
            Return ""
        End Try
    End Function
    Public Shared Function GetMAC(IPAddresse As String) As String
        Try
            Dim ab As Byte() = New Byte(5) {}
            Dim len As Integer = ab.Length
            Dim ipa As IPAddress = IPAddress.Parse(IPAddresse)
            Dim ip_addr As UInteger = BitConverter.ToUInt32(ipa.GetAddressBytes, 0)
            'Dim result As Integer = SendARP(ip_addr, 0, ab, len)
            Dim mac As String = BitConverter.ToString(ab, 0, 6)
            Return mac
        Catch generatedExceptionName As Exception
            Return "MAC Not available"
        End Try
    End Function
    Public Shared Function IsValidIP(ByVal sIP As String) As Boolean
        Dim bValid As Boolean = False
        If sIP.Length > 0 Then
            ' IP in Blöcke aufteilen
            Dim sNumber() As String = sIP.Split(CType(".", Char))
            ' wenn 4 Blöcke enthalten...
            If sNumber.Length = 4 Then
                bValid = True
                ' ... prüfen, ob Block numerisch
                For i As Integer = 0 To 3
                    If Not IsNumeric(sNumber(i)) OrElse sNumber(i).Length > 3 Then
                        bValid = False
                        Exit For
                    End If
                    Dim Value As Integer = Integer.Parse(sNumber(i))
                    If i = 0 Then
                        ' 1. Block muss im Bereich 10 bis 255 liegen
                        If Value < 10 OrElse Value > 255 Then
                            bValid = False
                            Exit For
                        End If
                    Else
                        ' alle anderen Blöcke müssen zwischen 0 und 255 liegen
                        If Value < 0 OrElse Value > 255 Then
                            bValid = False
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        Return bValid
    End Function
    Public Shared Function UrlExists(ByVal url As String) As Boolean
        'url = "http: //www.schlaf-und-wohn.de"
        Dim uri As New Uri(url)
        url = uri.Scheme & "://" & uri.Host
        Dim exists As Boolean
        Try
            Dim Response As WebResponse = Nothing
            Dim WebReq As HttpWebRequest = CType(HttpWebRequest.Create(url), HttpWebRequest)
            Response = WebReq.GetResponse
            Response.Close()
            exists = True
        Catch ex As Exception
            exists = False
        End Try
        Return exists
    End Function
    Public Shared Function UrlExistsAlternative(ByVal url As String) As Boolean
        Dim exists As Boolean
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
        Dim response As System.Net.HttpWebResponse
        Try
            req = System.Net.WebRequest.Create(url)
            response = DirectCast(req.GetResponse(), System.Net.HttpWebResponse)
            exists = True
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return exists
    End Function
    Public Shared Sub WakeOnLAN(ByVal MACAddress As String)
        Dim Client As New UdpClient
        Dim Count As Integer = 0
        Dim ByteArray(1024) As Byte
        Dim Pos As Integer = 0
        Client.Connect("255.255.255.255", 12287)
        Client.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.Broadcast, 0)
        For i As Integer = 0 To 5
            Count += 1
            ByteArray(Count) = &HFF
        Next
        For i As Integer = 0 To 15
            Pos = 0
            For n As Integer = 0 To 5
                Count += 1
                ByteArray(Count) = Byte.Parse(MACAddress.Substring(Pos, 2), NumberStyles.HexNumber)
                Pos += 2
            Next
        Next
        Client.Send(ByteArray, 1024)
    End Sub
    Public Shared Function IsIP(ByVal IP As String) As Boolean
        Return System.Text.RegularExpressions.Regex.IsMatch(IP, "\b((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$\b")
    End Function
    Public Shared Function UrlIsValid(ByVal url As String) As Boolean
        'Dim is_valid As Boolean = False
        'If url.ToLower().StartsWith("www.") Then url = "http://" & url

        'Dim web_response As HttpWebResponse = Nothing
        'Try
        '    Dim web_request As WebRequest = HttpWebRequest.Create(url)
        '    web_response = DirectCast(web_request.GetResponse(), HttpWebResponse)
        '    Return True
        'Catch ex As Exception
        '    Return False
        'Finally
        '    If Not (web_response Is Nothing) Then web_response.Close()
        'End Try

        Dim pattern As String
        pattern = "http(s)?://([\w+?\.\w+])+([a-zA-Z0-9\~\!\@\#\$\%\^\&\*\(\)_\-\=\+\\\/\?\.\:\;\'\,]*)?"
        If Regex.IsMatch(url, pattern) Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region
#Region "Registry"
    Public Shared Function KeyExist(Hive As RegistryHive, RegPathLocal As String) As Boolean
        Dim regKey As RegistryKey = Nothing
        Try
            Select Case Hive
                Case RegistryHive.ClassesRoot
                    regKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal)
                Case RegistryHive.CurrentConfig
                    regKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal)
                Case RegistryHive.CurrentUser
                    regKey = Registry.CurrentUser.OpenSubKey(RegPathLocal)
                Case RegistryHive.LocalMachine
                    regKey = Registry.LocalMachine.OpenSubKey(RegPathLocal)
                Case RegistryHive.PerformanceData
                    regKey = Registry.PerformanceData.OpenSubKey(RegPathLocal)
                Case RegistryHive.Users
                    regKey = Registry.Users.OpenSubKey(RegPathLocal)
            End Select
            If regKey Is Nothing Then
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return False
    End Function
    Public Shared Function CopyRecursive(ByVal HiveOld As RegistryHive,
                                     ByVal RegPathOld As String,
                                     ByVal KeyNameOld As String,
                                     ByVal HiveNew As RegistryHive,
                                     ByVal RegPathNew As String,
                                     ByVal KeyNameNew As String) As Boolean
        Dim parentKeyOld As RegistryKey = Nothing
        Dim parentKeyNew As RegistryKey = Nothing
        Select Case HiveOld
            Case RegistryHive.ClassesRoot
                parentKeyOld = Registry.ClassesRoot.OpenSubKey(RegPathOld, True)
            Case RegistryHive.CurrentConfig
                parentKeyOld = Registry.CurrentConfig.OpenSubKey(RegPathOld, True)
            Case RegistryHive.CurrentUser
                parentKeyOld = Registry.CurrentUser.OpenSubKey(RegPathOld, True)
            Case RegistryHive.LocalMachine
                parentKeyOld = Registry.LocalMachine.OpenSubKey(RegPathOld, True)
            Case RegistryHive.PerformanceData
                parentKeyOld = Registry.PerformanceData.OpenSubKey(RegPathOld, True)
            Case RegistryHive.Users
                parentKeyOld = Registry.Users.OpenSubKey(RegPathOld, True)
        End Select
        Select Case HiveNew
            Case RegistryHive.ClassesRoot
                parentKeyNew = Registry.ClassesRoot.OpenSubKey(RegPathNew, True)
            Case RegistryHive.CurrentConfig
                parentKeyNew = Registry.CurrentConfig.OpenSubKey(RegPathNew, True)
            Case RegistryHive.CurrentUser
                parentKeyNew = Registry.CurrentUser.OpenSubKey(RegPathNew, True)
            Case RegistryHive.LocalMachine
                parentKeyNew = Registry.LocalMachine.OpenSubKey(RegPathNew, True)
            Case RegistryHive.PerformanceData
                parentKeyNew = Registry.PerformanceData.OpenSubKey(RegPathNew, True)
            Case RegistryHive.Users
                parentKeyNew = Registry.Users.OpenSubKey(RegPathNew, True)
        End Select
        Return CopyKey(parentKeyOld, KeyNameOld, parentKeyNew, KeyNameNew)

    End Function
    Public Shared Function CopyKey(ByVal parentKeyOld As RegistryKey, SubKeyNameOld As String, ByVal parentKeyNew As RegistryKey, SubKeyNameNew As String) As Boolean
        Dim destinationKey As RegistryKey = parentKeyNew.CreateSubKey(SubKeyNameNew)
        Dim sourceKey As RegistryKey = parentKeyOld.OpenSubKey(SubKeyNameOld, True)
        RecurseCopyKey(sourceKey, destinationKey)
        Return True
    End Function
    Public Shared Function CopyKey(ByVal parentKey As RegistryKey, SubKeyNameOld As String, SubKeyNameNew As String) As Boolean
        Dim destinationKey As RegistryKey = parentKey.CreateSubKey(SubKeyNameNew)
        Dim sourceKey As RegistryKey = parentKey.OpenSubKey(SubKeyNameOld, True)
        RecurseCopyKey(sourceKey, destinationKey)
        Return True
    End Function
    Public Shared Sub RecurseCopyKey(sourceKey As RegistryKey, destinationKey As RegistryKey)
        For Each valueName As String In sourceKey.GetValueNames()
            Dim objValue As Object = sourceKey.GetValue(valueName)
            Dim valKind As RegistryValueKind = sourceKey.GetValueKind(valueName)
            destinationKey.SetValue(valueName, objValue, valKind)
        Next
        For Each sourceSubKeyName As String In sourceKey.GetSubKeyNames()
            Dim sourceSubKey As RegistryKey = sourceKey.OpenSubKey(sourceSubKeyName)
            Dim destSubKey As RegistryKey = destinationKey.CreateSubKey(sourceSubKeyName)
            RecurseCopyKey(sourceSubKey, destSubKey)
        Next
    End Sub
    Public Shared Function RegistryCreateKey(Hive As RegistryHive, RegPathLocal As String) As RegistryKey
        Dim regKey As RegistryKey = Nothing
        Try
            Select Case Hive
                Case RegistryHive.ClassesRoot
                    If Not KeyExist(Hive, RegPathLocal) Then
                        regKey = Registry.ClassesRoot.CreateSubKey(RegPathLocal)
                    Else
                        regKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal)
                    End If
                Case RegistryHive.CurrentConfig
                    If Not KeyExist(Hive, RegPathLocal) Then
                        regKey = Registry.CurrentConfig.CreateSubKey(RegPathLocal)
                    Else
                        regKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal)
                    End If
                Case RegistryHive.CurrentUser
                    If Not KeyExist(Hive, RegPathLocal) Then
                        regKey = Registry.CurrentUser.CreateSubKey(RegPathLocal)
                    Else
                        regKey = Registry.CurrentUser.OpenSubKey(RegPathLocal)
                    End If
                Case RegistryHive.LocalMachine
                    If Not KeyExist(Hive, RegPathLocal) Then
                        regKey = Registry.LocalMachine.CreateSubKey(RegPathLocal)
                    Else
                        regKey = Registry.LocalMachine.OpenSubKey(RegPathLocal)
                    End If
                Case RegistryHive.PerformanceData
                    If Not KeyExist(Hive, RegPathLocal) Then
                        regKey = Registry.PerformanceData.CreateSubKey(RegPathLocal)
                    Else
                        regKey = Registry.PerformanceData.OpenSubKey(RegPathLocal)
                    End If
                Case RegistryHive.Users
                    If Not KeyExist(Hive, RegPathLocal) Then
                        regKey = Registry.Users.CreateSubKey(RegPathLocal)
                    Else
                        regKey = Registry.Users.OpenSubKey(RegPathLocal)
                    End If
            End Select

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return regKey
    End Function
    Public Shared Function RegistryCreateSubKey(Hive As RegistryHive, RegPathLocal As String) As RegistryKey
        Dim regKey As RegistryKey = Nothing
        Select Case Hive
            Case RegistryHive.ClassesRoot
                If Not KeyExist(Hive, RegPathLocal) Then
                    regKey = Registry.ClassesRoot.CreateSubKey(RegPathLocal)
                End If
            Case RegistryHive.CurrentConfig
                If Not KeyExist(Hive, RegPathLocal) Then
                    regKey = Registry.CurrentConfig.CreateSubKey(RegPathLocal)
                End If
            Case RegistryHive.CurrentUser
                If Not KeyExist(Hive, RegPathLocal) Then
                    regKey = Registry.CurrentUser.CreateSubKey(RegPathLocal)
                End If
            Case RegistryHive.LocalMachine
                If Not KeyExist(Hive, RegPathLocal) Then
                    regKey = Registry.LocalMachine.CreateSubKey(RegPathLocal)
                End If
            Case RegistryHive.PerformanceData
                If Not KeyExist(Hive, RegPathLocal) Then
                    regKey = Registry.PerformanceData.CreateSubKey(RegPathLocal)
                End If
            Case RegistryHive.Users
                If Not KeyExist(Hive, RegPathLocal) Then
                    regKey = Registry.Users.CreateSubKey(RegPathLocal)
                End If
        End Select
        Return regKey
    End Function
    Public Shared Sub RegistryDeleteKey(ByVal Hive As RegistryHive, ByVal RegPathLocal As String, ByVal SubKey2Delete As String)
        RegPathLocal = RegPathLocal.Trim & "\" & SubKey2Delete.Trim
        Select Case Hive
            Case RegistryHive.ClassesRoot
                If KeyExist(Hive, RegPathLocal) Then
                    Registry.ClassesRoot.DeleteSubKey(RegPathLocal, False)
                End If
            Case RegistryHive.CurrentConfig
                If KeyExist(Hive, RegPathLocal) Then
                    Registry.CurrentConfig.DeleteSubKey(RegPathLocal, False)
                End If
            Case RegistryHive.CurrentUser
                If KeyExist(Hive, RegPathLocal) Then
                    Registry.CurrentUser.DeleteSubKey(RegPathLocal, False)
                End If
            Case RegistryHive.LocalMachine
                If KeyExist(Hive, RegPathLocal) Then
                    Registry.LocalMachine.DeleteSubKey(RegPathLocal, False)
                End If
            Case RegistryHive.PerformanceData
                If KeyExist(Hive, RegPathLocal) Then
                    Registry.PerformanceData.DeleteSubKey(RegPathLocal, False)
                End If
            Case RegistryHive.Users
                If KeyExist(Hive, RegPathLocal) Then
                    Registry.Users.DeleteSubKey(RegPathLocal, False)
                End If
        End Select
    End Sub
    Public Shared Function RegistryDeleteValue(Hive As RegistryHive, RegPathLocal As String, Key As String) As Boolean
        Dim regKey As RegistryKey = Nothing
        Select Case Hive
            Case RegistryHive.ClassesRoot
                regKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.CurrentConfig
                regKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.CurrentUser
                regKey = Registry.CurrentUser.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.LocalMachine
                regKey = Registry.LocalMachine.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.PerformanceData
                regKey = Registry.PerformanceData.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.Users
                regKey = Registry.Users.OpenSubKey(RegPathLocal, True)
        End Select
        If regKey IsNot Nothing Then
            regKey.DeleteValue(Key)
        End If
        Return CType(regKey.GetValue(Key, True), Boolean)
    End Function
    Public Shared Function RegistryReadSubKeys(ByVal Hive As RegistryHive, ByVal RegPathLocal As String, ByVal StartsWithString As String) As List(Of String)
        Dim regKey As RegistryKey = Nothing
        Dim SubKeyList As New List(Of String)
        Select Case Hive
            Case RegistryHive.ClassesRoot
                regKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.CurrentConfig
                regKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.CurrentUser
                regKey = Registry.CurrentUser.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.LocalMachine
                regKey = Registry.LocalMachine.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.PerformanceData
                regKey = Registry.PerformanceData.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.Users
                regKey = Registry.Users.OpenSubKey(RegPathLocal, True)
        End Select
        Try
            For Each Value As String In regKey.GetSubKeyNames().ToList
                If Value.ToUpper.StartsWith(StartsWithString.ToUpper) Then
                    SubKeyList.Add(Value)
                End If
            Next
        Catch ex As Exception
            SubKeyList = Nothing
        End Try
        Return SubKeyList
    End Function

    Friend Shared ReadOnly IniPathFile As String = Path.Combine(Helper.ValPath(Application.StartupPath, True), "prostruktura.ini")
    Dim IniFile1 As IniFile = New IniFile(IniPathFile)


    'Read String Public String ReadString(String section, String key)
    'Read Boolean: Public bool ReadBoolean(String section, String key)
    'Read Decimal: Public Decimal ReadDecimal(String section, String key)
    'Read Double: Public Double ReadDouble(String section, String key)
    'Read Float(Single) : Public float ReadFloat(String section, String key)
    'Read Integer: Public int ReadInteger(String section, String key)
    'Read DateTime : Public DateTime ReadDateTime(String section, String key)
    'Read Date: Public DateTime ReadDate(String section, String key)

    Friend Function CheckIni() As Boolean
        If Not File.Exists(IniPathFile) Then
            Return False
        Else
            Dim IniFile2 As IniFile = New IniFile(IniPathFile)
            If Not IniFile2.KeyExists("General", "Test") Then
                IniFile2.WriteString("General", "Test", Application.ProductVersion)
                If Application.ProductVersion = IniFile2.ReadString("General", "Test") Then
                    IniFile2.DeleteKey("General", "Test")
                    Return True
                Else
                    Return False
                End If
            End If
        End If
        Return False

    End Function


    Public Shared Function SelectReadSettingsFrom(ByVal IniSection As String,
                                                  ByVal RegPathLocal As String,
                                                  ByVal IniKey As String,
                                                  ByVal Optional VarType As Helper.VarTypeEnu = Helper.VarTypeEnu.Type_String,
                                                  ByVal Optional AllowNothing As Boolean = False,
                                                  ByVal Optional DefaultValue As Object = Nothing) As Object

        Try

            Return IniReadValue(IniSection,
                                IniKey,
                                VarType,
                                AllowNothing,
                                DefaultValue)

        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Thread.CurrentThread.ManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return False

    End Function
    Public Shared Function SelectReadSettingsFrom(ByVal Hive As RegistryHive,
                                                  ByVal RegPathLocal As String,
                                                  ByVal QueryKey As String,
                                                  ByVal Optional VarType As Helper.VarTypeEnu = Helper.VarTypeEnu.Type_String,
                                                  ByVal Optional AllowNothing As Boolean = False,
                                                  ByVal Optional DefaultValue As Object = Nothing) As Object
        Try
            Return RegistryReadValueDirect(Hive,
                                     RegPathLocal,
                                     QueryKey,
                                     VarType,
                                     AllowNothing,
                                     DefaultValue)


        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Thread.CurrentThread.ManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return False
    End Function

    Friend Shared Function IniReadValue(ByVal IniSection As String,
                                 ByVal IniKey As String,
                                 ByVal Optional VarType As Helper.VarTypeEnu = Helper.VarTypeEnu.Type_String,
                                 ByVal Optional AllowNothing As Boolean = False,
                                 ByVal Optional DefaultValue As Object = Nothing) As Object

        Dim IniFile1 As IniFile = New IniFile(IniPathFile)
        Dim IniValueString As String = ""
        Dim IniValueBoolean As Boolean = False
        Dim IniValueDouble As Double = 0
        Dim IniValueinteger As Integer = 0
        Dim IniValueDate As Date
        Dim IniValueDateTime As DateTime
        Try
            Select Case VarType
                Case Helper.VarTypeEnu.Type_Boolean
                    IniValueBoolean = IniFile1.ReadBoolean(IniSection, IniKey)
                    Return IniValueBoolean
                Case Helper.VarTypeEnu.Type_Double
                    IniValueDouble = IniFile1.ReadDouble(IniSection, IniKey)
                    Return IniValueDouble
                Case Helper.VarTypeEnu.Type_Integer
                    IniValueinteger = IniFile1.ReadInteger(IniSection, IniKey)
                    Return IniValueinteger
                Case Helper.VarTypeEnu.Type_String
                    IniValueString = IniFile1.ReadString(IniSection, IniKey)
                    If Not String.IsNullOrWhiteSpace(IniValueString) Then
                        If DefaultValue Is Nothing Then
                            IniValueString = ""
                        Else
                            IniValueString = CStr(DefaultValue)
                        End If
                    End If
                    Return IniValueString
                Case Helper.VarTypeEnu.Type_DateAsString
                    IniValueDate = IniFile1.ReadDate(IniSection, IniKey)
                    Return CStr(IniValueDate)
                Case Helper.VarTypeEnu.Type_Object
                    RetVal = Nothing
                Case Helper.VarTypeEnu.Type_Date
                    IniValueDate = IniFile1.ReadDate(IniSection, IniKey)
                    Return IniValueDate
                Case Helper.VarTypeEnu.Type_DateTime
                    IniValueDateTime = IniFile1.ReadDateTime(IniSection, IniKey)
                    Return IniValueDateTime
                Case Else
                    If DefaultValue Is Nothing Then
                        RetVal = ""
                    Else
                        RetVal = CStr(DefaultValue)
                    End If
            End Select

        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Thread.CurrentThread.ManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return "Error"
    End Function
    Public Shared Function RegistryReadValueDirect(ByVal Hive As RegistryHive,
                                         ByVal RegPathLocal As String,
                                         ByVal QueryKey As String,
                                         ByVal Optional VarType As Helper.VarTypeEnu = Helper.VarTypeEnu.Type_String,
                                         ByVal Optional AllowNothing As Boolean = False,
                                         ByVal Optional DefaultValue As Object = Nothing) As Object
        Dim regKey As RegistryKey = Nothing
        Dim rd As Double = 0
        Dim rs As String = ""
        Dim value As New Object
        Dim RetVal As New Object
        Try
            Select Case Hive
                Case RegistryHive.ClassesRoot
                    regKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.CurrentConfig
                    regKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.CurrentUser
                    regKey = Registry.CurrentUser.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.LocalMachine
                    regKey = Registry.LocalMachine.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.PerformanceData
                    regKey = Registry.PerformanceData.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.Users
                    regKey = Registry.Users.OpenSubKey(RegPathLocal, True)
            End Select
            If regKey Is Nothing Then
                value = Nothing
            Else
                value = regKey.GetValue(QueryKey)
            End If
            If value Is Nothing AndAlso AllowNothing Then
                RetVal = Nothing
            ElseIf value Is Nothing AndAlso Not AllowNothing Then
                Select Case VarType
                    Case Helper.VarTypeEnu.Type_Boolean
                        If DefaultValue Is Nothing Then
                            RetVal = False
                        Else
                            RetVal = CBool(DefaultValue)
                        End If
                    Case Helper.VarTypeEnu.Type_Double
                        If DefaultValue Is Nothing Then
                            RetVal = 0
                        Else
                            RetVal = Helper_VarConvert.ConvertToDouble(DefaultValue)
                        End If
                    Case Helper.VarTypeEnu.Type_Integer
                        If DefaultValue Is Nothing Then
                            RetVal = 0
                        Else
                            RetVal = Helper_VarConvert.ConvertToInteger(DefaultValue)
                        End If
                    Case Helper.VarTypeEnu.Type_String
                        If DefaultValue Is Nothing Then
                            RetVal = ""
                        Else
                            RetVal = CStr(DefaultValue)
                        End If
                    Case Helper.VarTypeEnu.Type_DateAsString
                        If DefaultValue Is Nothing Then
                            RetVal = ""
                        Else
                            RetVal = CStr(DefaultValue)
                        End If
                    Case Helper.VarTypeEnu.Type_Object
                        RetVal = Nothing
                    Case Helper.VarTypeEnu.Type_Date
                        If DefaultValue Is Nothing Then
                            Dim DOB As Nullable(Of Date) = CDate(Now)
                            Dim datestring As String = DOB.Value.ToString("d")
                            DOB = Nothing
                            RetVal = datestring
                        Else
                            RetVal = CDate(DefaultValue)
                        End If
                    Case Else
                        If DefaultValue Is Nothing Then
                            RetVal = ""
                        Else
                            RetVal = CStr(DefaultValue)
                        End If
                End Select
            Else
                Select Case VarType
                    Case Helper.VarTypeEnu.Type_Boolean
                        RetVal = CBool(value)
                    Case Helper.VarTypeEnu.Type_Double
                        RetVal = Helper_VarConvert.ConvertToDouble(Val(value))
                    Case Helper.VarTypeEnu.Type_Integer
                        RetVal = Helper_VarConvert.ConvertToInteger(Val(value))
                    Case Helper.VarTypeEnu.Type_String
                        RetVal = value.ToString
                    Case Helper.VarTypeEnu.Type_DateAsString
                        RetVal = value.ToString
                    Case Helper.VarTypeEnu.Type_Date
                        RetVal = CDate(value)
                    Case Helper.VarTypeEnu.Type_Object
                        RetVal = Nothing
                    Case Else
                        RetVal = value.ToString
                End Select
            End If

        Catch ex As Exception
            'Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            'If Helper.IsIDE() Then Stop
            Return ""
        End Try
        Return RetVal
    End Function

    Public Shared Function RegistryReadValuePattern(Hive As RegistryHive, RegPathLocal As String, KeyPattern As String, Optional VariType As Helper.VarTypeEnu = Helper.VarTypeEnu.Type_String) As List(Of RegFindResult)
        Dim regKey As RegistryKey = Nothing
        Dim rd As Double = 0
        Dim rs As String = ""
        'Dim value As New Object
        Dim RetVal As New List(Of RegFindResult)
        Try
            Select Case Hive
                Case RegistryHive.ClassesRoot
                    regKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.CurrentConfig
                    regKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.CurrentUser
                    regKey = Registry.CurrentUser.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.LocalMachine
                    regKey = Registry.LocalMachine.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.PerformanceData
                    regKey = Registry.PerformanceData.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.Users
                    regKey = Registry.Users.OpenSubKey(RegPathLocal, True)
            End Select
            If regKey IsNot Nothing Then
                If regKey.ValueCount > 0 Then
                    For Each key As String In regKey.GetValueNames
                        If key Like KeyPattern Then
                            Dim RetValTmp As New RegFindResult
                            RetValTmp.Hive = Hive
                            RetValTmp.RegPath = RegPathLocal
                            RetValTmp.Key = key
                            RetValTmp.VarType = VariType
                            RetValTmp.Value = SelectReadSettingsFrom(Hive, RegPathLocal, key, VariType)
                            RetVal.Add(RetValTmp)
                        End If
                    Next
                End If
            End If
            'regKey.Close()
            'regKey.Dispose()
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return RetVal
    End Function

    Public Shared Function RegistryWriteValue(ByVal Hive As RegistryHive,
                                              ByVal RegPathLocal As String,
                                              ByVal Key As String,
                                              ByVal Value As Object) As Boolean
        Dim regKey As RegistryKey = Nothing
        Try
            '\SOFTWARE\TimeShareIT\ConDrop\TestApplication\OttoDll\Parameter
            '\SOFTWARE\TimeShareIT\ConDrop\TestApplication\OttoDll\Parameter
            If Not IsNothing(Value) Then
                If IsNothing(Hive) OrElse Hive = 0 Then
                    'Stop
                Else
                    regKey = RegistryCreateKey(Hive, RegPathLocal)
                    Select Case Hive
                        Case RegistryHive.ClassesRoot
                            regKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal, True)
                        Case RegistryHive.CurrentConfig
                            regKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal, True)
                        Case RegistryHive.CurrentUser
                            regKey = Registry.CurrentUser.OpenSubKey(RegPathLocal, True)
                        Case RegistryHive.LocalMachine
                            regKey = Registry.LocalMachine.OpenSubKey(RegPathLocal, True)
                        Case RegistryHive.PerformanceData
                            regKey = Registry.PerformanceData.OpenSubKey(RegPathLocal, True)
                        Case RegistryHive.Users
                            regKey = Registry.Users.OpenSubKey(RegPathLocal, True)
                    End Select
                    Registry.SetValue(regKey.Name, Key, Value)
                    'regKey.Close()
                    'regKey.Dispose()
                End If
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
            Return False
        End Try
        Return True
    End Function

    Public Shared Function RenameSubKey(ByVal Hive As RegistryHive, ByVal RegPathLocal As String, ByVal SubKeyNameOld As String, ByVal SubKeyNameNew As String) As Boolean
        Try
            Dim parentKey As RegistryKey = Nothing
            Select Case Hive
                Case RegistryHive.ClassesRoot
                    parentKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.CurrentConfig
                    parentKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.CurrentUser
                    parentKey = Registry.CurrentUser.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.LocalMachine
                    parentKey = Registry.LocalMachine.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.PerformanceData
                    parentKey = Registry.PerformanceData.OpenSubKey(RegPathLocal, True)
                Case RegistryHive.Users
                    parentKey = Registry.Users.OpenSubKey(RegPathLocal, True)
            End Select
            Dim rb As Boolean = CopyKey(parentKey, SubKeyNameOld, SubKeyNameNew)
            If rb Then
                parentKey.DeleteSubKeyTree(SubKeyNameOld)
            Else
                rb = False
            End If
            Return rb

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
            Return False
        End Try
    End Function
    Public Shared Function RegistryKeyExist(ByVal Hive As RegistryHive, ByVal RegPathLocal As String, ByVal RegKey As String) As Boolean
        Dim CheckRegKey As RegistryKey = Nothing
        'RegPath = RegPath & RegKey
        Select Case Hive
            Case RegistryHive.ClassesRoot
                CheckRegKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.CurrentConfig
                CheckRegKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.CurrentUser
                CheckRegKey = Registry.CurrentUser.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.LocalMachine
                CheckRegKey = Registry.LocalMachine.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.PerformanceData
                CheckRegKey = Registry.PerformanceData.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.Users
                CheckRegKey = Registry.Users.OpenSubKey(RegPathLocal, True)
        End Select

        If CheckRegKey Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Shared Function RegistryKeyExist(Hive As RegistryHive, RegPathLocal As String) As Boolean
        Dim regKey As RegistryKey = Nothing
        Select Case Hive
            Case RegistryHive.ClassesRoot
                regKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.CurrentConfig
                regKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.CurrentUser
                regKey = Registry.CurrentUser.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.LocalMachine
                regKey = Registry.LocalMachine.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.PerformanceData
                regKey = Registry.PerformanceData.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.Users
                regKey = Registry.Users.OpenSubKey(RegPathLocal, True)
        End Select

        If regKey Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function
    Public Shared Function RegistrySearchSubKeys(ByVal Path As String, ByVal SearchStr As String) As List(Of String)
        Dim Result As New List(Of String)
        Dim ParentKey As RegistryKey = Registry.LocalMachine.OpenSubKey(Path, True)
        For Each valueName As String In ParentKey.GetValueNames()
            'Dim CurStr As String = ParentKey.GetValue(valueName).ToString
            '? umbauen auf Regex:
            Dim CurStr As String = valueName & " = " & ParentKey.GetValue(valueName).ToString

            If valueName Like SearchStr Then
                'MsgBox(CurStr)
                Result.Add(valueName)
                valueName = ""
            End If
        Next
        If ParentKey.SubKeyCount > 0 Then
            For Each subKeyName As String In ParentKey.GetSubKeyNames()
                Dim Thispath As String = Path & subKeyName & "\\"
                RegistrySearchSubKeys(Thispath, SearchStr)
            Next
        End If
        Return Result
    End Function
    Public Shared Function RegistryListValues(ByVal Hive As RegistryHive, ByVal RegPathLocal As String, ByVal Optional Pattern As String = "") As List(Of String)
        Dim regKey As RegistryKey = Nothing
        Dim rd As Double = 0
        Dim rs As String = ""
        Dim value As New List(Of String)
        Dim count As Integer = 0

        Select Case Hive
            Case RegistryHive.ClassesRoot
                regKey = Registry.ClassesRoot.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.CurrentConfig
                regKey = Registry.CurrentConfig.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.CurrentUser
                regKey = Registry.CurrentUser.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.LocalMachine
                regKey = Registry.LocalMachine.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.PerformanceData
                regKey = Registry.PerformanceData.OpenSubKey(RegPathLocal, True)
            Case RegistryHive.Users
                regKey = Registry.Users.OpenSubKey(RegPathLocal, True)
        End Select
        Try
            count = regKey.ValueCount
            For Each valueName As String In regKey.GetValueNames()
                value.Add(valueName)
            Next
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return value
    End Function
    Public Shared Sub CopyRegistryFrom32BitTo64Bit(Hive As RegistryHive, RegPathLocal As String)
        Try
            'Call Helper_Logger.writelog(Level.Info, "Die 32-Bit Registry-Einträge werden nach 64-Bit kopiert.")
            Dim rb As Boolean = Helper.RegistryWriteValue(Hive, RegPathLocal, "Test", "Test")
            Dim SourceView32 As RegistryKey = RegistryKey.OpenBaseKey(Hive, RegistryView.Registry32)
            Dim SourceStartKey As RegistryKey = SourceView32.OpenSubKey(RegPathLocal, False)

            Dim DestinationView64 As RegistryKey = RegistryKey.OpenBaseKey(Hive, RegistryView.Registry64)
            Dim DestinationStartKey As RegistryKey = DestinationView64.OpenSubKey(RegPathLocal, True)

            rb = Helper.RegistryWriteValue(Hive, RegPathLocal, "Test", "Test")
            Call RecurseCopyKey(SourceStartKey, DestinationStartKey)

            'Call Helper_Logger.writelog(Level.Info, "Die 32-Bit Registry-Einträge wurden erfolgreich nach 64-Bit kopiert.")
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
    End Sub
    Public Shared Sub CheckRegistryKey(ByVal RegHive As RegistryHive,
                                   ByVal RegPathLocal As String,
                                   Key As String,
                                   ByVal VarType As VarTypeEnu,
                                   ByVal DefaultValue As Object)

        Select Case VarType
            Case Helper.VarTypeEnu.Type_Boolean
                If Helper.SelectReadSettingsFrom(RegHive, RegPathLocal, Key, VarType, True) Is Nothing Then
                    Call Helper.RegistryWriteValue(RegHive, RegPathLocal, Key, DefaultValue)
                End If
            Case Helper.VarTypeEnu.Type_Double
                If Helper.SelectReadSettingsFrom(RegHive, RegPathLocal, Key, VarType, True) Is Nothing Then
                    Call Helper.RegistryWriteValue(RegHive, RegPathLocal, Key, DefaultValue)
                End If
            Case Helper.VarTypeEnu.Type_Integer
                If Helper.SelectReadSettingsFrom(RegHive, RegPathLocal, Key, VarType, True) Is Nothing Then
                    Call Helper.RegistryWriteValue(RegHive, RegPathLocal, Key, DefaultValue)
                End If
            Case Helper.VarTypeEnu.Type_String
                If Helper.SelectReadSettingsFrom(RegHive, RegPathLocal, Key, VarType, True) Is Nothing Then
                    Call Helper.RegistryWriteValue(RegHive, RegPathLocal, Key, DefaultValue)
                End If
            Case Helper.VarTypeEnu.Type_DateAsString
                If Helper.SelectReadSettingsFrom(RegHive, RegPathLocal, Key, VarType, True) Is Nothing Then
                    Call Helper.RegistryWriteValue(RegHive, RegPathLocal, Key, DefaultValue)
                End If
            Case Helper.VarTypeEnu.Type_Date
                Dim RetVal As Date = Helper_VarConvert.ConvertToDate(Helper.SelectReadSettingsFrom(RegHive, RegPathLocal, Key, VarType, True))
                'If RetVal Is Nothing Then
                '    Call Helper.RegistryWriteValue(RegHive, RegPathLocal, Key, DefaultValue)
                'End If
            Case Else
                Dim RetVal As String = Helper_VarConvert.ConvertToString(Helper.SelectReadSettingsFrom(RegHive, RegPathLocal, Key, VarType, True), False, "")
                If String.IsNullOrWhiteSpace(RetVal) Then
                    Call Helper.RegistryWriteValue(RegHive, RegPathLocal, Key, DefaultValue)
                End If
        End Select



    End Sub

#End Region
#Region "Security"
    Public Shared Function GenerateSHA512String(ByVal inputString As String) As String
        Dim sha512 As SHA512 = SHA512Managed.Create()
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(inputString)
        Dim hash As Byte() = sha512.ComputeHash(bytes)
        Dim stringBuilder As New StringBuilder()
        For i As Integer = 0 To hash.Length - 1
            stringBuilder.Append(hash(i).ToString("X2"))
        Next
        Return stringBuilder.ToString()
    End Function
    Public Shared Function HashPassword(ByVal Password As String, ByVal Salt As String) As String
        Dim pwd As String = Salt & Password & Salt
        Dim hasher As New SHA256Managed()
        Dim pwdb As Byte() = Encoding.UTF8.GetBytes(pwd)
        Dim pwdh As Byte() = hasher.ComputeHash(pwdb)
        Return (BitConverter.ToString(pwdh).Replace("-", ""))
    End Function
#End Region
#Region "SOAP"
    Public Shared Sub print_soap(request As Object)
        Dim serxml As New XmlSerializer(request.[GetType]())
        Dim ms As New MemoryStream()
        serxml.Serialize(ms, request)
        'Dim xml As String = Encoding.UTF8.GetString(ms.ToArray())
        'Debug.Print(xml)
    End Sub
#End Region
#Region "Strings"
    Public Shared Function AdjustString(Input As String) As String
        If Input Is Nothing OrElse Input.Trim = "" Then
            Input = ""
            Return Input
        End If
        Input = Input.Trim
        Input = Replace(Input, "&amp;", "&")
        Input = Replace(Input, "&AMP;", "&")
        Input = Replace(Input, "*", "")
        Input = Replace(Input, "?", "")
        Input = Replace(Input, "'", "")
        Input = Replace(Input, "`", "")
        Input = Replace(Input, """", "")
        Input = Replace(Input, "<", "")
        Input = Replace(Input, ">", "")
        Input = Replace(Input, "|", "")
        Input = Replace(Input, Chr(34), "")
        Input = Replace(Input, "\", " ")
        If Input.Replace(Tab, "".Trim) = "" Then
            Input = ""
        End If
        Return Input
    End Function
    Private Shared Function Correct_Syntax(strText As String) As String
        strText = Replace(strText, "%", "%25")
        strText = Replace(strText, vbCr, "%0A")
        strText = Replace(strText, Environment.NewLine, "%0A")
        strText = Replace(strText, Environment.NewLine, "%0A")
        strText = Replace(strText, " ", "%20")
        strText = Replace(strText, "!", "%21")
        strText = Replace(strText, "#", "%23")
        strText = Replace(strText, "*", "%2A")
        strText = Replace(strText, "/", "%2F")
        strText = Replace(strText, "?", "%3F")
        strText = Replace(strText, "Ä", "%C4")
        strText = Replace(strText, "Ö", "%D6")
        strText = Replace(strText, "Ü", "%DC")
        strText = Replace(strText, "ß", "%DF")
        strText = Replace(strText, "ä", "%E4")
        strText = Replace(strText, "ö", "%F6")
        strText = Replace(strText, "ü", "%FC")
        Return strText
    End Function
    Public Shared Function ReplacePatternInString(ByVal Text As String, ByVal pattern As String, ByVal replacement As String) As String
        Dim rgx As New Regex(pattern)
        Text = rgx.Replace(Text, replacement)
        Return Text
    End Function
    Public Shared Function MeasureString(ByVal Text As String, ByVal FontName As String, ByVal FontSize As Single) As SizeF
        Dim Bitmap As Bitmap
        Dim Graphic As Graphics
        Dim Font As New Font(FontName, FontSize)
        Dim RetVal As New SizeF
        Bitmap = New Bitmap(1, 1)
        Graphic = Graphics.FromImage(Bitmap)
        Graphic.PageUnit = GraphicsUnit.Pixel
        RetVal = Graphic.MeasureString(Text, Font)
        Graphic.Dispose()
        Bitmap.Dispose()
        Return RetVal
    End Function
    Public Shared Function RemoveHTML(ByVal Text As String) As String
        'Return System.Text.RegularExpressions.Regex.Replace(Text, "<.*?>", " ")
        Return Regex.Replace(Text, "<[\S\s]*?>", " ")
    End Function
    Public Shared Function UnicodeToAscii(ByVal unicodeString As String) As String
        Dim ascii As Encoding = Encoding.ASCII
        Dim unicode As Encoding = Encoding.Unicode
        ' Convert the string into a byte array. 
        Dim unicodeBytes As Byte() = unicode.GetBytes(unicodeString)

        ' Perform the conversion from one encoding to the other. 
        Dim asciiBytes As Byte() = Encoding.Convert(unicode, ascii, unicodeBytes)

        ' Convert the new byte array into a char array and then into a string. 
        Dim asciiChars(ascii.GetCharCount(asciiBytes, 0, asciiBytes.Length) - 1) As Char
        ascii.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0)
        Dim asciiString As New String(asciiChars)
        Return asciiString
    End Function
    'Public Shared Function ContainsSimilarString(ByVal WordOrSentence As String, ByVal String2Search As String, ByVal MinimalSimilarity As Double) As Boolean
    '    Dim RetVal As Boolean = False
    '    Dim inparray() As String = WordOrSentence.Split(CChar(" "))
    '    For Each str As String In inparray
    '        Dim ls As SimMetrics.Net.Metric.Levenstein = New SimMetrics.Net.Metric.Levenstein()
    '        Dim sim As Double = ls.GetSimilarity(String2Search.ToUpper, str.ToUpper)
    '        If sim >= MinimalSimilarity Then
    '            RetVal = True
    '            Exit For
    '        End If
    '    Next
    '    Return RetVal
    'End Function
    Public Shared Function ExtractPackstationNumber(ByVal InputString As String) As String
        Dim RetVal As Integer = 0
        Dim inparray() As String = InputString.Split(CChar(" "))
        For Each str As String In inparray
            If str.Length >= 6 AndAlso str.Length <= 10 Then
                Dim myInt As Integer = 0
                If Integer.TryParse(str, myInt) Then
                    RetVal = myInt
                    Exit For
                End If
            End If
        Next
        Return RetVal.ToString
    End Function
    Public Shared Function UTF8Decode(str As String) As String
        Dim UTF8 As Encoding = Encoding.UTF8
        Dim Unicode As Encoding = Encoding.Unicode
        Dim xBytes As Byte() = Unicode.GetBytes(str)
        Dim yBytes((xBytes.Length \ 2) - 1) As Byte
        For I As Integer = 0 To yBytes.Length - 1
            yBytes(I) = xBytes(I * 2)
        Next
        Return UTF8.GetString(yBytes)
    End Function
    Public Shared Function StringToByteArray(ByRef str As String) As Byte()
        Dim enc As System.Text.Encoding = System.Text.Encoding.Default
        Return enc.GetBytes(str)
    End Function
    Public Shared Function ByteArrayToString(ByRef Barr() As Byte) As String
        Dim enc As System.Text.Encoding = System.Text.Encoding.Default
        Return enc.GetString(Barr)
    End Function
#End Region
#Region "Security"
    Public Shared Function MD5StringHash(ByVal strString As String) As String
        Dim MD5 As New MD5CryptoServiceProvider
        Dim Data As Byte()
        Dim Result As Byte()
        Dim Res As String = ""
        Dim Tmp As String = ""

        Data = Encoding.ASCII.GetBytes(strString)
        Result = MD5.ComputeHash(Data)
        For i As Integer = 0 To Result.Length - 1
            Tmp = Hex(Result(i))
            If Tmp.Length = 1 Then Tmp = "0" & Tmp
            Res += Tmp
        Next
        Return Res
    End Function
    Public Shared Function ToBase64(ByVal sText As String) As String
        Dim nBytes() As Byte = Encoding.Default.GetBytes(sText)
        Return System.Convert.ToBase64String(nBytes)
    End Function
    Public Shared Function FromBase64(ByVal sText As String) As String
        Dim nBytes() As Byte = System.Convert.FromBase64String(sText)
        Return Encoding.Default.GetString(nBytes)
    End Function
    Public Shared Function GenerateRandomString(ByVal length As Integer) As String
        Dim str As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        Dim ran As New Random
        Dim sb As New StringBuilder
        For i As Integer = 1 To length
            Dim idx As Integer = ran.Next(0, 35)
            sb.Append(str.Substring(idx, 1))
        Next
        Return sb.ToString()
    End Function
#End Region
#Region "Active Directory"
    Public Shared Function ValidateActiveDirectoryLogin(ByVal Domain As String, ByVal Username As String, ByVal Password As String) As Boolean
        Dim group As String = "ConDropAdmin"
        Dim Success As Boolean = False
        Dim Entry As New DirectoryServices.DirectoryEntry("LDAP://" & Domain, Username, Password)
        Dim Searcher As New DirectorySearcher(Entry)
        Searcher.SearchScope = SearchScope.OneLevel
        Try
            Dim Results As SearchResult = Searcher.FindOne
            Success = Not (Results Is Nothing)
        Catch
            Success = False
        End Try

        Return Success
    End Function
    Public Shared Function ValidateActiveDirectoryLoginWithGroup(ByVal ADusername As String, ByVal ADpassword As String, ByVal ADdomain As String, ByVal ADgroup2Check As String) As Boolean
        Dim SuccessLogin As Boolean = False
        Dim ADPath As String = ""
        Dim myDirectoryEntry As New DirectoryEntry

        Try
            ADPath = "LDAP://" & ADdomain
            ADgroup2Check = ADgroup2Check.ToLower()
            If (ADusername <> "" AndAlso ADpassword <> "") Then
                myDirectoryEntry = New DirectoryEntry(ADPath, ADusername, ADpassword)
            Else
                myDirectoryEntry = New DirectoryEntry(ADPath)
            End If

            Dim myDirectorySearcher As New DirectorySearcher(myDirectoryEntry)

            myDirectorySearcher.Filter = "sAMAccountName=" & ADusername
            myDirectorySearcher.PropertiesToLoad.Add("MemberOf")

            Try
                Dim myresult As SearchResult = myDirectorySearcher.FindOne()
                SuccessLogin = Not (myresult Is Nothing)

                Dim NumberOfGroups As Integer = 0
                NumberOfGroups = myresult.Properties("memberOf").Count() - 1
                Dim tempString As String
                While (NumberOfGroups >= 0)
                    tempString = CStr(myresult.Properties("MemberOf").Item(NumberOfGroups))
                    tempString = tempString.Substring(0, tempString.IndexOf(",", 0, StringComparison.Ordinal))
                    tempString = tempString.Replace("CN=", "").ToLower().Trim()
                    If (ADgroup2Check = tempString) Then
                        Return True
                    End If
                    NumberOfGroups = NumberOfGroups - 1
                End While
                SuccessLogin = False
                MessageBoxEx.Show("Die Benutzeranmeldung war nicht erfolgreich." & Environment.NewLine & "" & Environment.NewLine & "Sie haben nicht die erforderlichen Berechtigungen.", "Benutzernameldung", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            Catch
                SuccessLogin = False
                MessageBoxEx.Show("Die Benutzeranmeldung war nicht erfolgreich." & Environment.NewLine & "" & Environment.NewLine & "Benutzername und/oder Passwort sind falsch.", "Benutzernameldung", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            End Try

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
            SuccessLogin = False
            Return False
        End Try

    End Function
    Public Shared Function Authenticate2Domain(ByVal domain As String, ByVal username As String, ByVal password As String) As Boolean
        Dim pwd As New SecureString()
        Dim bAuth As Boolean = False
        Dim entry As DirectoryEntry = Nothing

        'Durchlaufe das Passwort und hänge es dem SecureString an 
        For Each c As Char In password
            pwd.AppendChar(c)
        Next

        'Bewirkt, dass das Passwort nicht mehr verändert werden kann 
        pwd.MakeReadOnly()

        'Passwort wird einem Pointer übergeben, damit dieser später "entschlüsselt" werden kann 
        Dim pPwd As IntPtr = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(pwd)

        Try
            entry = New DirectoryEntry(String.Concat("LDAP://", domain), username, System.Runtime.InteropServices.Marshal.PtrToStringBSTR(pPwd))
            Dim nativeObject As Object = entry.NativeObject
            bAuth = True
        Catch ex As Exception
            bAuth = False
        Finally
            entry.Close()
            entry.Dispose()
        End Try
        Return bAuth
    End Function
    Public Shared Function GetUserMemberOfDomain(ByVal domain As String, ByVal username As String, ByVal password As String, Optional ByRef exeption As Exception = Nothing) As Collections.Generic.List(Of String)
        Dim searcher As DirectorySearcher = Nothing
        Dim colEntry As New Collections.Generic.List(Of String)
        Try
            searcher = New DirectorySearcher(New DirectoryEntry("LDAP://" & domain, username, password))
            searcher.Filter = String.Concat("(&(objectClass=User) (sAMAccountName=", username, "))")
            searcher.PropertiesToLoad.Add("MemberOf")

            Dim result As SearchResult = searcher.FindOne
            For i As Integer = 0 To result.Properties("MemberOf").Count - 1
                Dim sProp As String = CType(result.Properties("MemberOf")(i), String)
                colEntry.Add(sProp.Substring(3, sProp.IndexOf(",") - 3))
            Next

        Catch ex As Exception
            exeption = ex

        Finally
            searcher.Dispose()
        End Try

        Return colEntry

    End Function

#End Region
#Region "System"
    Public Shared Function GetGroups(ByVal userName As String) As List(Of String)
        Dim result As List(Of String) = New List(Of String)()
        Dim wi As WindowsIdentity = New WindowsIdentity(userName)
        For Each group As IdentityReference In wi.Groups
            Try
                result.Add(group.Translate(GetType(NTAccount)).ToString())
            Catch ex As Exception
                Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
                If Helper.IsIDE() Then Stop
            End Try
        Next
        result.Sort()
        Return result
    End Function
    Public Shared Function GetProcess(name As String) As Process
        Dim ReturnValue As Process = Nothing
        For Each clsProcess As Process In Process.GetProcesses()
            If clsProcess.ProcessName.StartsWith(name, True, New CultureInfo("en-US")) Then
                ReturnValue = clsProcess
                Exit For
            End If
        Next
        Return ReturnValue
    End Function
    Public Shared Function IsGuiThread() As Boolean
        If Thread.CurrentThread.Name = "GuiThread" Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Function IsProcessRunning(name As String) As Boolean
        Dim ReturnValue As Boolean = False
        For Each clsProcess As Process In Process.GetProcesses()
            'Debug.Print(clsProcess.ProcessName)
            If clsProcess.ProcessName.StartsWith(name, True, New CultureInfo("en-US")) Then
                ReturnValue = True
                Exit For
            End If
        Next
        Return ReturnValue
    End Function
    Public Shared Function GetArchitecture() As String
        Dim architecture As String = ""
        If Helper_VarConvert.ConvertToInteger(IntPtr.Size) > 4 Then
            architecture = "64 Bit"
        Else
            architecture = "32 Bit"
        End If
        Return architecture
    End Function
    Public Shared Function GetCallingProc() As String
        Dim Trace As New StackTrace()
        Try
            Return Trace.GetFrame(2).GetMethod().Name
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Public Shared Function IsApplicationRunning(ByVal Dateiname As String) As Boolean
        Try
            If Process.GetProcessesByName(Path.GetFileNameWithoutExtension(Dateiname)).Length > 1 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
            Return False
        End Try

    End Function
    Public Shared Function IsNewVersion(ByVal sCurrent As String, ByVal sNew As String) As Boolean
        Dim saCurrrent As String()
        Dim saNew As String()
        Dim nLength As Integer = 0
        Dim IsNewVersionVal As Boolean = False
        If String.IsNullOrWhiteSpace(sCurrent) OrElse String.IsNullOrWhiteSpace(sNew) Then
            Return IsNewVersionVal
        End If
        saCurrrent = Split(sCurrent, ".")
        saNew = Split(sNew, ".")
        nLength = saCurrrent.Length
        If saNew.Length <> nLength Then
            Return IsNewVersionVal
        End If
        For i As Integer = 0 To nLength
            Select Case Val(saNew(i)) - Val(saCurrrent(i))
                Case Is < 0
                    ' Teilnummer der "neuen" Version ist kleiner
                    Exit For
                Case 0
                    ' do nothing - nächste Stelle prüfen
                Case Is > 0
                    ' Teilnummer der "neuen" Version ist größer
                    IsNewVersionVal = True
                    Exit For
            End Select
        Next i
        Return IsNewVersionVal
    End Function

#End Region
#Region "Variablen"
    Public Shared Function Checkvar(ByVal var As Object,
                                ByVal VariType As VarTypeEnu,
                                Optional ByVal trimed As Boolean = False,
                                Optional ByVal DefaultValue As Object = Nothing) As Object
        Dim Var_Type_String As String = ""
        Dim Var_Type_Date As New Date
        'Dim Var_Type_Boolean As Boolean = False
        'Dim Var_Type_Integer As Integer = 0
        'Dim Var_Type_Double As Double = 0
        'Dim Var_Type_Single As Single = 0
        'Dim Var_Type_Long As Single = 0
        'Dim var_bool As Boolean = False
        'Dim var_dbl As Double = 0
        'Dim var_int As Integer = 0
        'Dim var_sng As Single = 0

        Try
            If VariType = Helper.VarTypeEnu.Type_Boolean Then
                Return Helper_VarConvert.ConvertToBoolean(var)
            ElseIf VariType = Helper.VarTypeEnu.Type_Double Then

                Return Helper_VarConvert.ConvertToDouble(var)
            ElseIf VariType = Helper.VarTypeEnu.Type_Integer Then
                Return Helper_VarConvert.ConvertToInteger(var)
            ElseIf VariType = Helper.VarTypeEnu.Type_Long Then
                Return Helper_VarConvert.ConvertToLong(var)
            ElseIf VariType = Helper.VarTypeEnu.Type_Single Then
                Return Helper_VarConvert.ConvertToSingle(var)
            ElseIf VariType = Helper.VarTypeEnu.Type_String Then
                If var Is Nothing OrElse System.Convert.IsDBNull(var) OrElse String.IsNullOrWhiteSpace(var.ToString) Then
                    var = ""
                End If
                If trimed Then
                    Var_Type_String = var.ToString.Trim
                Else
                    Var_Type_String = CStr(var)
                End If
                Return Var_Type_String
            ElseIf VariType = Helper.VarTypeEnu.Type_Date Then
                If var Is Nothing OrElse System.Convert.IsDBNull(var) OrElse String.IsNullOrWhiteSpace(var.ToString) OrElse Not IsDate(var) Then
                    'Var_Type_Date = Nothing
                    Return Nothing
                ElseIf CStr(var) = "1700-01-01 00:00:00" Then
                    'Var_Type_Date = Nothing
                    Return Nothing
                Else
                    Var_Type_Date = CDate(var)
                    Return CDate(var)
                End If
                'Return Var_Type_Date
            ElseIf VariType = Helper.VarTypeEnu.Type_DateAsString Then
                If var Is Nothing OrElse System.Convert.IsDBNull(var) OrElse String.IsNullOrWhiteSpace(var.ToString) Then
                    Var_Type_String = "1700-01-01 00:00:00"
                ElseIf CStr(var) = "1700-01-01 00:00:00" Then
                    Var_Type_String = "1700-01-01 00:00:00"
                Else
                    Var_Type_String = Format(CDate(var), "yyyy-MM-dd HH:mm:ss")
                End If
                Return Var_Type_String
            Else
                If var Is Nothing OrElse System.Convert.IsDBNull(var) OrElse String.IsNullOrWhiteSpace(var.ToString) Then
                    var = ""
                End If
                If trimed Then
                    Var_Type_String = var.ToString.Trim
                Else
                    Var_Type_String = var.ToString
                End If
                Return Var_Type_String
            End If
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False)
            If Helper.IsIDE() Then Stop
            Return ""
        End Try
    End Function
    Public Shared Function Convert2Double(wert As String) As Double
        wert = wert.Trim
        wert = wert.Replace(",", ".")
        If IsNumeric(Val(wert)) Then
            Return Helper_VarConvert.ConvertToDouble(Val(wert))
        Else
            Return 0
        End If
    End Function
    Public Shared Function ConvertHtmlToText(source As String) As String
        Dim result As String
        result = source.Replace(vbCr, " ")
        result = result.Replace(CrLf, " ")
        result = result.Replace(Tab, String.Empty)
        result = Regex.Replace(result, "( )+", " ")
        result = Regex.Replace(result, "<( )*head([^>])*>", "<head>", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "(<( )*(/)( )*head( )*>)", "</head>", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "(<head>).*(</head>)", String.Empty, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<( )*script([^>])*>", "<script>", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "(<( )*(/)( )*script( )*>)", "</script>", RegexOptions.IgnoreCase)
        'result = System.Text.RegularExpressions.Regex.Replace(result, 
        '         @"(<script>)([^(<script>\.</script>)])*(</script>)",
        '         string.Empty, 
        '         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        result = Regex.Replace(result, "(<script>).*(</script>)", String.Empty, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<( )*style([^>])*>", "<style>", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "(<( )*(/)( )*style( )*>)", "</style>", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "(<style>).*(</style>)", String.Empty, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<( )*td([^>])*>", Tab, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<( )*br( )*>", vbCr, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<( )*li( )*>", vbCr, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<( )*div([^>])*>", vbCr & vbCr, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<( )*tr([^>])*>", vbCr & vbCr, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<( )*p([^>])*>", vbCr & vbCr, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<[^>]*>", String.Empty, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "&nbsp;", " ", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "&bull;", " * ", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "&lsaquo;", "<", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "&rsaquo;", ">", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "&trade;", "(tm)", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "&frasl;", "/", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "<", "<", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, ">", ">", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "&copy;", "(c)", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "&reg;", "(r)", RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "&(.{2,6});", String.Empty, RegexOptions.IgnoreCase)
        result = result.Replace(CrLf, vbCr)
        result = Regex.Replace(result, "(" & vbCr & ")( )+(" & vbCr & ")", vbCr & vbCr, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "(" & Tab & ")( )+(" & Tab & ")", Tab & Tab, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "(" & Tab & ")( )+(" & vbCr & ")", Tab & vbCr, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "(" & vbCr & ")( )+(" & Tab & ")", vbCr & Tab, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "(" & vbCr & ")(" & Tab & ")+(" & vbCr & ")", vbCr & vbCr, RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "(" & vbCr & ")(" & Tab & ")+", vbCr & Tab, RegexOptions.IgnoreCase)
        Dim breaks As String = vbCr & vbCr & vbCr
        Dim tabs As String = Tab & Tab & Tab & Tab & Tab

        For index As Integer = 0 To result.Length - 1
            result = result.Replace(breaks, vbCr & vbCr)
            result = result.Replace(tabs, Tab & Tab & Tab & Tab)
            breaks = breaks & System.Convert.ToString(vbCr)
            tabs = tabs & System.Convert.ToString(Tab)
        Next
        Return result
    End Function
    Public Shared Function ExtractNummericFromString(ByVal Text As String) As String
        For i As Integer = 1 To Text.Length
            If Asc(Mid(Text, i, 1)) < 48 OrElse Asc(Mid(Text, i, 1)) > 57 Then
                Mid(Text, i, 1) = " "
            End If
        Next
        Text = Text.Trim.Replace(" ", "")
        Return Text
    End Function
    Public Shared Function IsDateWithinDays(ByVal CheckDate As String, Optional ByVal MaxDays As Integer = 30) As Boolean
        'Umbenennung von IsDateEx ==> IsDateWithinDays
        Dim ResultDate As Date
        Dim ddiff As Integer = Helper_VarConvert.ConvertToInteger(Math.Abs(DateDiff("d", CheckDate, Now)))
        If DateTime.TryParse(CheckDate, ResultDate) AndAlso ddiff <= MaxDays Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Function IsDateWithinDays(ByVal CheckDate As Date, Optional ByVal MaxDays As Integer = 30) As Boolean
        'Umbenennung von IsDateEx ==> IsDateWithinDays
        Dim ResultDate As Date
        Dim ddiff As Integer = Helper_VarConvert.ConvertToInteger(Math.Abs(DateDiff("d", CheckDate, Now)))
        If DateTime.TryParse(CheckDate.ToString, ResultDate) AndAlso ddiff <= MaxDays Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Function isEven(numToCheck As Integer) As Boolean
        Return (numToCheck And 1) = 0
    End Function
    Public Shared Function IncrementString(ByVal strString As String) As String
        Dim lngLenString As Long
        Dim strChar As String
        Dim lngI As Long

        lngLenString = strString.Length
        ' Start at far right
        For lngI = lngLenString To 0 Step -1
            ' If we reach the far left then add an A and exit
            If lngI = 0 Then
                strString = "a" & strString
                Exit For
            End If
            ' Consider next character
            strChar = Mid(strString, Helper_VarConvert.ConvertToInteger(lngI), 1)
            If strChar = "z" Then
                ' If we find Z then increment this to A
                ' and increment the character after this (in next loop iteration)
                strString = Left$(strString, Helper_VarConvert.ConvertToInteger(lngI - 1)) & "a" & Mid(strString, Helper_VarConvert.ConvertToInteger(lngI + 1), Helper_VarConvert.ConvertToInteger(lngLenString))
            Else
                ' Increment this non-Z and exit
                strString = Left$(strString, Helper_VarConvert.ConvertToInteger(lngI - 1)) & Chr(Asc(strChar) + 1) & Mid(strString, Helper_VarConvert.ConvertToInteger(lngI + 1), Helper_VarConvert.ConvertToInteger(lngLenString))
                Exit For
            End If
        Next lngI

        Return strString

    End Function
    Public Shared Function stripString(ByVal StringValue As String, ByVal Delimiter As String) As String
        Dim retval As String = StringValue
        Try
            If Not StringValue IsNot Nothing AndAlso Delimiter = Nothing AndAlso Not String.IsNullOrWhiteSpace(StringValue) AndAlso Not String.IsNullOrWhiteSpace(Delimiter) Then
                Dim strparts() As String = StringValue.Split(CType(Delimiter, Char))
                retval = strparts(0)
                'Dim po As Integer = StringValue.contains( Delimiter)
                'If po > 0 Then
                '    retval = Mid(StringValue, 1, po - 1)
                'End If
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try

        Return retval
    End Function
    Public Shared Function EncryptString(ByVal Decrypted As String, Optional ByVal KeyLength As Integer = 256) As String
        'Dim cCrypt As cCrypt = New cCrypt
        If Not String.IsNullOrWhiteSpace(Decrypted) Then
            Helper_cCrypt.Encrypt(KeyLength, Decrypted)
            Return Helper_cCrypt.EncryptedString
        Else
            Return ""
        End If
    End Function
    Public Shared Function DecryptString(ByVal Encrypted As String, Optional ByVal KeyLength As Integer = 256) As String
        'Dim cCrypt As cCrypt = New cCrypt
        If Not String.IsNullOrWhiteSpace(Encrypted) Then
            Helper_cCrypt.Decrypt(KeyLength, Helper_VarConvert.ConvertToString(Encrypted, False, ""))
            Return Helper_cCrypt.DecryptedString
        Else
            Return ""
        End If
    End Function
    Public Shared Function Inch2Millimeter(ByVal Inch As Double) As Double
        Dim mm As Double = Inch * 25.4 * 10 / 10
        Return mm
    End Function
    Public Shared Function StringList_SplitIntoChunks(keys As List(Of String), chunkSize As Integer) As List(Of List(Of String))
        Return keys _
.Select(Function(x, i) New With {Key .Index = i, Key .Value = x}) _
.GroupBy(Function(x) (x.Index \ chunkSize)) _
.Select(Function(x) x.Select(Function(v) v.Value).ToList()) _
.ToList()
    End Function
    Public Shared Function splitList(ByVal OriginalList As List(Of String), ByVal chunkSize As Integer) As List(Of List(Of String))
        Dim TargetList As New List(Of List(Of String))
        Dim i As Integer = 0
        While i < OriginalList.Count
            TargetList.Add(OriginalList.GetRange(i, Math.Min(chunkSize, OriginalList.Count - i)))
            i += chunkSize
        End While
        Return TargetList
    End Function
    Public Shared Function GetRandomByString(ByVal ResultLength As Integer,
                                         Optional SelectFrom As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz12345678901234567890") As String
        Dim ran As New Random
        Dim strb As New System.Text.StringBuilder
        Dim chararray() As Char = SelectFrom.ToCharArray
        For i As Integer = 1 To ResultLength
            Dim ti As Integer = ran.Next(0, chararray.Length)
            strb.Append(chararray(ti))
        Next
        Return strb.ToString
    End Function
    Private Shared Function ReplaceFirst(ByVal text As String, ByVal search As String, ByVal replace As String) As String
        Dim pos As Integer = text.IndexOf(search)
        If pos < 0 Then
            Return text
        End If
        Return text.Substring(0, pos) & replace & text.Substring(pos + search.Length)
    End Function

#End Region
#Region "Windows Firewall"
    Public Shared Sub CreateFirewallRule(ByVal Profile As FirewallProfiles,
                                     ByVal RuleName As String,
                                     ByVal PortNumber As UShort,
                                     ByVal Protocol As FirewallProtocol,
                                     ByVal Action As FirewallAction,
                                     ByVal direction As FirewallDirection)
        Dim ActionText As String = ""
        Dim DirectionText As String = ""
        Dim found As Boolean = False
        Dim allRules() As IFirewallRule = FirewallManager.Instance.Rules.ToArray()
        For Each FWRule As IFirewallRule In allRules
            If FWRule.Direction = direction AndAlso
            FWRule.Protocol = Protocol AndAlso
            FWRule.LocalPorts.Contains(PortNumber) Then
                found = True
                Exit For
            End If
        Next
        If Not found Then
            If Action = FirewallAction.Allow Then
                ActionText = "geöffnet"
            Else
                ActionText = "geschlossen"
            End If
            If direction = FirewallDirection.Inbound Then
                DirectionText = "eingehend"
            Else
                DirectionText = "ausgehend"
            End If
            Select Case MessageBox.Show("Für die ordnungsgemäße Funktion des Programms muss der Port " & PortNumber & " " & DirectionText & " für das Protokoll " & Protocol.ToString & " in der Windows Firewall " & ActionText & " sein." &
                                        Environment.NewLine & Environment.NewLine &
                                        "Soll dieser Port jetzt durch eine entsprechende Firewall-Regel " & ActionText & " werden?",
                                        "Fehlende Firewall-Regel für " & Application.ProductName,
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question)
                Case DialogResult.Yes
                    Dim rule As IFirewallRule = FirewallManager.Instance.CreatePortRule(Profile, RuleName, Action, PortNumber, Protocol)
                    rule.Direction = direction
                    FirewallManager.Instance.Rules.Add(rule)
                Case DialogResult.No
            End Select
        End If
    End Sub
#End Region

#Region "Unsortiert"
    Public Shared Sub LocalizeString(sender As Object, e As LocalizeEventArgs)
        If e.LocalizedValue.Trim = "" Then
            'Debug.Print(e.Key & " - " & e.LocalizedValue)
        End If
        Select Case e.Key
            Case LocalizationKeys.MessageBoxAbortButton
                e.LocalizedValue = "abbrechen"
                e.Handled = True
            Case LocalizationKeys.MessageBoxCancelButton
                e.LocalizedValue = "abbrechen"
                e.Handled = True
            Case LocalizationKeys.MessageBoxCloseButton
                e.LocalizedValue = "schließen"
                e.Handled = True
            Case LocalizationKeys.MessageBoxContinueButton
                e.LocalizedValue = "fortfahren"
                e.Handled = True
            Case LocalizationKeys.MessageBoxHelpButton
                e.LocalizedValue = "Hilfe"
                e.Handled = True
            Case LocalizationKeys.MessageBoxIgnoreButton
                e.LocalizedValue = "ignorieren"
                e.Handled = True
            Case LocalizationKeys.MessageBoxNoButton
                e.LocalizedValue = "Nein"
                e.Handled = True
            Case LocalizationKeys.MessageBoxOkButton
                e.LocalizedValue = "Ok"
                e.Handled = True
            Case LocalizationKeys.MessageBoxRetryButton
                e.LocalizedValue = "wiederholen"
                e.Handled = True
            Case LocalizationKeys.MessageBoxTryAgainButton
                e.LocalizedValue = "nochmal versuchen"
                e.Handled = True
            Case LocalizationKeys.MessageBoxYesButton
                e.LocalizedValue = "Ja"
                e.Handled = True
            Case LocalizationKeys.TimeSelectorClearButton
                e.LocalizedValue = "leeren"
                e.Handled = True
            Case LocalizationKeys.TimeSelectorOkButton
                e.LocalizedValue = "Ok"
                e.Handled = True
            Case LocalizationKeys.TimeSelectorHourLabel
                e.LocalizedValue = "Stunde"
                e.Handled = True
            Case LocalizationKeys.TimeSelectorMinuteLabel
                e.LocalizedValue = "Minute"
                e.Handled = True
            Case LocalizationKeys.MonthCalendarClearButtonText
                e.LocalizedValue = "leeren"
                e.Handled = True
            Case LocalizationKeys.MonthCalendarTodayButtonText
                e.LocalizedValue = "heute"
                e.Handled = True
            Case LocalizationKeys.MdiSystemItemMenuRestore
                e.LocalizedValue = "wiederherstellen"
                e.Handled = True
            Case LocalizationKeys.MdiSystemItemMenuMove
                e.LocalizedValue = "bewegen"
                e.Handled = True
            Case LocalizationKeys.MdiSystemItemMenuSize
                e.LocalizedValue = "Größe"
                e.Handled = True
            Case LocalizationKeys.MdiSystemItemMenuMinimize
                e.LocalizedValue = "minimieren"
                e.Handled = True
            Case LocalizationKeys.MdiSystemItemMenuMaximize
                e.LocalizedValue = "maximieren"
                e.Handled = True
            Case LocalizationKeys.MdiSystemItemMenuClose
                e.LocalizedValue = "schliessen"
                e.Handled = True
            Case LocalizationKeys.MdiSystemItemMenuNext
                e.LocalizedValue = "nächstes"
                e.Handled = True

                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
                'Case LocalizationKeys.
                '    e.LocalizedValue = ""
                '    e.Handled = True
        End Select
    End Sub
    Public Shared Function GetOrigin(ByVal PlatformSelect As Helper.Platform) As String
        Select Case PlatformSelect
            Case Helper.Platform.Afterbuy
                Return "Afterbuy"
            Case Helper.Platform.Plenty
                Return "Plenty"
            Case Helper.Platform.CSV
                Return "CSV"
            Case Helper.Platform.eBay
                Return "eBay"
            Case Helper.Platform.Amazon
                Return "Amazon"
            Case Helper.Platform.Manuell
                Return "Manuell"
            Case Helper.Platform.Shopware
                Return "Shopware"
            Case Helper.Platform.Otto
                Return "Otto"
            Case Else
                Return "CSV"
        End Select

    End Function

#End Region
End Class
