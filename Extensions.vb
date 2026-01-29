' ######################################################################
' ## Copyright (c) 2021 TimeShareIt GdbR
' ## by Thomas Steger
' ## File creation Date: 2021-1-29 04:37
' ## File update Date: 2021-3-19 18:54
' ## Filename: modHelper_Extensions.vb (F:\ConDrop\ConDrop_Server\modHelper_Extensions.vb)
' ## Project: ConDrop_Server
' ## Last User: stegert
' ######################################################################
'
'

Imports System.Globalization
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

Imports Limilabs.Client.DNS

Imports log4net
Imports log4net.Appender

Imports log4net.Layout

'Imports log4net
'Imports log4net.Appender
'Imports log4net.Layout

Public Module Extensions
#Region "ToInteger"
    <Extension()>
    Public Function IsInteger(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then
            Return False
        Else
            Return Integer.TryParse(value, Nothing)
        End If
    End Function

    <Extension()>
    Public Function ToInteger(ByVal value As Object) As Integer
        If IsNothing(value) OrElse IsDBNull(value) Then
            Return 0
        Else
            Dim valuestr As String = value.ToString
            If Integer.TryParse(valuestr, 0) Then
                Return Integer.Parse(valuestr)
            Else
                Return 0
            End If
        End If
    End Function
#End Region
#Region "ToDouble"
    <Extension()>
    Public Function IsDouble(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then
            Return False
        Else
            Return Double.TryParse(value, 0)
        End If
    End Function
    <Extension()>
    Public Function ToDouble(ByVal value As Object,
                             Optional DefaultValue As Double = Nothing) As Double
        Dim Result As Double = 0
        If IsNothing(value) OrElse IsDBNull(value) OrElse String.IsNullOrEmpty(CStr(value)) Then
            If Not IsNothing(DefaultValue) Then
                Return DefaultValue
            Else
                Return 0
            End If
        Else
            Dim valuestr As String = value.ToString
            If Application.CurrentCulture.Name = "de-DE" Then
                valuestr = value.ToString.Replace(".", ",")
            Else
                valuestr = value.ToString.Replace(",", ".")
            End If
Dim style As NumberStyles = NumberStyles.AllowDecimalPoint
            Dim rb As Boolean  = Double.TryParse(valuestr, style, Application.CurrentCulture, Result)
            If rb Then
                Return Result
            Else
                Return 0
            End If
        End If
    End Function
#End Region
#Region "ToString"
    Public Function ToString(ByVal value As Object, ByVal trimed As Boolean) As String
        If IsNothing(value) OrElse IsDBNull(value) OrElse CStr(value) = "" Then
            value = ""
        End If
        If trimed Then
            Return Trim(CStr(value))
        Else
            Return CStr(value)
        End If

    End Function

#End Region

#Region "ToSingle"
    <Runtime.CompilerServices.Extension()>
    Public Function IsSingle(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then
            Return False
        Else
            Return Single.TryParse(value, 0)
        End If
    End Function
    <Extension()>
    Public Function ToSingle(ByVal value As Object,
                             Optional DefaultValue As Single = Nothing) As Single
        If IsNothing(value) OrElse IsDBNull(value) OrElse String.IsNullOrEmpty(CStr(value)) Then
            If Not IsNothing(DefaultValue) Then
                Return DefaultValue
            Else
                Return 0
            End If
        Else
            Dim valuesng As String = value.ToString.Replace(",", ".")
            If Single.TryParse(valuesng, 0) Then
                Return Single.Parse(valuesng)
            Else
                Return 0
            End If
        End If
    End Function
#End Region
#Region "ToBoolean"
    <Extension()>
    Public Function IsBoolean(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then
            Return False
        Else
            Return Boolean.TryParse(value, False)
        End If
    End Function
    <Extension()>
    Public Function ToBoolean(ByVal value As Object) As Boolean
        If IsNothing(value) OrElse IsDBNull(value) Then
            Return False
        Else
            Dim valuebool As String = value.ToString
            If valuebool = "1" Then
                Return True
            ElseIf valuebool = "0" Then
                Return False
            ElseIf Boolean.TryParse(valuebool, False) Then
                Return Boolean.Parse(valuebool)
            Else
                Return False
            End If
        End If
    End Function
#End Region
#Region "ToLong"
    <Extension()>
    Public Function IsLong(ByVal value As String) As Boolean
        If String.IsNullOrEmpty(value) Then
            Return False
        Else
            Return Long.TryParse(value, Nothing)
        End If
    End Function
    <Extension()>
    Public Function ToLong(ByVal value As Object) As Long
        If IsNothing(value) OrElse IsDBNull(value) Then
            Return 0
        Else
            Dim valuestr As String = value.ToString
            If Long.TryParse(valuestr, 0) Then
                Return Long.Parse(valuestr)
            Else
                Return 0
            End If
        End If
    End Function
#End Region
#Region "Log4Net"
    <Extension()>
    Public Sub Notice(ByVal log As ILog, ByVal message As Object)
        log.Logger.Log(Nothing, log4net.Core.Level.Notice, message, Nothing)
    End Sub
    <Extension()>
    Public Sub Trace(ByVal log As ILog, ByVal message As Object)
        log.Logger.Log(Nothing, log4net.Core.Level.Trace, message, Nothing)
    End Sub
    <Extension()>
    Public Sub Verbose(ByVal log As ILog, ByVal message As Object)
        log.Logger.Log(Nothing, log4net.Core.Level.Verbose, message, Nothing)
    End Sub
    <Extension()>
    Public Sub AddStringParameterToAppender(ByVal appender As log4net.Appender.AdoNetAppender, ByVal paramName As String, ByVal size As Integer, ByVal conversionPattern As String)
        Dim param As AdoNetAppenderParameter = New AdoNetAppenderParameter()
        param.ParameterName = paramName
        param.DbType = System.Data.DbType.String
        param.Size = size
        param.Layout = New Layout2RawLayoutAdapter(New PatternLayout(conversionPattern))
        appender.AddParameter(param)
    End Sub
    <Extension()>
    Public Sub AddErrorParameterToAppender(ByVal appender As log4net.Appender.AdoNetAppender, ByVal paramName As String, ByVal size As Integer)
        Dim param As AdoNetAppenderParameter = New AdoNetAppenderParameter()
        param.ParameterName = paramName
        param.DbType = System.Data.DbType.String
        param.Size = size
        param.Layout = New Layout2RawLayoutAdapter(New ExceptionLayout())
        appender.AddParameter(param)
    End Sub

    <Extension()>
    Public Sub AddDateTimeParameterToAppender(ByVal appender As log4net.Appender.AdoNetAppender, ByVal paramName As String)
        Dim param As AdoNetAppenderParameter = New AdoNetAppenderParameter()
        param.ParameterName = paramName
        param.DbType = DbType.DateTime
        'param.Layout = New RawTimeStampLayout
        param.Layout = New log4net.Layout.RawUtcTimeStampLayout
        'param.Layout = New RawPropertyLayout
        appender.AddParameter(param)
    End Sub
    <Extension()>
    Public Sub AddInt32ParameterToAppender(ByVal appender As log4net.Appender.AdoNetAppender, ByVal paramName As String, ByVal conversionPattern As String)
        Dim param As AdoNetAppenderParameter = New AdoNetAppenderParameter()
        param.ParameterName = paramName
        param.DbType = DbType.Int32
        param.Layout = New Layout2RawLayoutAdapter(New PatternLayout(conversionPattern))
        appender.AddParameter(param)
    End Sub

#End Region


    <Extension()>
    Friend Function GetFilesByExtensions(ByVal dir As DirectoryInfo, ParamArray extensions As String()) As IEnumerable(Of FileInfo)
        If extensions Is Nothing Then Throw New ArgumentNullException("extensions")
        Dim files As IEnumerable(Of FileInfo) = dir.EnumerateFiles()
        Return files.Where(Function(f) extensions.Contains(f.Extension))
    End Function

    <Extension()>
    Friend Function TryAdd(Of TKey, TValue)(ByVal dictionary As IDictionary(Of TKey, TValue), ByVal key As TKey, ByVal value As TValue) As Boolean
        Try
            If dictionary Is Nothing Then
                Throw New ArgumentNullException(NameOf(dictionary))
            End If
            If Not dictionary.ContainsKey(key) Then
                dictionary.Add(key, value)
                Return True
            End If
        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return False
    End Function

End Module
