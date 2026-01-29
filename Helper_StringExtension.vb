' ######################################################################
' ## Copyright (c) 2021 TimeShareIt GdbR
' ## by Thomas Steger
' ## File creation Date: 2021-3-2 08:08
' ## File update Date: 2021-3-19 18:54
' ## Filename: Helper_StringExtension.vb (F:\ConDrop\ConDrop_Server\Helper_StringExtension.vb)
' ## Project: ConDrop_Server
' ## Last User: stegert
' ######################################################################
'
'

Imports System.Runtime.CompilerServices

Module Helper_StringExtension
    <Extension()>
    Function ReplaceFirst(ByVal text As String, ByVal search As String, ByVal replace As String) As String
        Dim pos As Integer = text.IndexOf(search)
        If pos < 0 Then
            Return text
        End If
        Return text.Substring(0, pos) & replace & text.Substring(pos + search.Length)
    End Function
    <Extension()>
    Function ReplaceLast(ByVal text As String, ByVal search As String, ByVal replace As String) As String
        Dim pos As Integer = text.LastIndexOf(search)
        If pos < 0 Then
            Return text
        End If
        Return text.Substring(0, pos) & replace & text.Substring(pos + search.Length)
    End Function
End Module
