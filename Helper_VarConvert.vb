' ######################################################################
' ## Copyright (c) 2021 TimeShareIt GdbR
' ## by Thomas Steger
' ## File creation Date: 2021-1-29 04:37
' ## File update Date: 2021-3-19 19:33
' ## Filename: Helper_Convert.vb
' ## Project: ConDrop_Server
' ## Last User: stegert
' ######################################################################
'
'

Imports System.Globalization
Imports System.DateTime
Public Class Helper_VarConvert
    Public Shared Function ConvertToDouble(ByVal var_Object As Object, ByVal Optional DefaultValue As Double = 0, ByVal Optional Culture As CultureInfo = nothing) As Double
        Dim retVal As Double = DefaultValue
        Try
            Culture = CultureInfo.CurrentCulture
            If IsNothing(var_Object) Then
                retVal = DefaultValue
            ElseIf TypeOf var_Object Is Double Then
                Return DirectCast(var_Object, Double)
            Else
                If var_Object.ToString.Contains(",") AndAlso var_Object.ToString.Contains(".") Then
                    Dim var_string As String = var_Object.ToString.Trim
                    var_string = var_string.Replace(".", ",")
                    var_string = var_string.ReplaceLast(",", ".").Replace(",", "")
                End If
                If (Double.TryParse(System.Convert.ToString(var_Object).Trim, System.Globalization.NumberStyles.Float, Culture, retVal)) Then
                    Return retVal
                Else
                    retVal = DefaultValue
                End If
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retVal
    End Function
    Public Shared Function ConvertToInteger(ByVal var_Object As Object, ByVal Optional DefaultValue As Integer = 0) As Integer
        Dim retVal As Integer = DefaultValue
        Try
            If IsNothing(var_Object) Then
                Return DefaultValue
            ElseIf TypeOf var_Object Is Integer Then
                Return DirectCast(var_Object, Integer)
            ElseIf Not Integer.TryParse(System.Convert.ToString(var_Object).Trim, retVal) Then
                Return retVal
            Else
                Return retVal
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retVal
    End Function
    Public Shared Function ConvertToSingle(ByVal var_Object As Object, ByVal Optional DefaultValue As Single = 0) As Single
        Dim retVal As Single = DefaultValue
        Try
            If Not Single.TryParse(System.Convert.ToString(var_Object).Trim, retVal) Then
                retVal = DefaultValue
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retVal
    End Function
    Public Shared Function ConvertToShort(ByVal var_Object As Object, ByVal Optional DefaultValue As Short = 0) As Short
        Dim retVal As Short = DefaultValue
        Try
            If Not Short.TryParse(System.Convert.ToString(var_Object).Trim, retVal) Then
                retVal = DefaultValue
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retVal
    End Function
    Public Shared Function ConvertToUShort(ByVal var_Object As Object, ByVal Optional DefaultValue As UShort = 0) As UShort
        Dim retVal As UShort = DefaultValue
        Try
            If Not UShort.TryParse(System.Convert.ToString(var_Object).Trim, retVal) Then
                retVal = DefaultValue
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retVal
    End Function
    Public Shared Function ConvertToLong(ByVal var_Object As Object, ByVal Optional DefaultValue As Long = 0) As Long
        Dim retVal As Long = DefaultValue
        Try
            If Not Long.TryParse(System.Convert.ToString(var_Object).Trim, retVal) Then
                retVal = DefaultValue
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retVal
    End Function
    Public Shared Function ConvertToString(ByVal var_Object As Object, ByVal Optional trimmed As Boolean = False, ByVal Optional DefaultValue As String = "") As String
        Dim retVal As String = DefaultValue
        Try
            If var_Object Is Nothing Then
                retVal = DefaultValue
            Else
                retVal = System.Convert.ToString(var_Object).Trim
            End If
            If trimmed Then
                Return retVal.Trim
            Else
                Return retVal
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retVal
    End Function
    Public Shared Function ConvertToBoolean(ByVal var_Object As Object, ByVal Optional DefaultValue As Boolean = False) As Boolean
        Dim retVal As Boolean = DefaultValue
        Try
            If var_Object Is Nothing Then
                retVal = DefaultValue
            ElseIf var_Object.ToString.Trim = "1" OrElse var_Object.ToString.Trim.ToLower = "true" Then
                retVal = true
            ElseIf var_Object.ToString.Trim = "0" OrElse var_Object.ToString.Trim.ToLower = "false" Then
                retVal = false
            Else
                If Not Boolean.TryParse(var_Object.ToString.Trim, retVal) Then
                    retVal = DefaultValue
                End If
            End If

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retVal
    End Function
    Public Shared Function ConvertToDate(ByVal var_Object As Object, ByVal Optional DefaultValue As Date = Nothing) As Date
        Dim retVal As New Date
        Try
            If DefaultValue = Date.MinValue Then
                DefaultValue = Now
            End If
            If Not DateTime.TryParse(var_Object.ToString.Trim, retVal) Then
                retVal = DefaultValue
            End If
        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retVal
    End Function
    Public Shared Function ConvertDateToString(ByVal var_Object As Object, ByVal Optional StringFormat As String = "dd.MM.yyyy") As String
        Dim retValD As New Date
        Dim retVal As String = ""
        Try
            If Not DateTime.TryParse(var_Object.ToString.Trim, retValD) Then
                retValD = CDate(Nothing)
            End If
            retValD = DateTime.Parse(retValD.ToString.Trim)

            retVal = retValD.ToString(StringFormat).Trim

        Catch ex As Exception
            Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return retVal
    End Function

End Class