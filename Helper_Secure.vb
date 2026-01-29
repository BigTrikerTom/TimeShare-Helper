
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO
'Imports log4net
'Imports MySql.Data.MySqlClient
Imports System.Threading
Imports TimeShare_Helper
Imports TimeShare_Error

Public NotInheritable Class Helper_Secure
    Private Shared enc As System.Text.UTF8Encoding
    Private Shared encryptor As ICryptoTransform
    Private Shared decryptor As ICryptoTransform

    Public Enum CWAction
        Encrypt
        Decrypt
    End Enum
    Public Structure ReturnStruct
        Public Success As Boolean
        Public StrValue As String
        Public SaltValue As String
    End Structure

    Public Shared Function DoPwd(ByVal UserNameInput As String, ByVal password As String) As ReturnStruct  ' Create New User
        Dim RetVal As New ReturnStruct
        Dim Salt As String = ""
        RetVal.StrValue = ""
        RetVal.SaltValue = ""
        RetVal.Success = False

        Salt = Helper_Secure.CreateRandomSalt()
        RetVal.SaltValue = CryptWrapper(Salt, CWAction.Encrypt) ' ==> Länge: 152
        RetVal.StrValue = Hash512(password, RetVal.SaltValue)
        RetVal.Success = True
        Dim SaltLlength = RetVal.SaltValue.Length
        Dim PwdLength = RetVal.StrValue.Length
        Return RetVal
    End Function
    Public Shared Function DoPwd(ByVal UserNameInput As String, ByVal PasswordInput As String, ByVal password As String, ByVal Salt As String) As ReturnStruct ' Validate User
        Dim RetVal As New ReturnStruct
        RetVal.StrValue = ""
        RetVal.SaltValue = ""
        RetVal.Success = False

        'Salt = CryptWrapper(Salt, CWAction.Decrypt)
        Dim StrCheck As String = Hash512(PasswordInput, Salt)
        If StrCheck = password Then
            RetVal.Success = True
        Else
            RetVal.Success = False
        End If
        Return RetVal
    End Function
    Private Shared Function CryptWrapper(ByVal Value As String, ByVal Action As CWAction) As String
        Dim RetVal As String = ""
        Dim KEY_128 As Byte() = {42, 1, 52, 67, 231, 13, 94, 101, 123, 6, 0, 12, 32, 91, 4, 111, 31, 70, 21, 141, 123, 142, 234, 82, 95, 129, 187, 162, 12, 55, 98, 23}
        Dim IV_128 As Byte() = {234, 12, 52, 44, 214, 222, 200, 109, 2, 98, 45, 76, 88, 53, 23, 78}
        Dim symmetricKey As RijndaelManaged = New RijndaelManaged()
        Try
            symmetricKey.Mode = CipherMode.CBC
            enc = New System.Text.UTF8Encoding
            encryptor = symmetricKey.CreateEncryptor(KEY_128, IV_128)
            decryptor = symmetricKey.CreateDecryptor(KEY_128, IV_128)
            If Action = CWAction.Decrypt Then
                RetVal = Helper_Secure.StrDecrypt(Value)
            ElseIf Action = CWAction.Encrypt Then
                RetVal = Helper_Secure.StrEncrypt(Value)
            End If
            enc = Nothing
            encryptor = Nothing
            decryptor = Nothing
            symmetricKey = Nothing
        Catch ex As Exception
            call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return RetVal


    End Function
    Private Shared Function StrEncrypt(ByVal Value As String) As String
        Dim RetVal As String = ""
        Try
            If Not String.IsNullOrEmpty(Value) Then
                Dim memoryStream As MemoryStream = New MemoryStream()
                Dim cryptoStream As CryptoStream = New CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write)
                cryptoStream.Write(enc.GetBytes(Value), 0, Value.Length)
                cryptoStream.FlushFinalBlock()
                RetVal = Convert.ToBase64String(memoryStream.ToArray())
                memoryStream.Close()
                cryptoStream.Close()
            End If

        Catch ex As Exception
            call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return RetVal
    End Function
    Private Shared Function StrDecrypt(ByVal Value As String) As String
        Dim RetVal As String = ""
        Try
            'Debug.Print(Value)
            Dim cypherTextBytes As Byte() = Convert.FromBase64String(Value)
            Dim memoryStream As MemoryStream = New MemoryStream(cypherTextBytes)
            Dim cryptoStream As CryptoStream = New CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read)
            Dim plainTextBytes(cypherTextBytes.Length) As Byte
            Dim decryptedByteCount As Integer = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)
            memoryStream.Close()
            cryptoStream.Close()
            RetVal = enc.GetString(plainTextBytes, 0, decryptedByteCount)

        Catch ex As Exception
            call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return RetVal
    End Function

    Public Shared Function Hash512(ByVal password As String, ByVal salt As String) As String
        salt = CryptWrapper(salt, CWAction.Decrypt)
        Dim convertedToBytes As Byte() = Encoding.UTF8.GetBytes(password & salt)
        Dim hashType As HashAlgorithm = New SHA512Managed()
        Dim hashBytes As Byte() = hashType.ComputeHash(convertedToBytes)
        Dim hashedResult As String = Convert.ToBase64String(hashBytes)
        Return hashedResult
    End Function
    Private Shared Function CreateRandomSalt() As String
        'the following is the string that will hold the salt charachters
        Dim mix As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+=][}{<>"
        Dim salt As String = ""
        Dim rnd As New Random
        'Dim sb As New StringBuilder
        For i As Integer = 1 To 100 'Length of the salt
            Dim x As Integer = rnd.Next(0, mix.Length - 1)
            salt &= (mix.Substring(x, 1))
        Next
        Return salt
    End Function
End Class
