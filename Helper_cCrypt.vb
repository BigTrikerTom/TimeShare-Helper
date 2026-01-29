' ######################################################################
' ## Copyright (c) 2021 TimeShareIt GdbR
' ## by Thomas Steger
' ## File creation Date: 2020-12-15 17:18
' ## File update Date: 2021-8-27 19:52
' ## Filename: Helper_cCrypt.vb (F:\++++ Code Share\classes\Helper_cCrypt.vb)
' ## Project: ConDrop_Server
' ## Last User: stegert
' ######################################################################
'
'

Option Strict On

Imports System
Imports System.IO
'Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports DevComponents.DotNetBar
Imports MySql.Data.MySqlClient
Imports System.DateTime
'Imports System.Security.Cryptography.X509Certificates
'Imports Org.BouncyCastle.Pkcs
'Imports Org.BouncyCastle.Security
'Imports Org.BouncyCastle.Crypto

'Imports System.IO
'Imports iText.Bouncycastle.Crypto
'Imports iText.Bouncycastle.X509
'Imports iText.Commons.Bouncycastle.Cert
'Imports iText.Commons.Bouncycastle.Crypto
'Imports iText.Commons.Utils
'Imports iText.Forms.Form.Element
'Imports iText.Kernel.Colors
'Imports iText.Kernel.Pdf
'Imports iText.Layout
'Imports iText.Layout.Element
'Imports iText.Signatures
Imports TimeShare_Error
'Imports iText.Signatures
'Imports iText.Commons.Utils
'Imports iText.Kernel.Pdf
'Imports iText.Layout
'Imports iText.Layout.Element
'Imports iText.Kernel.Colors
'Imports iText.Forms.Form.Element
'Imports iText.IO.Util
Imports Org.BouncyCastle.Pkcs
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports Org.BouncyCastle.Security
Imports Org.BouncyCastle.X509
Imports X509Certificate = Org.BouncyCastle.X509.X509Certificate
'Imports NPOI.POIFS.Crypt
Imports Newtonsoft.Json








Public Class Helper_cCrypt
    Private Shared _saltPrefix As String
    Friend Shared Property SaltPrefix() As String
        Get
            Return _saltPrefix
        End Get
        Set(ByVal value As String)
            _saltPrefix = value
        End Set
    End Property
#Region "Zustandsvariablen"
    Private Shared EncryptedString_ As String
    Private Shared DecryptedString_ As String
#End Region


#Region "Methoden"
    ' Verschlüsseln
    Public Shared Sub Encrypt(ByVal AESKeySize As Int32, ByVal DecryptedStringRes As String)
        Dim Password As String = "wvlIelaulv2fHRMFWPICqNW5d"
        Dim Salt() As Byte
        If SaltPrefix Is Nothing Then
            SaltPrefix = ""
        End If
        Salt = System.Text.Encoding.UTF8.GetBytes(SaltPrefix & "9Tbd4JWZe6FNPibLadXzE6lJSZbC6bRZAELL4iqtBTLu5nN")
        Dim GenerierterKey As New Rfc2898DeriveBytes(Password, Salt)
        Using AES As New AesManaged
            AES.KeySize = AESKeySize
            AES.BlockSize = 128
            AES.Key = GenerierterKey.GetBytes(AES.KeySize \ 8)
            AES.IV = GenerierterKey.GetBytes(AES.BlockSize \ 8)
            Using ms As New IO.MemoryStream
                Using cs As New CryptoStream(ms, AES.CreateEncryptor(), CryptoStreamMode.Write)
                    Dim Data() As Byte
                    Data = System.Text.Encoding.UTF8.GetBytes(DecryptedStringRes)
                    cs.Write(Data, 0, Data.Length)
                    cs.FlushFinalBlock()
                End Using
                Try
                    EncryptedString_ = System.Convert.ToBase64String(ms.ToArray)
                Catch ex As Exception
                    EncryptedString_ = ""
                End Try
            End Using
            AES.Clear()
        End Using
    End Sub
    ' Entschlüsseln
    Public Shared Sub Decrypt(ByVal AESKeySize As Int32, ByVal EncryptedStringVal As String)
        Dim Password As String = "wvlIelaulv2fHRMFWPICqNW5d"
        Dim Salt() As Byte

        Try
            Using AES As New AesManaged
                If EncryptedStringVal.Trim = "" OrElse EncryptedStringVal.Trim = "Oy4ZYRTB1iU2iMhA0GjAzw==" Then
                    DecryptedString_ = ""
                    Exit Sub
                End If
                Salt = System.Text.Encoding.UTF8.GetBytes("9Tbd4JWZe6FNPibLadXzE6lJSZbC6bRZAELL4iqtBTLu5nN")
                Dim GenerierterKey As New Rfc2898DeriveBytes(Password, Salt)
                AES.KeySize = AESKeySize
                AES.BlockSize = 128
                AES.Key = GenerierterKey.GetBytes(AES.KeySize \ 8)
                AES.IV = GenerierterKey.GetBytes(AES.BlockSize \ 8)
                Using ms As New IO.MemoryStream
                    Using cs As New CryptoStream(ms, AES.CreateDecryptor(), CryptoStreamMode.Write)
                        Dim Data() As Byte
                        Data = System.Convert.FromBase64String(EncryptedStringVal)
                        cs.Write(Data, 0, Data.Length)
                        cs.FlushFinalBlock()
                    End Using
                    Try
                        DecryptedString_ = System.Text.Encoding.UTF8.GetString(ms.ToArray)
                    Catch ex As Exception
                        DecryptedString_ = ""
                    End Try
                End Using
                AES.Clear()
            End Using

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try

    End Sub
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

    Public Shared Function EncodeString(ByVal StringToEncode As String, ByVal EncodingKey As String) As String ' ehemals: Verschluesseln
        Dim Ausgabe As String = StringToEncode

        Try
            If Not String.IsNullOrWhiteSpace(StringToEncode) AndAlso Not String.IsNullOrWhiteSpace(EncodingKey) Then
                Using rd As New RijndaelManaged
                    Using md5 As New MD5CryptoServiceProvider
                        Dim key() As Byte = md5.ComputeHash(Encoding.UTF8.GetBytes(EncodingKey)) ' das hier
                        rd.Key = key
                    End Using
                    rd.GenerateIV()
                    Dim iv() As Byte = rd.IV
                    Using ms As New MemoryStream
                        ms.Write(iv, 0, iv.Length)
                        Using cs As New CryptoStream(ms, rd.CreateEncryptor, CryptoStreamMode.Write)
                            Dim data() As Byte = System.Text.Encoding.UTF8.GetBytes(StringToEncode) ' mit dem getauscht
                            cs.Write(data, 0, data.Length)
                            cs.FlushFinalBlock()
                            Dim encdata() As Byte = ms.ToArray()
                            Ausgabe = System.Convert.ToBase64String(encdata)
                        End Using
                    End Using
                End Using
            End If

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return Ausgabe
    End Function
    Public Shared Function DecodeString(ByVal StringToDecode As String, ByVal DecodingKey As String) As String ' ehemals: Entschluesseln
        Dim rijndaelIvLength As Integer = 16
        Dim Ausgabe As String = StringToDecode

        Try
            If Not String.IsNullOrWhiteSpace(StringToDecode) AndAlso Not String.IsNullOrWhiteSpace(DecodingKey) AndAlso StringToDecode.Length >= 16 Then
                Using rd As New RijndaelManaged
                    Using md5 As New MD5CryptoServiceProvider
                        Dim key() As Byte = md5.ComputeHash(Encoding.UTF8.GetBytes(DecodingKey))
                        rd.Key = key
                    End Using
                    Dim encdata() As Byte = System.Convert.FromBase64String(StringToDecode)
                    Using ms As New MemoryStream(encdata)
                        Dim iv(15) As Byte
                        ms.Read(iv, 0, rijndaelIvLength)
                        rd.IV = iv
                        Using cs As New CryptoStream(ms, rd.CreateDecryptor, CryptoStreamMode.Read)
                            Dim data(CInt(ms.Length) - rijndaelIvLength) As Byte
                            Dim i As Integer = cs.Read(data, 0, data.Length) '############HIER#############
                            Ausgabe = System.Text.Encoding.UTF8.GetString(data, 0, i)
                        End Using
                    End Using
                End Using
            End If

        Catch ex As Exception
            Call Helper_ErrorHandling.HandleErrorCatch(ex, Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return Ausgabe
    End Function



#Region "SignPDF"

    'Private _privateSignature As IExternalSignature
    'Private _signChain As Org.BouncyCastle.X509.X509Certificate()



    'Public Shared Function GetPkcs12Store(ByVal cert As X509Certificate2, ByVal password As String) As Pkcs12Store
    '    Dim bcCert As Org.BouncyCastle.X509.X509Certificate
    '    bcCert = DotNetUtilities.FromX509Certificate(cert)
    '    Dim pk12 As Pkcs12Store = New Pkcs12StoreBuilder().Build()
    '    Dim certEntry As X509CertificateEntry = New X509CertificateEntry(bcCert)
    '    pk12.SetCertificateEntry(bcCert.SubjectDN.ToString(), certEntry)
    '    Dim keyEntry As AsymmetricKeyEntry = New AsymmetricKeyEntry(bcCert.GetPublicKey())
    '    Return pk12
    'End Function


    ''Private Shared ReadOnly SRC As String = "hello.pdf"
    ''Private Shared ReadOnly DOC_TO_SIGN As String = "signatureAddedUsingAppearance.pdf"
    ''Private Shared ReadOnly DEST As String = "signSignatureAddedUsingAppearance.pdf"
    ''Private Shared ReadOnly SIGNATURE_NAME As String = "Signature1"
    ''Private Shared ReadOnly PASSWORD As Char() = "password".ToCharArray()
    ''Private Shared ReadOnly CERT_PATH As String = "YOUR_CERTIFICATE_PATH"

    ''Public Shared Sub Main(ByVal args As String())
    ''    Dim file = New FileInfo(DEST)
    ''    file.Directory?.Create()
    ''    Call  CreateAndSignSignature(SRC, DOC_TO_SIGN, DEST, SIGNATURE_NAME)
    ''End Sub

    'Public Shared Sub CreateAndSignSignature(ByVal src As String,
    '                                         ByVal doc As String,
    '                                         ByVal dest As String,
    '                                         ByVal signatureName As String,
    '                                         ByVal CERT_PATH As String,
    '                                         ByVal PASSWORD As Char())

    '    Try
    '        AddSignatureToTheDocument(src, dest, signatureName)
    '        Dim FileInfoIn As New FileInfo(doc)
    '        Dim FileInfoOut As New FileInfo(dest)
    '        Dim pr As PdfReader = New PdfReader(New FileStream(doc, FileMode.Open)
    '        Dim padesSigner = New PdfPadesSigner(pr, New FileStream(dest, FileMode.Create)
    '        Dim CertChain As Org.BouncyCastle.X509.X509Certificate() = CType(GetCertificateChain(CERT_PATH, PASSWORD), X509Certificate())
    '        Dim signerProperties = CreateSignerProperties(signatureName)
    '        padesSigner.SignWithBaselineBProfile(signerProperties,
    '                                             CertChain,
    '                                             GetPrivateKey(CERT_PATH, PASSWORD))

    '    Catch ex As Exception
    '        Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
    '        If Helper.IsIDE() Then Stop
    '    End Try
    'End Sub

    'Private Shared Sub AddSignatureToTheDocument(ByVal src As String, ByVal dest As String, ByVal signatureName As String)
    '    Try
    '        Dim document As Document = New Document(New PdfDocument(New PdfReader(src), New PdfWriter(dest)))
    '        'Dim document = New Document(New PdfDocument(New PdfWriter(dest)))
    '        Dim table = New Table(2)
    '        Dim cell = New Cell(0, 2).Add(New Paragraph("Test signature").SetFontColor(ColorConstants.WHITE))
    '        cell.SetBackgroundColor(ColorConstants.GREEN)
    '        table.AddCell(cell)
    '        cell = New Cell().Add(New Paragraph("Signer"))
    '        cell.SetBackgroundColor(ColorConstants.LIGHT_GRAY)
    '        table.AddCell(cell)
    '        cell = New Cell().Add(New SignatureFieldAppearance(signatureName).SetContent("Sign here").SetHeight(50).SetWidth(100).SetInteractive(True))
    '        table.AddCell(cell)
    '        document.Add(table)
    '        document.Close()

    '    Catch ex As Exception
    '        Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
    '        If Helper.IsIDE() Then Stop
    '    End Try
    'End Sub

    'Private Shared Function CreateSignerProperties(ByVal signatureName As String) As SignerProperties
    '    Dim signerProperties As New SignerProperties
    '    Try
    '        signerProperties = New SignerProperties().SetFieldName(signatureName)
    '        Dim appearance = New SignatureFieldAppearance(signerProperties.GetFieldName()).SetContent("Signer", "Signature description. " & "Signer is replaced to the one from the certificate.").SetBackgroundColor(ColorConstants.YELLOW)
    '        signerProperties.SetSignatureAppearance(appearance).SetReason("Reason").SetLocation("Location")

    '    Catch ex As Exception
    '        Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
    '        If Helper.IsIDE() Then Stop
    '    End Try
    '    Return signerProperties
    'End Function

    'Private Shared Function GetPrivateKey(ByVal certificatePath As String, ByVal password As Char()) As PrivateKeySignature
    '    Dim pk As PrivateKey
    '    Try
    '        Dim [alias] As String = Nothing
    '        Dim pk12 = New Pkcs12StoreBuilder().Build()
    '        pk12.Load(New FileStream(certificatePath, FileMode.Open, FileAccess.Read), password)

    '        For Each a In pk12.Aliases
    '            [alias] = (CStr(a))

    '            If pk12.IsKeyEntry([alias]) Then
    '                Exit For
    '            End If
    '        Next

    '        pk = TryCast(pk12.GetKey([alias]).Key, IPrivateKey)
    '        'Dim pk1 As IPrivateKey = New PrivateKeyBC(pk12.GetKey([alias]).Key)

    '    Catch ex As Exception
    '        Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
    '        If Helper.IsIDE() Then Stop
    '    End Try
    '    Return New PrivateKeySignature(pk, DigestAlgorithms.SHA512)


    'End Function

    'Private Shared Function GetCertificateChain(ByVal certificatePath As String, ByVal password As Char()) As Org.BouncyCastle.X509.X509Certificate()
    '    Dim chain As Org.BouncyCastle.X509.X509Certificate()
    '    '.IX509Certificate()
    '    Try
    '        Dim [alias] As String = Nothing
    '        Dim pk12 = New Pkcs12StoreBuilder().Build()
    '        pk12.Load(New FileStream(certificatePath, FileMode.Open, FileAccess.Read), password)
    '        For Each a In pk12.Aliases
    '            [alias] = (CStr(a))
    '            If pk12.IsKeyEntry([alias]) Then
    '                Exit For
    '            End If
    '        Next

    '        Dim ce As X509CertificateEntry() = pk12.GetCertificateChain([alias])
    '        'Dim chain0 = New IX509Certificate(ce.Length - 1) {}
    '        chain = New Org.BouncyCastle.X509.X509Certificate(ce.Length - 1) {}
    '        Dim chain1 As Org.BouncyCastle.X509.X509Certificate() = New Org.BouncyCastle.X509.X509Certificate(ce.Length - 1) {}
    '        For k = 0 To ce.Length - 1
    '            'chain(k) = CType(ce(k).Certificate, iX509Certificate)
    '            chain(k) = ce(k).Certificate
    '        Next
    '        'chain = trycast(chain, iX509Certificate)
    '        Return chain

    '    Catch ex As Exception
    '        Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
    '        If Helper.IsIDE() Then Stop
    '        Return Nothing
    '    End Try
    'End Function

#End Region
#End Region

#Region "Eigenschaften"
    Public Shared ReadOnly Property EncryptedString() As String
        Get
            Return EncryptedString_
        End Get
    End Property
    Public Shared ReadOnly Property DecryptedString() As String
        Get
            Return DecryptedString_
        End Get
    End Property
    Public Shared Function CreateMD5Sum(ByVal filename As String) As String
        Dim md5checksum As String = ""
        Using md5 As MD5 = MD5.Create()
            Using stream As FileStream = File.OpenRead(filename)
                md5checksum = BitConverter.ToString(md5.ComputeHash(stream)).Replace("-", String.Empty)
            End Using
        End Using
        Return md5checksum
    End Function
#End Region

End Class