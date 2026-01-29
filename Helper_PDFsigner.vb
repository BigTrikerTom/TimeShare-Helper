Imports System
Imports System.IO
Imports System.Linq
'Imports iTextSharp.text.pdf
'Imports iTextSharp.text.pdf.security
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.Pkcs

Public Class Helper_PDFsigner


        'Private Shared Sub test(ByVal args As String())
        '    Dim pdfFilePath As String = "C:\Users\Admin\Documents\MyPDF.pdf"
        '    Dim pdfReader As PdfReader = New PdfReader(pdfFilePath)
        '    Dim pfxFilePath As String = "D:\Uday Dodiya\Digital_Sign\Uday Dodiya.pfx"
        '    Dim pfxPassword As String = "uday1234"
        'Dim pdfStamper As PdfStamper = PdfStamper.CreateSignature(pdfReader, New FileStream("C:\Users\Admin\Documents\MyPDF_Signed.pdf", FileMode.Create), CChar(vbNullChar), Nothing, True)
        'Dim signatureAppearance As PdfSignatureAppearance = pdfStamper.SignatureAppearance
        '    signatureAppearance.Reason = "Digital Signature Reason"
        '    signatureAppearance.Location = "Digital Signature Location"
        '    Dim x As Single = 360
        '    Dim y As Single = 130
        '    signatureAppearance.Acro6Layers = False
        '    signatureAppearance.Layer4Text = PdfSignatureAppearance.questionMark
        '    signatureAppearance.SetVisibleSignature(New iTextSharp.text.Rectangle(x, y, x + 150, y + 50), 1, "signature")
        'Dim pfxKeyStore As Org.BouncyCastle.Pkcs.Pkcs12Store(New FileStream(pfxFilePath, FileMode.Open, FileAccess.Read), pfxPassword.ToCharArray)
        'Dim [alias] As String = pfxKeyStore.Aliases.Cast(Of String)().FirstOrDefault(Function(entryAlias) pfxKeyStore.IsKeyEntry(entryAlias))

        'If [alias] IsNot Nothing Then
        '        Dim privateKey As ICipherParameters = pfxKeyStore.GetKey([alias]).Key
        '        Dim pks As IExternalSignature = New PrivateKeySignature(privateKey, DigestAlgorithms.SHA256)
        '        MakeSignature.SignDetached(signatureAppearance, pks, New Org.BouncyCastle.X509.X509Certificate() {pfxKeyStore.GetCertificate([alias]).Certificate}, Nothing, Nothing, Nothing, 0, CryptoStandard.CMS)
        '    Else
        '        Console.WriteLine("Private key not found in the PFX certificate.")
        '    End If

        '    pdfStamper.Close()
        '    Console.WriteLine("PDF signed successfully!")
        'End Sub



End Class
