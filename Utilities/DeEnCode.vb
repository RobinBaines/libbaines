'------------------------------------------------
'Name: Module 'DeEnCode.vb
'Function: 
'Copyright Baines 2012. All rights reserved.
'Notes:  'See Walkthrough: Encrypting and Decrypting Strings in Visual Basic
'Modifications: 
'This does not solve the basic problem of sql authentication which is that the password has to be somewhere.
'Convert a plain text password into a encrypted string in a file called cipherText.txt using a password
'and the default key, see below. Do with the EnCode application. 
'Then put the encrypted string in the code and decrypt from there.
'So if robin2012 is encrypted to "KW6xdELlM57NgAAMR3psE5sh6/RkdJ1o" using the default key do the following to get it back
'Private Function GetPassword() As String

'    'The sql password was encrypted using the default key in Utilities.DeEnCode(), see parameter below. 
'    'It is decrypted using the default key in Utilities.DeEnCode().
'    Dim deCode As New Utilities.DeEnCode()
'    Return (deCode.DecryptData("KW6xdELlM57NgAAMR3psE5sh6/RkdJ1o"))
'End Function

'So if at any time in the future the password is changed and there is no reason to recompile
'use Utilities.EnCode application to create a cijpherTest.txt file and put this in together with the exe.
'If re- compiling do the same but copy the encrypted password from Utilities.EnCode cijpherTest.txt file to the parameter in the Main 
'form of the application.
Imports System.Security.Cryptography
Public NotInheritable Class DeEnCode
    Dim TripleDes As New TripleDESCryptoServiceProvider
    Dim cipherText As String = ""
    'default key
    Dim key As String = "5AvlfEGHWAYM98QU94Tc"
    Sub New(ByVal key As String)
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

    Sub New()
        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

    'Look for the cipher in a file. If it's not there set it from default.
    'This allows a cipher file to be placed should the password change.
    Sub New(FileName As String, DefaultcipherText As String)
        If My.Computer.FileSystem.FileExists(FileName) Then
            cipherText = My.Computer.FileSystem.ReadAllText(FileName)
        End If
        If cipherText = "" Then cipherText = DefaultcipherText

        ' Initialize the crypto provider.
        TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

    Private Function TruncateHash(
    ByVal key As String,
    ByVal length As Integer) As Byte()

        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the key. 
        Dim keyBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(key)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash. 
        ReDim Preserve hash(length - 1)
        Return hash
    End Function
    Public Function EncryptData(ByVal plaintext As String) As String

        ' Convert the plaintext string to a byte array. 
        Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plaintext)

        ' Create the stream. 
        Dim ms As New System.IO.MemoryStream

        ' Create the encoder to write to the stream. 
        Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string. 
        Return Convert.ToBase64String(ms.ToArray)
    End Function
    Public Function DecryptData()
        Return DecryptData(cipherText)
    End Function
    Public Function DecryptData(ByVal encryptedtext As String) As String

        ' Convert the encrypted text string to a byte array. 
        Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

        ' Create the stream. 
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream. 
        Dim decStream As New CryptoStream(ms,
            TripleDes.CreateDecryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
        decStream.FlushFinalBlock()

        ' Convert the plaintext stream to a string. 
        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
    End Function

End Class
