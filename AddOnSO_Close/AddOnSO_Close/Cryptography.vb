#Region "Summary"
'-----------------------------------------------------------------------------------------
'  Author	    : Vinno Ramadian Tedjo
'  Create Date	: 23 Januari 2006
'  Purpose	    : Encryption for Login and Connection string
'  Special Note	: This encryption using RSA and Hash methods from Microsoft .Net.
'                 Thanks for tutorial how to use this methods to 
'                 Wei-Meng Lee an Microsoft .NET MVP and co-founder of Active Developer, 
'                 a training company specializing in .NET and wireless technologies
' ========================================================================================
'  History	    : 
'  Release	    : 
'-----------------------------------------------------------------------------------------
#End Region

#Region "Option"
Option Explicit On
Option Strict On
#End Region

#Region "Imports"
Imports Microsoft.VisualBasic
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
#End Region

#Region "Please Readme"
'---------------------------------------------------------------
'How to Use
'---------------------------------------------------------------
'1. HashEncryption digunakan apabaila anda 
'   hanya ingin membandingkan dua nilai yang sudah di encrypt
'2. RSAEncryption dan RSADecryption digunakan apabila anda 
'   ingin mengirimkan suatu informasi yang sangat penting
'   untuk anda bandingkan atau anda simpan
'---------------------------------------------------------------
#End Region

#Region "Namespace CMNPCommon.Cryptography"
Namespace Common.Cryptography
#Region "Class enkripsi"
    Public Class CryptoControl

#Region "Private Variables"
        Private Shared RSA As RSACryptoServiceProvider
        Private Shared strPublicKey As String = String.Empty
        Private Shared strPrivateKey As String = String.Empty
#End Region

#Region "Public Methods"
        Public Enum DynamicEncrypt
            Symmetric
            Asymmetric
        End Enum

        Public Function HashEncryption(ByVal strData As String) As String
            Try
                'Convert the string to ASCII
                'Return the value
                Return HashEncrypt(Encoding.ASCII.GetBytes(strData))
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Function RsaStaticEncryption(ByVal strData As String) As String
            'Get the encryption string using RSA Symmetric Encryption Methods
            'And Return the value
            Return SymmetricEncryption(strData, StaticKey, StaticKey)
        End Function

        Public Function RsaDynamicEncryption(ByVal strData As String, ByVal strEncryption As DynamicEncrypt) As String()
            Select Case strEncryption
                Case DynamicEncrypt.Symmetric
                    'Get the private and public key
                    Dim strTemp As String = Convert.ToBase64String(RandomByte())
                    strPublicKey = strTemp
                    strPrivateKey = strTemp

                    'Get the encryption string using RSA Symmetric Encryption Methods
                    'And Return the value
                    Dim strTempEncryption() As String
                    ReDim strTempEncryption(2)
                    strTempEncryption(0) = SymmetricEncryption(strData, strPublicKey, strPrivateKey)
                    strTempEncryption(1) = strPrivateKey

                    Return strTempEncryption
                Case DynamicEncrypt.Asymmetric
                    'Creates a new instance of RSACryptoServiceProvider
                    RSA = New RSACryptoServiceProvider()

                    'Get the Public Key & Private Key
                    strPublicKey = RSA.ToXmlString(False)
                    strPrivateKey = RSA.ToXmlString(True)


                    'Get the encryption string using RSA Asymmetric Encryption Methods
                    'And Return the value
                    Dim bytTemp() As Byte = Encoding.ASCII.GetBytes(strPrivateKey)
                    Dim strTemp As String = Convert.ToBase64String(bytTemp)
                    strPrivateKey = strTemp

                    Dim strTempEncryption() As String
                    ReDim strTempEncryption(2)
                    strTempEncryption(0) = AsymmetricEncryption(strData, strPublicKey)
                    strTempEncryption(1) = strPrivateKey

                    Return strTempEncryption
                Case Else
                    Return Nothing
            End Select
        End Function

        Public Function RsaStaticDecryption(ByVal strData As String) As String
            'Get the decryption string using RSA Symmetric Decryption Methods
            'And Return the value
            Return SymmetricDecryption(strData, StaticKey, StaticKey)
        End Function

        Public Function RsaDynamicDecryption(ByVal strData As String, ByVal strKey As String, ByVal strEncryption As DynamicEncrypt) As String
            Select Case strEncryption
                Case DynamicEncrypt.Symmetric
                    'Get the decryption string using RSA Asymmetric Encryption Methods
                    'And Return the value
                    Return SymmetricDecryption(strData, strKey, strKey)
                Case DynamicEncrypt.Asymmetric
                    'Convert Private Key
                    Dim bytTemp() As Byte = Convert.FromBase64String(strKey)
                    Dim strTemp As String = Encoding.ASCII.GetString(bytTemp)
                    strPrivateKey = strTemp

                    'Get the decryption string using RSA Asymmetric Encryption Methods
                    'And Return the value
                    Return AsymmetricDecryption(strData, strPrivateKey)
                Case Else
                    Return Nothing
            End Select
        End Function
#End Region

#Region "Private Methods"
        Private Function HashEncrypt(ByVal bytData() As Byte) As String
            Try
                'Declare Hash algoritma
                Dim hashAlgoritma As New SHA1CryptoServiceProvider()
                'Compute data type byte using hash algoritma
                Dim hashValue() As Byte = hashAlgoritma.ComputeHash(bytData)
                'return the value
                'HashEncrypt = ByteArrayToString(hashValue)
                HashEncrypt = Convert.ToBase64String(hashValue)
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Private Function AsymmetricEncryption(ByVal strData As String, ByVal strPublicKey As String) As String
            Try
                'Creates a new instance of RSACryptoServiceProvider
                RSA = New RSACryptoServiceProvider()

                'Loads the public key
                RSA.FromXmlString(strPublicKey)
                Dim EncryptedStr() As Byte

                'Encrypts the string
                EncryptedStr = RSA.Encrypt(Encoding.ASCII.GetBytes(strData), False)

                'Converts the encrypted byte array to string
                'Dim i As Integer
                'Dim strBuild As New StringBuilder()
                'For i = 0 To EncryptedStr.Length - 1
                '    If i <> EncryptedStr.Length - 1 Then
                '        strBuild.Append(EncryptedStr(i) & " ")
                '    Else
                '        strBuild.Append(EncryptedStr(i))
                '    End If
                'Next

                'Return the value
                'AsymmetricEncryption = ByteArrayToString(EncryptedStr) 'strBuild.ToString()

                AsymmetricEncryption = Convert.ToBase64String(EncryptedStr)
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Private Function AsymmetricDecryption(ByVal strData As String, ByVal strPrivateKey As String) As String
            Try
                'Creates a new instance of RSACryptoServiceProvider
                RSA = New RSACryptoServiceProvider()

                'Loads the private key
                RSA.FromXmlString(strPrivateKey)

                'Decrypts the string
                'Dim DecryptedStr As Byte() = RSA.Decrypt(StringToByteArray(strData), False)
                Dim DecryptedStr As Byte() = RSA.Decrypt(Convert.FromBase64String(strData), False)

                'Converts the decrypted byte array to string
                Dim i As Integer
                Dim strBuild As New StringBuilder()
                For i = 0 To DecryptedStr.Length - 1
                    strBuild.Append(Convert.ToChar(DecryptedStr(i)))
                Next
                AsymmetricDecryption = strBuild.ToString()
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Private Function SymmetricDecryption(ByVal strData As String, ByVal strPublicKey As String, ByVal strPrivateKey As String) As String
            Try
                Dim strTemp As String = String.Empty

                'converts the encrypted string into a byte array
                'Dim bytTemp As Byte() = stringToByteArray(strData)
                Dim bytTemp As Byte() = Convert.FromBase64String(strData)

                'converts the publickey and privatekey from string to byte array
                Dim bytPublicKey As Byte() = Encoding.UTF8.GetBytes(strPublicKey)
                Dim bytPrivateKey As Byte() = Encoding.UTF8.GetBytes(strPrivateKey)

                'converts the byte array into a memory stream for decryption
                Dim memStream As New MemoryStream(bytTemp)
                Dim RMCrypto As New RijndaelManaged()
                Dim CryptStream As New CryptoStream(memStream, RMCrypto.CreateDecryptor(bytPublicKey, bytPrivateKey), CryptoStreamMode.Read)

                'decrypting the stream
                Dim SReader As New StreamReader(CryptStream)
                strTemp = SReader.ReadToEnd
                SReader.Close()

                'converts the descrypted stream into a string
                strTemp.ToString()
                SymmetricDecryption = strTemp
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Private Function SymmetricEncryption(ByVal strData As String, ByVal strPublicKey As String, ByVal strPrivateKey As String) As String
            Try
                Dim memStream As New MemoryStream()

                'converts the publickey and privatekey from string to byte array
                Dim bytPublickey As Byte() = System.Text.Encoding.UTF8.GetBytes(strPublicKey)
                Dim bytPrivateKey As Byte() = System.Text.Encoding.UTF8.GetBytes(strPrivateKey)

                'creates a new instance of the RijndaelManaged class
                Dim RMCrypto As New RijndaelManaged()

                'creates a new instance of the CryptoStream class
                Dim CryptStream As New CryptoStream(memStream, RMCrypto.CreateEncryptor(bytPublickey, bytPrivateKey), CryptoStreamMode.Write)
                Dim SWriter As New StreamWriter(CryptStream)

                'encrypting the string
                SWriter.Write(strData)
                SWriter.Close()
                CryptStream.Close()

                'converts the encrypted stream into a string
                'Dim strBuild As New System.Text.StringBuilder()
                'Dim bytTemp As Byte() = memStream.ToArray
                'Dim i As Integer
                'For i = 0 To bytTemp.Length - 1
                '    If i <> bytTemp.Length - 1 Then
                '        strBuild.Append(bytTemp(i) & " ")
                '    Else
                '        strBuild.Append(bytTemp(i))
                '    End If
                'Next
                'SymmetricEncryption = ByteArrayToString(memStream.ToArray)
                SymmetricEncryption = Convert.ToBase64String(memStream.ToArray)
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Private Function StringToByteArray(ByVal strData As String) As Byte()
            'e.g. "1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16" to   
            '     {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16}
            Try
                Dim strTemp As String()
                strTemp = strData.Split(" ".ToCharArray)
                Dim bytTemp(strTemp.Length - 1) As Byte
                Dim i As Integer
                For i = 0 To strTemp.Length - 1
                    bytTemp(i) = Convert.ToByte(strTemp(i))
                Next
                StringToByteArray = bytTemp
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Private Function ByteArrayToString(ByVal bytData As Byte()) As String
            'e.g. {1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16} to   
            '     "1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16"

            Dim i As Integer
            Dim strBuild As New StringBuilder()
            For i = 0 To bytData.Length - 1
                If i <> bytData.Length - 1 Then
                    strBuild.Append(bytData(i) & " ")
                Else
                    strBuild.Append(bytData(i))
                End If
            Next

            ByteArrayToString = strBuild.ToString
        End Function

        Private Function RandomByte() As Byte()
            Dim bytTemp(10) As Byte
            Dim rng As New RNGCryptoServiceProvider
            rng.GetBytes(bytTemp)
            RandomByte = bytTemp
        End Function

        Private Function StaticKey() As String
            StaticKey = "brdW1riXSuru/v4="
        End Function

#End Region

    End Class
#End Region

End Namespace
#End Region

