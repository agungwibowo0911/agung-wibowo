Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports AddOnSO_Close.Common
Imports Microsoft.VisualBasic
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports SAPbobsCOM

Public Class AuthoCloseSO
    Private objCryptoFunction As New Cryptography.CryptoControl
    Public oCompany As SAPbobsCOM.Company
    Public strSAPDB As String
    Public lRetCode As Integer
    Public sErrMsg As String
    Public lErrCode As Integer

    Dim GenerateAutoSOClose As New GenerateAutoCloseSO

    Private pstrSBOConnection As String = objCryptoFunction.RsaDynamicDecryption(ConfigurationSettings.AppSettings("pstrSBOConnection").ToString, ConfigurationSettings.AppSettings("PrivateKeySBO").ToString, Cryptography.CryptoControl.DynamicEncrypt.Symmetric)


    Public Function ConnectionSbo(ByVal pstrUserSAP As String, ByVal pstrPassSAP As String, ByVal dbName As String) As Boolean

        Dim sConnectionString As String = System.Convert.ToString(Environment.GetCommandLineArgs.GetValue(1))
        Dim strSplitCon As String()
        Dim strserver As String()
        Dim strUser As String()
        Dim strPassword As String()
        Dim StrDB As String()

        ConnectionSbo = True

        strSplitCon = Split(pstrSBOConnection, ";")

        strserver = Split(strSplitCon(0), "=")
        StrDB = Split(strSplitCon(1), "=")
        strUser = Split(strSplitCon(2), "=")
        strPassword = Split(strSplitCon(3), "=")
        ' MsgBox("" & Split(pstrSBOConnection, ";") & "")
        oCompany = New SAPbobsCOM.Company

        ' Init Connection Properties
        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
        oCompany.Server = RTrim(strserver(1)) ' change to your company server"
        oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English ' change to your language
        oCompany.DbUserName = RTrim(strUser(1))
        oCompany.DbPassword = RTrim(strPassword(1))

        strSAPDB = RTrim(StrDB(1))
        oCompany.CompanyDB = strSAPDB

        oCompany.UserName = pstrUserSAP
        oCompany.Password = pstrPassSAP

        lRetCode = oCompany.Connect()

        If lRetCode <> 0 Then ' if the connection failed
            ConnectionSbo = False
            oCompany.GetLastError(lErrCode, sErrMsg)
            MsgBox("Connection SAP Failed - " & sErrMsg, MsgBoxStyle.Exclamation, "Default Connection Failed")
        End If

    End Function

    Private Sub AuthoCloseSO_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim sConnectionString As String = System.Convert.ToString(Environment.GetCommandLineArgs.GetValue(1))
        Dim objdataUsr As New DataAccess
        Dim t As Timer = New Timer()

        Try
            Dim Connection As String = ""
            Dim DocNum As String = ""
            Dim Flag As String = "N"
            Dim TransType As String = ""
            Dim i As Integer

            Dim ConnectionLog As String = ""
            Dim User As String = ""
            Dim Password As String = ""
            Dim DtvUser As DataView = objdataUsr.UsrsapB1(pstrSBOConnection)

            For i = 0 To DtvUser.Count - 1
                ConnectionLog = DtvUser.Item(i).Item("dbName").ToString.Trim()
                User = DtvUser.Item(i).Item("User").ToString.Trim()
                Password = DtvUser.Item(i).Item("Password").ToString.Trim()

                ConnectionSbo(User, Password, ConnectionLog)

                'GenerateOutgoing.CancelOutgoing(oCompany)
                'GenerateIncoming.CancelIncoming(oCompany)
                GenerateAutoSOClose.CLoseDocumentSO(oCompany)

            Next


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class
