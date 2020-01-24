Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports AddOnSO_Close.Common
Imports Microsoft.VisualBasic
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports SAPbobsCOM

Public Class GenerateAutoCloseSO
    Private objCryptoFunction As New Cryptography.CryptoControl
    Public lRetCode As Integer
    Public sErrMsg As String
    Public lErrCode As Integer
    Public strSAPDB As String
    Dim objRecSetConn As SAPbobsCOM.Recordset
    Dim pstrMaxDay As String
    Dim BlStatus As Boolean = True

    Private pstrSBOConnection As String = objCryptoFunction.RsaDynamicDecryption(ConfigurationSettings.AppSettings("pstrSBOConnection").ToString, ConfigurationSettings.AppSettings("PrivateKeySBO").ToString, Cryptography.CryptoControl.DynamicEncrypt.Symmetric)

    Public Function CLoseDocumentSO(ByRef oCompany As SAPbobsCOM.Company) As Boolean
        Dim SO As SAPbobsCOM.Documents = Nothing
        Dim lngResult As Long
        Dim ErrCode As Long
        Dim sql As String
        Dim RS As SAPbobsCOM.Recordset
        'Dim ErrMsg As String
        RS = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            If Not oCompany.InTransaction Then
                oCompany.StartTransaction()
            End If

            SO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

         
            sql = "EXEC [MIS_AddOn_AutoCloseSO]"
            RS.DoQuery(sql)
            ' RS.MoveFirst()
            For i = 1 To RS.RecordCount

                lngResult = SO.GetByKey(RS.Fields.Item("DocEntry").Value.ToString.Trim)

                'Close the record
                lngResult = SO.Close
                If oCompany.InTransaction Then Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                RS.MoveNext()
            Next


        Catch ex As Exception
            If lngResult <> 0 Then
                If oCompany.InTransaction Then Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                'Return Nothing
            End If
            If Err.Description <> "" Then
                If oCompany.InTransaction Then Call oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                'Return Nothing
            End If
        End Try


        System.Runtime.InteropServices.Marshal.ReleaseComObject(RS)
        RS = Nothing
        SO = Nothing
        Return Nothing
        
    End Function


End Class
