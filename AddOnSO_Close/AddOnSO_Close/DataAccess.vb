Option Strict Off
Option Explicit On
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System
Public Class DataAccess
    Private objTrans As New Common.SQLFunction
    Private strodbc As String

    Private Function SQLGetDataView(ByVal strSPName As String, ByVal parameter As SqlParameter(), ByVal scnMain As SqlConnection, ByVal stnMain As SqlTransaction) As DataView
        'agung wibowo
        Dim scmMain As SqlCommand = New SqlCommand()
        Dim sdaMain As SqlDataAdapter = New SqlDataAdapter()
        Dim dttMain As DataTable = New DataTable("MyDataTable")

        If scmMain.Parameters.Count > 0 Then
            scmMain.Parameters.Clear()
        End If

        scmMain.CommandText = strSPName
        scmMain.Connection = scnMain
        scmMain.Transaction = stnMain
        scmMain.CommandTimeout = 3000
        scmMain.CommandType = CommandType.StoredProcedure

        Try
            If Not (parameter Is Nothing) Then
                For Each prm As SqlParameter In parameter
                    scmMain.Parameters.Add(prm)
                Next
            End If

            sdaMain.SelectCommand = scmMain
            sdaMain.Fill(dttMain)
        Catch expodbc As SqlException
            Throw expodbc
        Catch expSystem As Exception
            Throw expSystem
        End Try

        Return dttMain.DefaultView()
    End Function

    Public Function UsrsapB1(ByVal pstrConnection As String) As DataView
        Dim strSQL As String = String.Empty
        Dim strodbc As String = String.Empty
        Dim strSplitCon As String()
        Dim StrDB As String()
        Dim SAPDB As String
        strSplitCon = Split(pstrConnection, ";")
        StrDB = Split(strSplitCon(1), "=")
        SAPDB = RTrim(StrDB(1))

        strodbc = "EXEC [MIS_FL_GetUsrLogin]  'FL '"

        'return value
        Dim dtv3 As DataView = objTrans.SQLGetDataView(pstrConnection, strodbc)
        Return dtv3
    End Function

#Region "FUNCTION"

    Public Function fctFormatDateSave(ByVal oCompany As SAPbobsCOM.Company, ByVal pdate As String, ByVal sngFormat As Integer) As String
        Dim strFormat As String
        Dim strMonth As String
        Dim intLength As Integer
        Static oGetCompanyService As SAPbobsCOM.CompanyService = Nothing
        Dim oAdminInfo As SAPbobsCOM.AdminInfo = Nothing

        On Error GoTo ErrorHandler

        strMonth = "JANUARY01FEBRUARY02MARCH03APRIL04MAY05JUNE06JULY07AUGUST08SEPTEMBER09OCTOBER10NOVEMBER11DECEMBER12"

        oGetCompanyService = oCompany.GetCompanyService
        oAdminInfo = oGetCompanyService.GetAdminInfo

        sngFormat = oAdminInfo.DateTemplate

        If pdate = "" Then
            GoTo ErrorHandler
        End If

        Select Case sngFormat
            Case 0
                fctFormatDateSave = "20" + Right(pdate, 2) + "/" + Mid(pdate, 4, 2) + "/" + Left(pdate, 2)
            Case 1
                fctFormatDateSave = Right(pdate, 4) + "/" + Mid(pdate, 4, 2) + "/" + Left(pdate, 2)
            Case 2
                fctFormatDateSave = "20" + Right(pdate, 2) + "/" + Left(pdate, 2) + "/" + Mid(pdate, 4, 2)
            Case 3
                fctFormatDateSave = Right(pdate, 4) + "/" + Left(pdate, 2) + "/" + Mid(pdate, 4, 2)
            Case 4
                fctFormatDateSave = Left(pdate, 4) + "/" + Mid(pdate, 6, 2) + "/" + Right(pdate, 2)
            Case 5
                intLength = InStr(1, strMonth, UCase(Mid(pdate, 4, Len(pdate) - 8))) + Len(Mid(pdate, 4, Len(pdate) - 8))
                fctFormatDateSave = Right(pdate, 4) + "/" + Mid(strMonth, intLength, 2) + "/" + Left(pdate, 2)
        End Select

        GoTo SetNothing

ErrorHandler:
        fctFormatDateSave = ""

SetNothing:
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oGetCompanyService)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oAdminInfo)
        oGetCompanyService = Nothing
        oAdminInfo = Nothing
    End Function

#End Region

End Class
