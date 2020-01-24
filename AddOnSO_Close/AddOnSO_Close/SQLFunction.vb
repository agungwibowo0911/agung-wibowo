Option Strict On
Option Explicit On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices
Imports System.Data.Odbc
Namespace Common
    Public Class SQLFunction
        Public Function SQLReader(ByVal strSQL As String, ByVal pscnMain As SqlConnection) As SqlDataReader
            'define query
            Dim scmMain As New SqlCommand(strSQL, pscnMain)

            'get data into reader
            Dim sdr As SqlDataReader = scmMain.ExecuteReader

            'Return Value
            Return sdr
        End Function

        Public Function SQLReader(ByVal strSQL As String, ByVal pstrConnection As String) As SqlDataReader
            'Open Connection
            Dim pscnMain As New SqlConnection(pstrConnection)
            pscnMain.Open()

            'define query
            Dim scmMain As New SqlCommand(strSQL, pscnMain)

            'get data into reader
            Dim sdr As SqlDataReader = scmMain.ExecuteReader

            'Return Value
            Return sdr
        End Function

        Public Sub CloseSqlReader(ByVal pscnMain As SqlConnection, ByVal sdr As SqlDataReader)
            sdr.Close()
            subCleanSQL(pscnMain)
        End Sub

        Public Sub CloseSqlReader(ByVal pscnMain As SqlConnection, ByVal sdr As SqlDataReader, ByVal pscmMain As SqlCommand)
            sdr.Close()
            subCleanSQL(pscnMain, pscmMain)
        End Sub

        Public Function SQLScalar(ByVal strSQL As String, ByVal pstrConnection As String) As Object

            'open connection
            Dim scnMain As New SqlConnection(pstrConnection)
            scnMain.Open()

            'define query
            Dim scmMain As New SqlCommand(strSQL, scnMain)

            'get data from scalar
            Dim scl As Object = scmMain.ExecuteScalar

            subCleanSQL(scnMain, scmMain)

            'Return Value
            SQLScalar = scl
        End Function

        Public Sub SQLSPTransactionCommand(ByVal strSPName As String, ByVal parameter As SqlParameter(), ByVal pstrConnection As String, Optional ByRef pscnMain As SqlConnection = Nothing, Optional ByRef pstnMain As SqlTransaction = Nothing)
            'create connection
            Dim scnMain As SqlConnection
            Dim stnMain As SqlTransaction

            'begin transaction
            If pscnMain Is Nothing Then
                scnMain = New SqlConnection(pstrConnection)
                scnMain.Open()
                stnMain = scnMain.BeginTransaction(IsolationLevel.Serializable)
            Else
                stnMain = pstnMain
                scnMain = pscnMain
            End If

            'create command for header
            Dim scmHeader As SqlCommand = New SqlCommand(strSPName, scnMain, stnMain)
            scmHeader.CommandType = CommandType.StoredProcedure
            scmHeader.CommandTimeout = 3600
            Try

                '  define parameters for detail 
                If Not parameter Is Nothing Then
                    Dim Param As SqlParameter
                    For Each Param In parameter
                        scmHeader.Parameters.Add(Param)
                    Next
                End If

                'execute 
                scmHeader.ExecuteNonQuery()

                'commit transaction
                If pscnMain Is Nothing Then stnMain.Commit()

            Catch expSQL As OdbcException
                If pscnMain Is Nothing Then stnMain.Rollback()
                Throw expSQL
            Catch expSystem As Exception
                If pscnMain Is Nothing Then stnMain.Rollback()
                Throw expSystem
            Finally
                'close connection
                If pscnMain Is Nothing Then subCleanSQL(scnMain, scmHeader)
            End Try
        End Sub

        Public Sub SQLTransactionCommand(ByVal strSQL As String, ByVal pstrConnection As String, Optional ByRef pscnMain As SqlConnection = Nothing, Optional ByRef pstnMain As SqlTransaction = Nothing)
            'create connection
            Dim scnMain As SqlConnection
            Dim stnMain As SqlTransaction

            'begin transaction
            If pscnMain Is Nothing Then
                scnMain = New SqlConnection(pstrConnection)
                scnMain.Open()
                stnMain = scnMain.BeginTransaction
            Else
                stnMain = pstnMain
                scnMain = pscnMain
            End If

            Dim scmMain As New SqlCommand(strSQL, scnMain, stnMain)
            scmMain.CommandType = CommandType.Text
            scmMain.CommandTimeout = 3600

            Try
                'execute command
                scmMain.ExecuteNonQuery()

                'commit transaction
                If pscnMain Is Nothing Then stnMain.Commit()

            Catch expodbc As SqlException
                If pscnMain Is Nothing Then stnMain.Rollback()
                Throw expodbc
            Catch expSystem As System.Exception
                If pscnMain Is Nothing Then stnMain.Rollback()
                Throw expSystem
            Finally
                'close connection
                If pscnMain Is Nothing Then subCleanSQL(scnMain, scmMain)
            End Try
        End Sub

        Public Function SQLSPGetDataView(ByVal strSPName As String, ByVal parameter As SqlParameter(), ByVal pstrConnection As String, Optional ByRef pscnMain As SqlConnection = Nothing, Optional ByRef pstnMain As SqlTransaction = Nothing) As DataView
            ' 'open connection
            Dim scnMain As New SqlConnection(pstrConnection)
            scnMain.Open()

            'define command           
            Dim scmHeader As New SqlCommand(strSPName, scnMain)
            scmHeader.CommandType = CommandType.StoredProcedure

            'define parameters
            If Not parameter Is Nothing Then
                Dim Param As SqlParameter
                For Each Param In parameter
                    scmHeader.Parameters.Add(Param)
                Next
            End If

            'create data adapter for the query
            Dim sdaMain As New SqlDataAdapter(scmHeader)

            'create data set and populate it with table
            Dim dstMain As New DataSet()
            sdaMain.Fill(dstMain, "MyTable")

            'return value
            Dim dvwItemClass As New DataView(dstMain.Tables("MyTable"), "", "", DataViewRowState.CurrentRows)

            subCleanSQL(scnMain, scmHeader)

            Return dvwItemClass
        End Function

        Public Function SQLGetDataView(ByRef pstrConnection As String, ByRef pstrSQL As String) As DataView
            Try
                Dim adcMain As New SqlConnection(pstrConnection)
                adcMain.Open()
                'create data adapter
                Dim sdaAccount As New SqlDataAdapter(pstrSQL, adcMain)

                'load data set
                Dim dstItemClass As New System.Data.DataSet
                sdaAccount.Fill(dstItemClass, "MyTable")

                'close connection
                subCleanSQL(adcMain)

                'return value
                Dim dvwItemClass As New System.Data.DataView(dstItemClass.Tables("MyTable"), "", "", DataViewRowState.CurrentRows)
                Return dvwItemClass
            Catch ex As Exception
                MsgBox("" & ex.Message & "")
            End Try

            Return Nothing
        End Function

        Public Function SQLGetDataTable(ByRef pstrConnection As String, ByRef pstrSQL As String) As DataTable
            Dim scn As New SqlConnection(pstrConnection)

            Dim scm As New SqlCommand
            scm.Connection = scn
            scm.CommandText = pstrSQL
            scm.CommandType = CommandType.Text

            'create data adapter
            Dim sdaAccount As New SqlDataAdapter(scm)

            'load data set
            Dim dtt As New DataTable("MyTable")
            sdaAccount.Fill(dtt)

            'return value
            Return dtt
        End Function

        Public Function SQLSPGetDataSet(ByVal strSPName As String, ByVal parameter As SqlParameter(), ByVal pstrConnection As String, Optional ByRef pscnMain As SqlConnection = Nothing, Optional ByRef pstnMain As SqlTransaction = Nothing) As DataSet
            ' 'open connection
            Dim scnMain As New SqlConnection(pstrConnection)
            scnMain.Open()

            'define command           
            Dim scmHeader As New SqlCommand(strSPName, scnMain)
            scmHeader.CommandType = CommandType.StoredProcedure
            scmHeader.CommandTimeout = 3600
            'define parameters
            If Not parameter Is Nothing Then
                Dim Param As SqlParameter
                For Each Param In parameter
                    scmHeader.Parameters.Add(Param)
                Next
            End If

            'create data adapter for the query
            Dim sdaMain As New SqlDataAdapter(scmHeader)

            'create data set and populate it with table
            Dim dstMain As New DataSet()
            sdaMain.Fill(dstMain)

            'return value
            Return dstMain
        End Function

        Public Function SQLGetDataSet(ByRef pstrConnection As String, ByRef pstrSQL As String) As DataSet
            Dim adcMain As New SqlConnection(pstrConnection)
            adcMain.Open()

            'create data adapter
            Dim sdaAccount As New SqlDataAdapter(pstrSQL, adcMain)

            'load data set
            Dim dstMain As New System.Data.DataSet
            sdaAccount.Fill(dstMain)

            'close connection
            subCleanSQL(adcMain)

            'return value
            Return dstMain
        End Function

        Public Sub subCleanSQL(ByRef pscnConnect As SqlConnection, Optional ByRef pcmCommand As SqlCommand = Nothing)
            'Clean Memory From scmCommand and scnConnection
            If Not pcmCommand Is Nothing Then
                pcmCommand.Dispose()
                pcmCommand = Nothing
            End If

            If Not pscnConnect Is Nothing Then
                If Not pscnConnect.State <> ConnectionState.Closed Then pscnConnect.Close()
                pscnConnect.Dispose()
                pscnConnect = Nothing
            End If
        End Sub
    End Class

    Public Class odbcFunction
        Public Function OdbcReader(ByVal strSQL As String, ByVal pscnMain As OdbcConnection) As OdbcDataReader
            'define query
            Dim scmMain As New OdbcCommand(strSQL, pscnMain)

            'get data into reader
            Dim sdr As OdbcDataReader = scmMain.ExecuteReader

            'Return Value
            Return sdr
        End Function

        Public Function OdbcReader(ByVal strSQL As String, ByVal pstrConnection As String) As OdbcDataReader
            'Open Connection
            Dim pscnMain As New OdbcConnection(pstrConnection)
            pscnMain.Open()

            'define query
            Dim scmMain As New OdbcCommand(strSQL, pscnMain)

            'get data into reader
            Dim sdr As OdbcDataReader = scmMain.ExecuteReader

            'Return Value
            Return sdr
        End Function

        Public Sub CloseOdbcReader(ByVal pscnMain As OdbcConnection, ByVal sdr As OdbcDataReader)
            sdr.Close()
            subClean(pscnMain)
        End Sub

        Public Sub CloseOdbcReader(ByVal pscnMain As OdbcConnection, ByVal sdr As OdbcDataReader, ByVal pscmMain As OdbcCommand)
            sdr.Close()
            subClean(pscnMain, pscmMain)
        End Sub

        Public Function odbcScalar(ByVal strSQL As String, ByVal pstrConnection As String) As Object

            'open connection
            Dim scnMain As New OdbcConnection(pstrConnection)
            scnMain.Open()

            'define query
            Dim scmMain As New OdbcCommand(strSQL, scnMain)
            scmMain.CommandTimeout = 3600

            'get data from scalar
            Dim scl As Object = scmMain.ExecuteScalar

            subClean(scnMain, scmMain)

            'Return Value
            odbcScalar = scl
        End Function

        Public Sub odbcTransactionCommand(ByVal strSQL As String, ByVal pstrConnection As String, Optional ByRef pscnMain As OdbcConnection = Nothing, Optional ByRef pstnMain As OdbcTransaction = Nothing)
            'create connection
            Dim scnMain As OdbcConnection
            Dim stnMain As OdbcTransaction

            'begin transaction
            If pscnMain Is Nothing Then
                scnMain = New OdbcConnection(pstrConnection)
                scnMain.Open()
                stnMain = scnMain.BeginTransaction
            Else
                stnMain = pstnMain
                scnMain = pscnMain
            End If

            Dim scmMain As New OdbcCommand(strSQL, scnMain, stnMain)
            scmMain.CommandType = CommandType.Text
            scmMain.CommandTimeout = 3600

            Try
                'execute command
                scmMain.ExecuteNonQuery()

                'commit transaction
                If pscnMain Is Nothing Then stnMain.Commit()

            Catch expodbc As OdbcException
                If pscnMain Is Nothing Then stnMain.Rollback()
                Throw expodbc
            Catch expSystem As System.Exception
                If pscnMain Is Nothing Then stnMain.Rollback()
                Throw expSystem
            Finally
                'close connection
                If pscnMain Is Nothing Then subClean(scnMain, scmMain)
            End Try
        End Sub

        Public Sub odbcSPTransactionCommand(ByVal strSPName As String, ByVal parameter As OdbcParameter(), ByVal pstrConnection As String, Optional ByRef pscnMain As OdbcConnection = Nothing, Optional ByRef pstnMain As OdbcTransaction = Nothing)
            'create connection
            Dim scnMain As OdbcConnection
            Dim stnMain As OdbcTransaction

            'begin transaction
            If pscnMain Is Nothing Then
                scnMain = New OdbcConnection(pstrConnection)
                scnMain.Open()
                stnMain = scnMain.BeginTransaction(IsolationLevel.Serializable)
            Else
                stnMain = pstnMain
                scnMain = pscnMain
            End If

            'create command for header
            Dim scmHeader As OdbcCommand = New OdbcCommand(strSPName, scnMain, stnMain)
            scmHeader.CommandType = CommandType.StoredProcedure
            scmHeader.CommandTimeout = 3600

            Try

                'define parameters for detail 
                If Not parameter Is Nothing Then
                    Dim Param As OdbcParameter
                    For Each Param In parameter
                        scmHeader.Parameters.Add(Param)
                    Next
                End If

                'execute 
                scmHeader.ExecuteNonQuery()

                'commit transaction
                If pscnMain Is Nothing Then stnMain.Commit()

            Catch expSQL As OdbcException
                If pscnMain Is Nothing Then stnMain.Rollback()
                Throw expSQL
            Catch expSystem As Exception
                If pscnMain Is Nothing Then stnMain.Rollback()
                Throw expSystem
            Finally
                'close connection
                If pscnMain Is Nothing Then subClean(scnMain, scmHeader)
            End Try
        End Sub

        Public Function OdbcSPGetDataView(ByVal strSPName As String, ByVal parameter As OdbcParameter(), ByVal pstrConnection As String, Optional ByRef pscnMain As OdbcConnection = Nothing, Optional ByRef pstnMain As OdbcTransaction = Nothing) As DataView
            ' 'open connection
            Dim scnMain As New OdbcConnection(pstrConnection)
            scnMain.Open()

            'define command           
            Dim scmHeader As New OdbcCommand(strSPName, scnMain)
            scmHeader.CommandType = CommandType.StoredProcedure

            'define parameters
            If Not parameter Is Nothing Then
                Dim Param As OdbcParameter
                For Each Param In parameter
                    scmHeader.Parameters.Add(Param)
                Next
            End If

            'create data adapter for the query
            Dim sdaMain As New OdbcDataAdapter(scmHeader)

            'create data set and populate it with table
            Dim dstMain As New DataSet()
            sdaMain.Fill(dstMain, "MyTable")

            'return value
            Dim dvwItemClass As New DataView(dstMain.Tables("MyTable"), "", "", DataViewRowState.CurrentRows)

            subClean(scnMain, scmHeader)

            Return dvwItemClass
        End Function

        Public Function SQLGetDataView(ByRef pstrConnection As String, ByRef pstrSQL As String) As DataView
            Dim adcMain As New SqlConnection(pstrConnection)
            adcMain.Open()

            'create data adapter
            Dim sdaAccount As New SqlDataAdapter(pstrSQL, adcMain)

            'load data set
            Dim dstItemClass As New System.Data.DataSet
            sdaAccount.Fill(dstItemClass, "MyTable")

            'close connection
            subClean(adcMain)

            'return value
            Dim dvwItemClass As New System.Data.DataView(dstItemClass.Tables("MyTable"), "", "", DataViewRowState.CurrentRows)
            Return dvwItemClass
        End Function

        Public Function OdbcGetDataTable(ByRef pstrConnection As String, ByRef pstrSQL As String) As DataTable
            Dim scn As New OdbcConnection(pstrConnection)

            Dim scm As New OdbcCommand
            scm.Connection = scn
            scm.CommandText = pstrSQL
            scm.CommandType = CommandType.Text

            'create data adapter
            Dim sdaAccount As New OdbcDataAdapter(scm)

            'load data set
            Dim dtt As New DataTable("MyTable")
            sdaAccount.Fill(dtt)

            'return value
            Return dtt
        End Function

        Public Function OdbcSPGetDataSet(ByVal strSPName As String, ByVal parameter As OdbcParameter(), ByVal pstrConnection As String, Optional ByRef pscnMain As OdbcConnection = Nothing, Optional ByRef pstnMain As OdbcTransaction = Nothing) As DataSet
            ' 'open connection
            Dim scnMain As New OdbcConnection(pstrConnection)
            scnMain.Open()

            'define command           
            Dim scmHeader As New OdbcCommand(strSPName, scnMain)
            scmHeader.CommandType = CommandType.StoredProcedure
            scmHeader.CommandTimeout = 3600
            'define parameters
            If Not parameter Is Nothing Then
                Dim Param As OdbcParameter
                For Each Param In parameter
                    scmHeader.Parameters.Add(Param)
                Next
            End If

            'create data adapter for the query
            Dim sdaMain As New OdbcDataAdapter(scmHeader)

            'create data set and populate it with table
            Dim dstMain As New DataSet()
            sdaMain.Fill(dstMain)

            'return value
            Return dstMain
        End Function

        Public Function OdbcGetDataSet(ByRef pstrConnection As String, ByRef pstrSQL As String) As DataSet
            Dim adcMain As New OdbcConnection(pstrConnection)
            adcMain.Open()

            'create data adapter
            Dim sdaAccount As New OdbcDataAdapter(pstrSQL, adcMain)

            'load data set
            Dim dstMain As New System.Data.DataSet
            sdaAccount.Fill(dstMain)

            'close connection
            subClean(adcMain)

            'return value
            Return dstMain
        End Function

        Public Sub subClean(ByRef pscnConnect As OdbcConnection, Optional ByRef pcmCommand As OdbcCommand = Nothing)
            'Clean Memory From scmCommand and scnConnection
            If Not pcmCommand Is Nothing Then
                pcmCommand.Dispose()
                pcmCommand = Nothing
            End If

            If Not pscnConnect Is Nothing Then
                If Not pscnConnect.State <> ConnectionState.Closed Then pscnConnect.Close()
                pscnConnect.Dispose()
                pscnConnect = Nothing
            End If
        End Sub

        Private Sub subClean(adcMain As SqlConnection)
            Throw New NotImplementedException
        End Sub

    End Class
End Namespace




