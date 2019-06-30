Imports System.Data.SqlClient
Imports System.Text


Namespace GGBackOffice



    Public Class GGDatabaseController

        Private m_strPayrollConnection As String = vbNullString
        Private m_strGGSSConnection As String = vbNullString
        Private m_strWebConnection As String = vbNullString


        Sub New()
            If Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug" Or Application.StartupPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release" Then
                m_strPayrollConnection = "Data Source=YVETTE\SQLEXPRESS;Initial Catalog=Payroll;Integrated Security=True"
                m_strGGSSConnection = "Data Source=YVETTE\SQLEXPRESS;Initial Catalog=Gateway;Integrated Security=True"
                m_strWebConnection = "Data Source=YVETTE\SQLEXPRESS;Initial Catalog=DotNetNuke4;Integrated Security=True"
            Else
                m_strPayrollConnection = "Data Source=" & My.Settings.DBServerName & ";Initial Catalog=Payroll;Integrated Security=True"
                m_strGGSSConnection = "Data Source=" & My.Settings.DBServerName & ";Initial Catalog=Gateway;Integrated Security=True"
                m_strWebConnection = "Data Source=" & My.Settings.DBServerName & ";Initial Catalog=DotNetNuke4;Integrated Security=True"
            End If

        End Sub



        Public Function GetTempAdjustmentInvoices(ByVal lngStartNo As Long, ByVal lngEndNo As Long) As SqlClient.SqlDataReader
            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            strSQL = "SELECT client.invdept, client.invadd1, client.invadd2, client.invcity,  client.invstate, " &
                    "client.invzip, client.charge_int, client.invcontact, 0 as thispo, adjust.* FROM adjust " &
                    "INNER JOIN client ON client.clcode=adjust.clcode WHERE invnumber BETWEEN " & lngStartNo & " AND " & lngEndNo &
                    "AND adjust.site_id=101 " &
                    "ORDER BY invnumber, lname "

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Debug.Print(strSQL)

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                'conn.Close()
                Return reader

            Catch

                conn.Close()
                Return reader
            End Try


        End Function

        Public Function GetTempInvoices(ByVal lngStartNo As Long, ByVal lngEndNo As Long) As SqlClient.SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            'strSQL = "SELECT gateway..clients.OnlineTimeSlips, gateway..clients.HideInvoicePage2,client.invdept, client.invadd1, client.invadd2, client.invcity,  client.invstate, client.invzip, " &
            '        "client.billemail, client.charge_int, client.invcontact, client.latedays as [CLATEDAYS], joborder.* FROM joborder " &
            '        "INNER JOIN client ON client.clcode=joborder.clcode  " &
            '         " INNER JOIN gateway..clients On gateway..clients.clientno=joborder.clcode " &
            '        " WHERE invnumber BETWEEN " & lngStartNo & " And " & lngEndNo &
            '        "And joborder.site_id=101 " &
            '        "ORDER BY invnumber, assignedto, lname "


            strSQL = "SELECT gateway..clients.OnlineTimeSlips, gateway..clients.HideInvoicePage2,client.clname, client.invdept, client.invadd1, client.invadd2, client.invcity,  client.invstate, client.invzip, " &
                    "client.billemail, client.charge_int, client.invcontact, client.latedays as [CLATEDAYS], 'J' AS [tablecode], JOBORDER.WEDATE, " &
                    "JOBORDER.ACCTEXEC, JOBORDER.BILLADDRESSID, JOBORDER.CLCODE, JOBORDER.INVNUMBER, JOBORDER.INVDATE, JOBORDER.PAYOT, JOBORDER.PAYDT, JOBORDER.PAYST," &
                    "JOBORDER.fname, JOBORDER.lname, JOBORDER.billst, JOBORDER.stbillrate, JOBORDER.billot, JOBORDER.otbillrate, JOBORDER.billdt, JOBORDER.dtbillrate, JOBORDER.miscbill, JOBORDER.originvamt, " &
                    "JOBORDER.jobnumber, JOBORDER.thisPO, JOBORDER.typeassign, JOBORDER.reason, JOBORDER.latedays, JOBORDER.assignedto FROM joborder " &
                    "INNER JOIN client ON client.clcode=joborder.clcode  " &
                     " INNER JOIN gateway..clients On gateway..clients.clientno=joborder.clcode " &
                      " WHERE invnumber BETWEEN " & lngStartNo & " And " & lngEndNo &
                    " And joborder.site_id=101 " &
                    "ORDER BY invnumber, assignedto, lname "



            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                'conn.Close()
                Return reader

            Catch

                conn.Close()
                Return reader
            End Try




        End Function

        Public Function GetGABillingAddress(ByVal lngAddressID As Long) As SqlClient.SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strGGSSConnection)
            Dim reader As SqlDataReader

            conn.Open()


            strSQL = "SELECT * FROM clientbill WHERE pk_id=" & lngAddressID

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                Return reader

            Catch

                Return reader
            End Try

        End Function


        Public Function GetCompanyInformation() As SqlClient.SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString
            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()


            strSQL = "SELECT * FROM mastvars WHERE site_id=101"
            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                Return reader

            Catch

                Return reader
            End Try

        End Function



        Public Function GetInvoicePageCount(ByVal lngInvNumber As Long) As Integer
            Dim dcJobCount As Decimal = 0, intPageCount As Integer = 0
            Dim cmd As New SqlCommand
            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            conn.Open()

            cmd.Connection = conn
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "bak_GetInvoiceJobCount"

            Dim param1 As New SqlParameter
            param1.Direction = ParameterDirection.Input
            param1.ParameterName = "@InvoiceNumber"
            param1.Value = lngInvNumber
            param1.SqlDbType = SqlDbType.Int

            Dim param2 As New SqlParameter
            With param2
                .Direction = ParameterDirection.Output
                .ParameterName = "@Count"
                .SqlDbType = SqlDbType.Int
            End With

            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)


            Try
                cmd.ExecuteNonQuery()
                dcJobCount = cmd.Parameters(1).Value

            Catch ex As Exception

            End Try

            If dcJobCount > 2 Then
                dcJobCount = dcJobCount / 2

                intPageCount = Math.Round(dcJobCount, 0, MidpointRounding.AwayFromZero)
            Else
                intPageCount = 1
            End If



            conn.Close()
            Return intPageCount

        End Function

        Public Function GetInvoiceJobCount(ByVal lngInvNumber As Long) As Integer
            Dim intJobCount As Integer = 0
            Dim cmd As New SqlCommand
            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            conn.Open()

            cmd.Connection = conn
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "bak_GetInvoiceJobCount"

            Dim param1 As New SqlParameter
            param1.Direction = ParameterDirection.Input
            param1.ParameterName = "@InvoiceNumber"
            param1.Value = lngInvNumber
            param1.SqlDbType = SqlDbType.Int

            Dim param2 As New SqlParameter
            With param2
                .Direction = ParameterDirection.Output
                .ParameterName = "@Count"
                .SqlDbType = SqlDbType.Int
            End With

            cmd.Parameters.Add(param1)
            cmd.Parameters.Add(param2)


            Try
                cmd.ExecuteNonQuery()
                intJobCount = cmd.Parameters(1).Value

            Catch ex As Exception

            End Try
            conn.Close()
            Return intJobCount

        End Function



        Public Function GetTimeSlipCount(ByVal lngInvnumber As Long)
            Dim intTSCount As Integer = 0
            'Dim cmd As New SqlCommand
            'Dim conn As SqlConnection = New SqlConnection(Me.m_strWebConnection)
            'conn.Open()

            'cmd.Connection = conn
            'cmd.CommandType = CommandType.StoredProcedure
            'cmd.CommandText = "GGGetTimeSlipCount"

            'Dim param1 As New SqlParameter
            'param1.Direction = ParameterDirection.Input
            'param1.ParameterName = "@InvoiceNumber"
            'param1.Value = lngInvnumber
            'param1.SqlDbType = SqlDbType.Int

            'Dim param2 As New SqlParameter
            'With param2
            '    .Direction = ParameterDirection.Output
            '    .ParameterName = "@rCount"
            '    .SqlDbType = SqlDbType.Int
            'End With

            'cmd.Parameters.Add(param1)
            'cmd.Parameters.Add(param2)


            'Try
            '    cmd.ExecuteNonQuery()
            '    intTSCount = cmd.Parameters(1).Value

            'Catch ex As Exception

            'End Try

            'conn.Close()
            Return intTSCount

        End Function

        Public Function GetJobRecord(ByVal lngJobnumber As Long) As SqlDataReader
            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            strSQL = "SELECT joborder.* FROM joborder " &
                    "WHERE jobnumber=" & lngJobnumber &
                    " And joborder.site_id=101 " &
                    "ORDER BY invnumber, lname "


            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                'conn.Close()
                Return reader

            Catch

                conn.Close()
                Return reader
            End Try
        End Function

        Public Function GetPermInvoices(ByVal lngStartNo As Long, ByVal lngEndNo As Long) As SqlClient.SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            strSQL = "SELECT client.billemail, client.invdept, client.invadd1, client.invadd2, client.invcity,  client.invstate, " &
                "client.invzip, ays_perm.attn AS [invcontact], client.permpo, ays_perm.chargeint AS [charge_int], ays_perm.*  FROM ays_perm " &
                "INNER JOIN client ON client.clcode=ays_perm.clcode " &
                " WHERE invnumber BETWEEN " & lngStartNo & " And " & lngEndNo &
                " ORDER BY ays_perm.clname, ays_perm.CLCODE "


            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                'conn.Close()
                Return reader

            Catch

                conn.Close()
                Return reader
            End Try




        End Function

        Public Function GetPermAdjustmentInvoices(ByVal lngStartNo As Long, ByVal lngEndNo As Long) As SqlClient.SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            strSQL = "SELECT client.billemail, client.invdept, client.invadd1, client.invadd2, client.invcity,  client.invstate, " &
                "client.invzip, client.charge_int, permadjust.attn AS [invcontact], client.permpo, permadjust.*  FROM permadjust " &
                "INNER JOIN client ON client.clcode=permadjust.clcode " &
                " WHERE adjinvnumber BETWEEN " & lngStartNo & " And " & lngEndNo &
                " ORDER BY permadjust.clname, permadjust.CLCODE "


            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                'conn.Close()
                Return reader

            Catch

                conn.Close()
                Return reader
            End Try




        End Function


        Public Function GetTempStatementClients(ByVal strClcode As String, ByVal blPrintCurrent As Boolean)

            Dim d1630 As Date
            d1630 = DateAdd("d", -16, Now)

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            If strClcode = "-99" Then

                If blPrintCurrent Then

                    strSQL = "SELECT 0 as billaddressid, charge_int, latedays, clcode, clname, invcontact, invadd1, invadd2, invcity, invstate, invzip, invdept, statementmsg1, statementmsg2, statementmsg3  FROM client WHERE " &
                        " clcode IN (SELECT clcode FROM joborder WHERE originvamt + miscbill  > 0 And paidinfull=0 And " &
                        "(invnumber > 0 And invnumber Is Not NULL)) ORDER BY ltrim(CLNAME), clcode "
                Else

                    strSQL = "SELECT 0 as billaddressid, charge_int, latedays, clcode, clname, invcontact, invadd1, invadd2, invcity, invstate, invzip, invdept, statementmsg1, statementmsg2, statementmsg3  FROM client WHERE " &
                        "  clcode IN (SELECT clcode FROM joborder WHERE originvamt + miscbill > 0 And paidinfull=0 And " &
                        "(invnumber > 0 And invnumber Is Not NULL) And invdate < '" & d1630 & "') ORDER BY ltrim(CLNAME), clcode "

                End If


            Else

                strSQL = "SELECT 0 as billaddressid, charge_int, latedays, clcode, clname, invcontact, invadd1, invadd2, invcity, invstate, invzip, invdept," & _
                    "statementmsg1, statementmsg2, statementmsg3 FROM client WHERE clcode='" & strClcode & "'"

            End If

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                'conn.Close()
                Return reader

            Catch

                conn.Close()
                Return reader
            End Try

        End Function

        Public Function GetTempStatementInvoices(ByVal strClCode As String, ByVal blPrintCurrent As Boolean, ByVal dDate As Date) As SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            If blPrintCurrent Then

                strSQL = "SELECT joborder.latedays, joborder.invnumber, invdate, sum(originvamt + miscbill) AS [InvTotal], null AS [AdjAmt] FROM joborder " & _
                            "WHERE joborder.paidinfull=0 AND (joborder.originvamt + miscbill  > 0) " & _
                             " AND (joborder.invnumber IS NOT NULL AND joborder.invnumber > 0) AND " & _
                             " joborder.clcode='" & strClCode & "' GROUP BY invnumber, invdate, joborder.latedays "

                strSQL = strSQL & "UNION SELECT adjust.latedays, invnumber, invdate, null AS [InvTotal], sum(invadjust) AS [AdjAmt] FROM adjust " & _
                                "WHERE paidinfull = 0 And (invadjust > 0) AND invnumber <> '-999' " & _
                                "AND (adjust.invnumber IS NOT NULL AND adjust.invnumber > 0) " & _
                                " AND adjust.clcode='" & strClCode & "'  GROUP BY invnumber, invdate, adjust.latedays  " & _
                                "ORDER BY invnumber"
            Else

                strSQL = "SELECT joborder.latedays, invnumber, invdate, sum(originvamt + miscbill) AS [InvTotal], null AS [AdjAmt] FROM joborder " & _
                                "WHERE joborder.paidinfull=0 AND (joborder.originvamt + miscbill  > 0) " & _
                             " AND (joborder.invnumber IS NOT NULL AND joborder.invnumber > 0) AND " & _
                             " joborder.clcode='" & strClCode & "' AND invdate < '" & dDate & "' GROUP BY invnumber, invdate, joborder.latedays  "

                strSQL = strSQL & "UNION SELECT adjust.latedays, invnumber, invdate, null AS [InvTotal], sum(invadjust) AS [AdjAmt] FROM adjust " & _
                                "WHERE paidinfull = 0 And (invadjust > 0) AND invnumber <> '-999' " & _
                                "AND (adjust.invnumber IS NOT NULL AND adjust.invnumber > 0) " & _
                                " AND adjust.clcode='" & strClCode & "' AND invdate < '" & dDate & "' GROUP BY invnumber, invdate, adjust.latedays " & _
                                "ORDER BY invnumber"
                Debug.Print(strSQL)

            End If

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                Return reader

            Catch

                conn.Close()

            End Try

        End Function

        Public Function GetTempStatementOtherInvoices(ByVal strClCode As String, ByVal blPrintCurrent As Boolean, ByVal dDate As Date) As SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            If blPrintCurrent Then
                strSQL = "SELECT invnumber, invdate, sum(amount) AS [InvTotal], null AS [AdjAmt] FROM MISCBILL WHERE clcode='" & Trim(strClCode) & "' GROUP BY invnumber, invdate" & _
                                    " ORDER BY invnumber"
            Else

                strSQL = "SELECT invnumber, invdate, sum(amount) AS [InvTotal], null AS [AdjAmt] FROM MISCBILL WHERE clcode='" & Trim(strClCode) & "' AND invdate < '" & dDate & "' GROUP BY invnumber, invdate" & _
                                   " ORDER BY invnumber"
            End If

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                Return reader

            Catch

                conn.Close()

            End Try

        End Function

        Public Function GetInvoicePayments(ByVal strInvNumber As String) As Decimal

            Dim dcPayments As Decimal = 0

            'create connection
            Dim con As SqlConnection
            con = New SqlConnection(m_strPayrollConnection)
            con.Open()

            'create command object
            Dim cmd As New SqlCommand()
            cmd.CommandText = "bak_GetSumInvoicePayments"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = con

            'create parameters for command object
            Dim parameter1 As New SqlParameter()
            parameter1.ParameterName = "@InvoiceNumber"
            parameter1.SqlDbType = SqlDbType.VarChar
            parameter1.Direction = ParameterDirection.Input
            parameter1.Value = strInvNumber

            Dim parameter2 As New SqlParameter()
            parameter2.ParameterName = "@rPayments"
            parameter2.SqlDbType = SqlDbType.Money
            parameter2.Direction = ParameterDirection.Output
            parameter2.Size = 100

            cmd.Parameters.Add(parameter1)
            cmd.Parameters.Add(parameter2)

            Try
                cmd.ExecuteNonQuery()

                If Not cmd.Parameters("@rPayments").Value Is System.DBNull.Value Then dcPayments = CDec(cmd.Parameters("@rPayments").Value)


            Catch ex As Exception

                dcPayments = -99

            Finally
                If con.State = ConnectionState.Open Then con.Close()

            End Try

            Return dcPayments


        End Function

        Public Function GetInvoiceCreditMemos(ByVal strInvNumber As String) As Decimal

            Dim dcPayments As Decimal = 0

            'create connection
            Dim con As SqlConnection
            con = New SqlConnection(m_strPayrollConnection)
            con.Open()

            'create command object
            Dim cmd As New SqlCommand()
            cmd.CommandText = "bak_GetSumInvoiceCreditMemos"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = con

            'create parameters for command object
            Dim parameter1 As New SqlParameter()
            parameter1.ParameterName = "@InvoiceNumber"
            parameter1.SqlDbType = SqlDbType.VarChar
            parameter1.Direction = ParameterDirection.Input
            parameter1.Value = strInvNumber

            Dim parameter2 As New SqlParameter()
            parameter2.ParameterName = "@rPayments"
            parameter2.SqlDbType = SqlDbType.Money
            parameter2.Direction = ParameterDirection.Output
            parameter2.Size = 100

            cmd.Parameters.Add(parameter1)
            cmd.Parameters.Add(parameter2)

            Try
                cmd.ExecuteNonQuery()

                If Not cmd.Parameters("@rPayments").Value Is System.DBNull.Value Then dcPayments = CDec(cmd.Parameters("@rPayments").Value)


            Catch ex As Exception

                dcPayments = -99

            Finally
                If con.State = ConnectionState.Open Then con.Close()

            End Try

            Return dcPayments


        End Function

        Public Function GetGlobalStatementMessages() As String

            Dim strMessages As String = vbNullString

            'create connection
            Dim con As SqlConnection
            con = New SqlConnection(m_strPayrollConnection)
            con.Open()

            'create command object
            Dim cmd As New SqlCommand()
            cmd.CommandText = "bak_GetGlobalStatementMessages"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = con

            Dim parameter1 As New SqlParameter()
            parameter1.ParameterName = "@rMessages"
            parameter1.SqlDbType = SqlDbType.VarChar
            parameter1.Direction = ParameterDirection.Output
            parameter1.Size = 600

            cmd.Parameters.Add(parameter1)


            Try
                cmd.ExecuteNonQuery()

                If Not cmd.Parameters("@rMessages").Value Is System.DBNull.Value Then strMessages = cmd.Parameters("@rMessages").Value


            Catch ex As Exception

                strMessages = vbNullString

            Finally
                If con.State = ConnectionState.Open Then con.Close()

            End Try

            Return strMessages

        End Function

        Public Function GetInterestPercent() As Integer

            Dim intPercent As Integer = 0

            'create connection
            Dim con As SqlConnection
            con = New SqlConnection(m_strPayrollConnection)
            con.Open()

            'create command object
            Dim cmd As New SqlCommand()
            cmd.CommandText = "bak_GetInterestRate"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = con

            Dim parameter1 As New SqlParameter()
            parameter1.ParameterName = "@rInterest"
            parameter1.SqlDbType = SqlDbType.Int
            parameter1.Direction = ParameterDirection.Output

            cmd.Parameters.Add(parameter1)


            Try
                cmd.ExecuteNonQuery()

                If Not cmd.Parameters("@rInterest").Value Is System.DBNull.Value Then intPercent = cmd.Parameters("@rInterest").Value


            Catch ex As Exception

                intPercent = 0

            Finally
                If con.State = ConnectionState.Open Then con.Close()

            End Try

            Return intPercent

        End Function

        Public Function GetInterestStartDate() As Date

            Dim dStartDate As Date

            'create connection
            Dim con As SqlConnection
            con = New SqlConnection(m_strPayrollConnection)
            con.Open()

            'create command object
            Dim cmd As New SqlCommand()
            cmd.CommandText = "bak_GetInterestStartDate"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = con

            Dim parameter1 As New SqlParameter()
            parameter1.ParameterName = "@rDate"
            parameter1.SqlDbType = SqlDbType.SmallDateTime
            parameter1.Direction = ParameterDirection.Output

            cmd.Parameters.Add(parameter1)


            Try
                cmd.ExecuteNonQuery()

                If Not cmd.Parameters("@rDate").Value Is System.DBNull.Value Then dstartdate = cmd.Parameters("@rDate").Value


            Catch ex As Exception

                dStartDate = "1/1/1900"

            Finally
                If con.State = ConnectionState.Open Then con.Close()

            End Try

            Return dStartDate

        End Function

        Public Function GetInterestExpirationDays() As Integer

            Dim intPercent As Integer = 0

            'create connection
            Dim con As SqlConnection
            con = New SqlConnection(m_strPayrollConnection)
            con.Open()

            'create command object
            Dim cmd As New SqlCommand()
            cmd.CommandText = "bak_GetInterestExpirationDays"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = con

            Dim parameter1 As New SqlParameter()
            parameter1.ParameterName = "@rDays"
            parameter1.SqlDbType = SqlDbType.Int
            parameter1.Direction = ParameterDirection.Output

            cmd.Parameters.Add(parameter1)


            Try
                cmd.ExecuteNonQuery()

                If Not cmd.Parameters("@rDays").Value Is System.DBNull.Value Then intPercent = cmd.Parameters("@rDays").Value


            Catch ex As Exception

                intPercent = 0

            Finally
                If con.State = ConnectionState.Open Then con.Close()

            End Try

            Return intPercent

        End Function

        Public Function GetStatementUnpaidInterestItems(ByVal strClCode As String, ByVal dCutOff As Date, Optional ByVal blPerm As Boolean = False) As SqlClient.SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            If Not blPerm Then

                strSQL = "SELECT * FROM UnpaidInterest WHERE Interest - (Paid + Discard) > 0 AND clcode='" & strClCode & "' AND dateadded > '" & dCutOff & "' AND (INVTYPE='J' OR INVTYPE='JA')"
            Else
                strSQL = "SELECT * FROM UnpaidInterest WHERE Interest - (Paid + Discard) > 0 AND clcode='" & strClCode & "' AND dateadded > '" & dCutOff & "' AND (INVTYPE='P' OR INVTYPE='PA')"
            End If


            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                Return reader

            Catch

                conn.Close()
                Return reader
            End Try




        End Function

        Public Function GetSumUnpaidTempInvoices(ByVal strClCode As String, ByVal blPrintCurrent As Boolean) As Decimal

            Dim cmd As New SqlCommand
            Dim strSQL As String
            Dim dcTotal As Decimal
            Dim d1630 As Date
            d1630 = DateAdd("d", -16, Now)

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            If blPrintCurrent Then
                strSQL = "SELECT sum(joborder.originvamt) AS [TotalTemp] FROM joborder " & _
                         "WHERE joborder.paidinfull=0 AND (joborder.originvamt + miscbill  > 0) " & _
                         " AND (joborder.invnumber > 0 AND joborder.invnumber IS NOT NULL) " & _
                         " AND joborder.clcode='" & strClCode & "' "

            Else
                strSQL = "SELECT sum(joborder.originvamt) AS [TotalTemp] FROM joborder " & _
                         "WHERE joborder.paidinfull=0 AND (joborder.originvamt + miscbill  > 0) " & _
                         " AND (joborder.invnumber > 0 AND joborder.invnumber IS NOT NULL) " & _
                         " AND joborder.clcode='" & strClCode & "' AND invdate < '" & d1630 & "'"

            End If

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader
                reader.Read()
                If Not reader("TotalTemp") Is System.DBNull.Value Then dcTotal = reader("TotalTemp")

            Catch



            End Try

            reader.Close()

            'check for unpaid adjustment invoices -----------------------------------------------------
            If blPrintCurrent Then
                strSQL = "SELECT sum(adjust.invadjust) AS [TotalTemp] FROM adjust " & _
                         "WHERE adjust.invadjust + miscbill  > 0 " & _
                         " AND (adjust.invnumber > 0 AND adjust.invnumber IS NOT NULL) " & _
                         " AND adjust.clcode='" & strClCode & "' "

            Else
                strSQL = "SELECT sum(adjust.invadjust) AS [TotalTemp] FROM adjust " & _
                         "WHERE adjust.invadjust + miscbill  > 0 " & _
                         " AND (adjust.invnumber > 0 AND adjust.invnumber IS NOT NULL) " & _
                         " AND adjust.clcode='" & strClCode & "' AND invdate < '" & d1630 & "'"

            End If

            cmd.CommandText = strSQL

            Try
                reader = cmd.ExecuteReader

                reader.Read()
                If Not reader("TotalTemp") Is System.DBNull.Value Then dcTotal = dcTotal + reader("TotalTemp")

            Catch



            End Try



            conn.Close()

            Return dcTotal

        End Function

        Public Function GetSumTempInvoicePayments(ByVal strClCode As String) As Decimal

            Dim cmd As New SqlCommand
            Dim strSQL As String
            Dim dcTotal As Decimal
            Dim d1630 As Date
            d1630 = DateAdd("d", -16, Now)

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            strSQL = "SELECT sum(amountpaid + ignore) AS [TotalRec] FROM receipts " & _
                     "WHERE (invnumber IN " & _
                     "(SELECT DISTINCT invnumber FROM joborder " & _
                     "WHERE joborder.paidinfull=0 AND (joborder.originvamt + miscbill > 0)" & _
                     "OR invnumber IN " & _
                     "(SELECT DISTINCT invnumber FROM adjust " & _
                     "WHERE invadjust + miscbill > 0))" & _
                     " AND joborder.clcode='" & strClCode & "' "

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                reader.Read()
                If Not reader("TotalRec") Is System.DBNull.Value Then dcTotal = reader("TotalRec")

            Catch



            End Try

            conn.Close()

            Return dcTotal

        End Function

        Public Function GetPermStatementClients(ByVal strClcode As String, ByVal blPrintCurrent As Boolean)

            Dim d1630 As Date
            d1630 = DateAdd("d", -16, Now)

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            If strClcode = "-99" Then

                If blPrintCurrent Then

                    strSQL = "SELECT charge_int, latedays, clcode, clname, invcontact, invadd1, invadd2, invcity, invstate, invzip, invdept, statementmsg1, statementmsg2, statementmsg3  FROM client WHERE " & _
                        " clcode IN (SELECT clcode FROM ays_perm WHERE (invnumber > 0 AND invnumber IS NOT NULL)) " & _
                        " ORDER BY ltrim(CLNAME), clcode "
                Else

                    strSQL = "SELECT charge_int, latedays, clcode, clname, invcontact, invadd1, invadd2, invcity, invstate, invzip, invdept, statementmsg1, statementmsg2, statementmsg3  FROM client WHERE " & _
                        " clcode IN (SELECT clcode FROM ays_perm WHERE (invnumber > 0 AND invnumber IS NOT NULL) AND invdate < '" & d1630 & "') " & _
                        " ORDER BY ltrim(CLNAME), clcode "

                End If


            Else

                'check for invoices
                'If Not InvoicesExist(strClcode, , , True) Then
                '    MsgBox("This client has no open invoices.", vbOKOnly + vbInformation)
                '    Exit Function
                'End If

                strSQL = "SELECT charge_int, latedays, clcode, clname, invcontact, invadd1, invadd2, invcity, invstate, invzip, invdept," & _
                    "statementmsg1, statementmsg2, statementmsg3 FROM client WHERE clcode='" & strClcode & "'"

            End If

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                'conn.Close()
                Return reader

            Catch

                conn.Close()
                Return reader
            End Try

        End Function

        Public Function GetPermStatementInvoices(ByVal strClCode As String, ByVal blPrintCurrent As Boolean, ByVal dDate As Date) As SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            If blPrintCurrent Then

                strSQL = "SELECT emplname, comment1, terms, startdate, latedays, chargeint, invnumber, invdate, sum(AMOUNT) AS [InvTotal] FROM ays_perm " & _
                             "WHERE clcode='" & strClCode & "' AND (amount IS NOT NULL AND amount > 0) " & _
                             " AND (invnumber IS NOT NULL AND invnumber > 0) " & _
                             " GROUP BY invnumber, invdate, latedays, chargeint, startdate, terms, emplname, comment1 "


                '- INCLUDE POSITIVE INVOICES FROM PERMADJUST -----------------
                strSQL = strSQL & "UNION SELECT '' AS [emplname], '' AS [comment1], '' AS [terms], 0 AS [startdate], 0 as [latedays], 'N' AS [chargeint], AdjInvnumber AS [invnumber], adjinvdate AS [invdate], " & _
                                "sum(grossadj) AS [invtotal] FROM permadjust " & _
                                "WHERE (grossadj > 0) AND " & _
                                "(adjinvnumber IS NOT NULL AND adjinvnumber > 0) " & _
                                " AND clcode='" & strClCode & "'  GROUP BY AdjInvnumber, adjinvdate " & _
                                "ORDER BY Invnumber"

            Else
                strSQL = "SELECT emplname, comment1, terms, startdate, latedays, chargeint, invnumber, invdate, sum(AMOUNT) AS [InvTotal] FROM ays_perm " & _
                                 "WHERE clcode='" & strClCode & "' AND (amount IS NOT NULL AND amount > 0) " & _
                                 " AND (invnumber IS NOT NULL AND invnumber > 0) AND invdate < '" & dDate & "' " & _
                                 " GROUP BY invnumber, invdate, latedays, chargeint, startdate, terms, emplname, comment1 "


                '- INCLUDE POSITIVE INVOICES FROM PERMADJUST -----------------
                strSQL = strSQL & "UNION SELECT '' AS [emplname], '' AS [comment1], '' AS [terms], 0 AS [startdate], 0 as [latedays], 'N' AS [chargeint], AdjInvnumber AS [invnumber], adjinvdate AS [invdate], " & _
                                "sum(grossadj) AS [invtotal] FROM permadjust " & _
                                "WHERE (grossadj > 0) AND " & _
                                "(adjinvnumber IS NOT NULL AND adjinvnumber > 0) AND adjinvdate < '" & dDate & "' " & _
                                " AND clcode='" & strClCode & "' GROUP BY AdjInvnumber, adjinvdate " & _
                                "ORDER BY Invnumber"


            End If

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                'conn.Close()
                Return reader

            Catch

                conn.Close()

            End Try

        End Function

        Public Function GetSumUnpaidPermInvoices(ByVal strClCode As String, ByVal blPrintCurrent As Boolean) As Decimal

            Dim cmd As New SqlCommand
            Dim strSQL As String
            Dim dcTotal As Decimal
            Dim d1630 As Date
            d1630 = DateAdd("d", -16, Now)

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            If blPrintCurrent Then
                strSQL = "SELECT sum(ays_perm.amount) AS [TotalPerm] FROM ays_perm " & _
                        "WHERE (amount > 0 AND amount IS NOT NULL) AND " & _
                        "(invnumber > 0 AND invnumber IS NOT NULL) " & _
                        "AND clcode='" & strClCode & "' "
            Else
                strSQL = "SELECT sum(ays_perm.amount) AS [TotalPerm] FROM ays_perm " & _
                        "WHERE (amount > 0 AND amount IS NOT NULL) AND " & _
                        "(invnumber > 0 AND invnumber IS NOT NULL) " & _
                        "AND clcode='" & strClCode & "' AND invdate < '" & d1630 & "'"

            End If

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader

                reader.Read()
                If Not reader("TotalPerm") Is System.DBNull.Value Then dcTotal = reader("TotalPerm")

                reader.Close()

            Catch



            End Try



            'get any positive perm adjustments
            If blPrintCurrent Then
                strSQL = "SELECT sum(grossadj) AS [TotalPerm] FROM permadjust " & _
                        "WHERE (grossadj > 0 AND grossadj IS NOT NULL) AND " & _
                        "(adjinvnumber > 0 AND adjinvnumber IS NOT NULL) " & _
                        "AND clcode='" & strClCode & "' "
            Else
                strSQL = "SELECT sum(grossadj) AS [TotalPerm] FROM permadjust " & _
                        "WHERE (grossadj > 0 AND grossadj IS NOT NULL) AND " & _
                        "(adjinvnumber > 0 AND adjinvnumber IS NOT NULL) " & _
                        "AND clcode='" & strClCode & "' AND adjinvdate < '" & d1630 & "'"

            End If

            cmd.CommandText = strSQL

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                reader.Read()
                If Not reader("TotalPerm") Is System.DBNull.Value Then dcTotal = dcTotal + reader("TotalPerm")

            Catch



            End Try

            conn.Close()

            Return dcTotal

        End Function

        Public Function GetSumPermInvoicePayments(ByVal strClCode As String) As Decimal

            Dim cmd As New SqlCommand
            Dim strSQL As String
            Dim dcTotal As Decimal
            Dim d1630 As Date
            d1630 = DateAdd("d", -16, Now)

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            strSQL = "SELECT sum(amountpaid + ignore) AS [TotalRec] FROM receipts " & _
                     "WHERE (invnumber IN " & _
                     "(SELECT DISTINCT invnumber FROM ays_perm " & _
                     "WHERE amount > 0)" & _
                     "OR invnumber IN " & _
                     "(SELECT DISTINCT invnumber FROM ays_perm " & _
                     "WHERE amount > 0))" & _
                     " AND receipts.clcode='" & strClCode & "' "

            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                reader.Read()
                If Not reader("TotalRec") Is System.DBNull.Value Then dcTotal = reader("TotalRec")

            Catch



            End Try

            conn.Close()

            Return dcTotal

        End Function

        Public Function GetInvoiceJobIDs(ByVal lngInvNumber As Long) As SqlDataReader

            Dim cmd As New SqlCommand
            Dim sbSQL As StringBuilder = New StringBuilder


            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            sbSQL.Append("SELECT pk_id FROM JOBORDER WHERE invnumber=")
            sbSQL.Append(CStr(lngInvNumber))
            sbSQL.Append(" ORDER BY lname DESC")

            cmd.CommandText = sbSQL.ToString
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                Return reader

            Catch

                conn.Close()

            End Try

        End Function

      

        Public Function GetTimeSlipByAssignmentID(ByVal lngAssignmentID As Long) As SqlDataReader

            Dim cmd As New SqlCommand
            Dim param As SqlParameter = New SqlParameter
            Dim conn As SqlConnection = New SqlConnection(m_strWebConnection)
            Dim reader As SqlDataReader

            conn.Open()

            cmd.CommandText = "GGGetTimeSlipByAssignmentID"
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = conn

            param.ParameterName = "@AssignmentID"
            param.Value = lngAssignmentID
            param.Direction = ParameterDirection.Input

            cmd.Parameters.Add(param)


            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                Return reader

            Catch ex As Exception
                conn.Close()

            End Try



        End Function


        Public Function GetTimeSlipHours(ByVal lngTimeSlipID As Long) As SqlDataReader
            Dim cmd As New SqlCommand
            Dim param As SqlParameter = New SqlParameter
            Dim reader As SqlDataReader
            Dim conn As SqlConnection = New SqlConnection(m_strWebConnection)
            conn.Open()

            cmd.Connection = conn
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "GGGetTimeSlipDays"

            param.ParameterName = "@TimeSlipID"
            param.Value = lngTimeSlipID
            param.Direction = ParameterDirection.Input

            cmd.Parameters.Add(param)


            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                Return reader

            Catch ex As Exception
                conn.Close()
                Return reader
            End Try



        End Function

        Public Function GetClientBillingEmail(ByVal lngInvNumber As Long, ByVal strInvType As String)

            If strInvType = "P" Then

                'look up clcode in AYS_PERM Table and use to retrieve bill email from client table

            Else
                'look up bill address ID in joborder table


                'if bill address ID is null

            End If

        End Function

        Public Function GetClientBillingEmail(ByVal strClCode As String)

        End Function


        Public Function GetOtherInvoices(ByVal lngStartNo As Long, ByVal lngEndNo As Long) As SqlClient.SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strPayrollConnection)
            Dim reader As SqlDataReader

            conn.Open()

            strSQL = "SELECT '' AS [permpo], NULL AS [billaddressID], client.invdept, client.invadd1, client.invadd2, client.invcity,  client.invstate, client.invzip, " &
                    "client.billemail, client.charge_int, client.invcontact, client.latedays, miscbill.* FROM miscbill " &
                    "INNER JOIN client ON client.clcode=miscbill.clcode  WHERE invnumber BETWEEN " & lngStartNo & " AND " & lngEndNo &
                    "AND miscbill.site_id=101 " &
                    "ORDER BY invnumber "


            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                'conn.Close()
                Return reader

            Catch

                conn.Close()
                Return reader
            End Try




        End Function



        Public Function GetBackgroundCheckInfo(ByVal lngEmpID As Long) As SqlClient.SqlDataReader

            Dim cmd As New SqlCommand
            Dim strSQL As String = vbNullString

            Dim conn As SqlConnection = New SqlConnection(m_strGGSSConnection)
            Dim reader As SqlDataReader

            conn.Open()



            strSQL = "SELECT * FROM GGOnlineFormsBackgroundCheck WHERE empid=" & lngEmpID



            cmd.CommandText = strSQL
            cmd.CommandType = CommandType.Text
            cmd.Connection = conn

            Try
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection)

                'conn.Close()
                Return reader

            Catch

                conn.Close()
                Return reader
            End Try




        End Function




    End Class

    
  


    

End Namespace
