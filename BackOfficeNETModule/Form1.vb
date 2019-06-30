Imports System.IO

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim strArgs() As String
        Dim i As Integer
        Dim strStartNo As String = vbNullString
        Dim lngEndNo As Long = 0
        Dim strClCode As String = vbNullString
        Dim strPrintType As String = vbNullString
        Dim strClientNo As String = vbNullString
        Dim strInvType As String = vbNullString
        Dim blSendEmail As Boolean = False
        Dim strEmailArgs As String = vbNullString
        Dim strEmailType As String = vbNullString
        Dim blExportFile As Boolean = False

        'for background check
        Dim lngEmpID As Long

        strArgs = Split(Command$, " ")


        For i = LBound(strArgs) To UBound(strArgs)

            '        MsgBox "Arg 1: " & Left(LCase(a_strArgs(i)), 2)

            Select Case Mid(LCase(strArgs(i)), 1, 2)

                Case "-b", "/b"
                    lngEmpID = CLng(Mid(strArgs(i), 3, 5))
                    strPrintType = "B"

                Case "-s", "/s"
                    strStartNo = CLng(Mid(strArgs(i), 3, 5))

                Case "-e", "/e"
                    lngEndNo = CLng(Mid(strArgs(i), 3, 5))

                Case "-t", "/t" 'invoices or statements (S: Statements O: Other Invoices)
                    strPrintType = Mid(strArgs(i), 3, 1)

                Case "-i", "/i" 'Temp or Perm (T/P)
                    strInvType = Mid(strArgs(i), 3, 1)

                Case "-c", "/c"
                    strClCode = Mid(strArgs(i), 3)

                Case "-m", "/m" 'email flag
                    blSendEmail = True

                Case "-f", "/f"
                    strEmailArgs = Mid(strArgs(i), 6)

                Case "-x", "/x"
                    strEmailType = Mid(strArgs(i), 3)

                Case "-w", "/w"
                    blExportFile = True


            End Select

        Next i


        Dim objPDF As GGBackOffice.PDFController


        'email
        If strPrintType = "E" Then
            HandleEmail(strEmailType, strStartNo)

        ElseIf strPrintType = "B" Then
            objPDF = New GGBackOffice.PDFController
            objPDF.PrintBackgroundCheck(lngEmpID)
            objPDF = Nothing
        Else


            objPDF = New GGBackOffice.PDFController

            If strPrintType = "S" Then

                strArgs = Split(strClCode, "&")

                If strInvType = "T" Then
                    'print temp statements
                    objPDF.PrintTempStatements(strArgs(0), strArgs(1))
                Else
                    objPDF.PrintPermStatements(strArgs(0), strArgs(1))
                End If

            ElseIf strPrintType = "O" Then
                objPDF.PrintOtherInvoices(CLng(strStartNo), lngEndNo, blSendEmail, blExportFile)

            Else

                If strStartNo > 0 And lngEndNo > 0 Then
                    If strInvType = "T" Then
                        objPDF.PrintInvoices(CLng(strStartNo), lngEndNo, blSendEmail, blExportFile)
                    Else
                        objPDF.PrintPermInvoices(CLng(strStartNo), lngEndNo, blSendEmail)
                    End If
                End If

            End If



            objPDF = Nothing

        End If

        Me.Dispose()


    End Sub

    Private Sub HandleEmail(ByVal strEmailType As String, ByVal lngStartNo As Long)
        Dim objE As GGBackOffice.BAKEmailManager = New GGBackOffice.BAKEmailManager
        Dim strArgs() As String
        Dim strEmail As String

        strArgs = Split(strEmailType, "&")
        strEmail = strArgs(1)

        Select Case UCase(Trim(strArgs(0)))

            Case "ONLINEINVOICEREADY"
                objE.SendOnlineInvoiceNotification(lngStartNo, strEmail)

            Case "UNPAIDINVOICEREMINDER"
                objE.SendUnpaidInvoiceReminder(strEmail, lngStartNo)

            Case "TIMESLIPREADY"

            Case "TIMESLIPREMINDER"
                objE.SendUnsignedTimeSlipReminder(strEmail)

        End Select

    End Sub

    
End Class
