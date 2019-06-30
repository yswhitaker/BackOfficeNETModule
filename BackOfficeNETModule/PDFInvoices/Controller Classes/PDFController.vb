Imports System
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports ceTe.DynamicPDF
Imports ceTe.DynamicPDF.PageElements
Imports BackOfficeNETClassLibrary
Imports System.IO
Imports System.Math
Imports System.Text
Imports System.Collections


Namespace GGBackOffice


    Public Class PDFController

#Region "Variables"

        Dim m_blPrintingTemp As Boolean = False

        Dim m_strImagePath As String = vbNullString
        Dim m_strMultiJobImagePath As String = vbNullString
        Dim m_strSImagePath As String = vbNullString
        Dim m_strTSImagePath As String = vbNullString
        Dim m_strPDFPath As String = vbNullString
        Dim m_strAppDirectory As String = Application.StartupPath
        Dim m_strTimeSlipTextPath As String = vbNullString
        Dim m_strBCHeaderPath As String = ""

        Private m_intTSCount As Integer = 0

        'columns - Temp Invoices ---------------------------------------
        Dim intATTNCol As Integer = 10
        Dim intCLCodeCol As Integer = 280
        Dim intDateCol As Integer = 350
        Dim intInvCol As Integer = 450
        Dim intPageCol As Integer = 515

        Dim intLowerCol1 As Integer = 0
        Dim intLowerCol2 As Integer = 70
        Dim intLowerCol3 As Integer = 140
        Dim intLowerCol4 As Integer = 240
        Dim intLowerCol5 As Integer = 340
        Dim intLowerCol6 As Integer = 430

        'columns - Temp Adjustment Invoices -----------------------------
        'top portion
        Dim adjAttnCOL As Integer = 10
        Dim adjCustNumCol As Integer = 280   'customer number
        Dim adjDateCol As Integer = 350
        Dim adjInvNumCol As Integer = 450  'invoice number

        Dim adjCol1 As Integer = 1
        Dim adjCol2 As Integer = 45

        'right- aligned
        Dim adjCol3 As Integer = 100
        Dim adjCol4 As Integer = 140
        Dim adjCol5 As Integer = 190
        Dim adjCol6 As Integer = 240
        Dim adjCol7 As Integer = 290
        Dim adjCol8 As Integer = 350
        Dim adjCol9 As Integer = 390
        Dim adjCol10 As Integer = 450

        Dim intCurrentY As Integer = 40
        Dim lblText As Label = New Label("", 0, 0, 100, 100)

        Dim m_objCompany As CompanyInfo

        Private m_dcThisInvoiceTotal As Decimal = 0

        'columns - Temp Statements ------------------------------------------
        Dim intSInvNumCol As Integer = 10
        Dim intSInvDateCol As Integer = 75

        'right-aligned
        Dim intSOrigInvCol As Integer = 100
        Dim intSCurrentCol As Integer = 155
        Dim intSInterestCol As Integer = 195
        Dim intS1630Col As Integer = 250
        Dim intS3160Col As Integer = 325
        Dim intS6190Col As Integer = 400
        Dim intSOver90Col As Integer = 450

        Dim m_intEndOfPage As Integer = 600

        Private m_blOnlineTimeSlips As Boolean
        Private m_blHidePage2 As Boolean = False


#End Region

#Region "Temp Invoices"

        Public Function SavePDF(ByVal MyDocument As ceTe.DynamicPDF.Document, ByVal objI As InvoiceInfo, ByVal strClientName As String, ByVal strInvNumber As String, ByVal blSendEMail As Boolean, Optional ByVal blExportFile As Boolean = False, Optional ByVal strClCode As String = vbNullString)
            Dim strDocumentName As String = vbNullString
            Dim strFile As String = vbNullString
            Dim objE As EmailInfo


            strDocumentName = strClientName & strInvNumber
            strFile = strDocumentName.Replace("\", " ")
            strFile = strFile.Replace("/", "")
            strFile = strFile.Replace(".", "")
            strFile = strFile.Replace("-", "")

            Try
                MyDocument.Draw(m_strPDFPath & strFile & ".pdf")

                If blSendEMail Then
                    Dim objM As BAKEmailManager = New BAKEmailManager
                    Dim strEmailInfo As String = vbNullString

                    'From
                    objE.From = "yvette@yswconsulting.com"

                    'To - need to use objC.BillAddress
                    objE.MailTo = "yvette@yswconsulting.com"

                    'Subject
                    objE.Subject = "Your Invoice From Gateway Group"

                    'Body ---------------------------------------------------------
                    objE.Body = "Attached, please find invoice #" & objI.InvoiceNumber & " from Gateway Group Personnel." & vbCrLf & vbCrLf
                    objE.Body = objE.Body & "You must have Adobe Acrobat Reader to view this PDF file. If you do not currently have this software, you can download it for free at www.Adobe.com." & vbCrLf & vbCrLf
                    objE.Body = objE.Body & "If you have questions about this invoice, please direct them to your Account Executive shown on the invoice. Please do not reply to this email, as it is sent from an unmonitored mailbox." & vbCrLf & vbCrLf
                    objE.Body = objE.Body & "We thank you for your business." & vbCrLf & vbCrLf
                    objE.Body = objE.Body & "Gateway Group Personnel."

                    'Mail Server
                    objE.MailServer = "mail.yswconsulting.com"

                    'Attachment 
                    objE.Attach = m_strPDFPath & strFile & ".pdf"

                    objM.SendEmail(objE)

                    objM = Nothing

                End If

            Catch ex As IOException
                MsgBox("Unable to create a PDF for Invoice #" & objI.InvoiceNumber & ". Please make sure an earlier version of this file is not open.")
                Exit Function

            Catch
                MsgBox("An error occurred while attempting to create a PDF for Invoice #" & objI.InvoiceNumber & ". The message from the system is -- " & _
                        Err.Description & ". If the problem persists, please report this message to Tech Support.", MsgBoxStyle.Exclamation)
                Exit Function
            End Try

            If blExportFile Then
                Me.ExportPDFSToWeb(strFile & ".pdf", strClCode)
            End If


        End Function


        Public Function PrintPageHeaders(ByVal MyPage As ceTe.DynamicPDF.Page, ByVal objClient As ClientInfo, ByVal objInvoice As InvoiceInfo) As Integer

            intCurrentY = 160

            'TOP PART OF INVOICE ---------------------------------------------------------------------------------------
            lblText = New Label("ATTN: " & objClient.Contact, intATTNCol, intCurrentY, 150, 100)
            lblText.FontSize = 10

            'Add label to MyPage
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'CLCODE --------------------------
            lblText = New Label(objClient.ClCode, intCLCodeCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'DATE -----------------------------
            lblText = New Label(objInvoice.InvoiceDate, intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'INV # ----------------------------
            lblText = New Label(objInvoice.InvoiceNumber, intInvCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'PAGE # ---------------------------
            lblText = New Label(objInvoice.CurrentPage & "/" & objInvoice.PageCount, intPageCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'second line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'company name ----------------------
            lblText = New Label(objClient.ClientName, intATTNCol, intCurrentY, 550, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'FED ID ----------------------------
            'lblText = New Label("FED. ID " & Left(m_objCompany.FedID, 6) & "xxxx", intDateCol, intCurrentY, 100, 100)
            'lblText.FontSize = 10
            'MyPage.Elements.Add(lblText)
            'lblText = Nothing

            intCurrentY = intCurrentY + 12


            'NET DAYS
            lblText = New Label("Net " & objInvoice.LateDays & " days", intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'DEPT ------------------------------
            If Not objClient.Dept = vbNullString Then
                lblText = New Label(objClient.Dept, intATTNCol, intCurrentY, 550, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing
                intCurrentY = intCurrentY + 12
            End If


            'client address ----------------------
            Dim strAddress As String = objClient.Address1
            If Not Trim(objClient.Address2) = vbNullString Then strAddress = strAddress & ", " & objClient.Address2
            lblText = New Label(strAddress, intATTNCol, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'fourth line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'client city, state, zip ----------------------
            lblText = New Label(objClient.City & ", " & objClient.State & " " & objClient.Zip, intATTNCol, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'Account Exec -----------------------------------
            lblText = New Label("Account Executive: " & objClient.AccountExec, intDateCol, intCurrentY, 150, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'INVOICE AREA HEADINGS -------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 100

            lblText = New Label("WE DATE", intLowerCol1, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Employee Name", intLowerCol2, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("ST Hours/Rate", intLowerCol3, intCurrentY, 100, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("OT Hours/Rate", intLowerCol4, intCurrentY, 100, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("DT Hours/Rate", intLowerCol5, intCurrentY, 100, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Amt Due", intLowerCol6, intCurrentY, 100, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            PrintPageHeaders = intCurrentY

        End Function

        Public Function PrintMultiJobPageHeaders(ByVal MyPage As ceTe.DynamicPDF.Page, ByVal objClient As ClientInfo, ByVal objInvoice As InvoiceInfo) As Integer

            intCurrentY = 120

            'TOP PART OF INVOICE ---------------------------------------------------------------------------------------
            lblText = New Label("ATTN: " & objClient.Contact, intATTNCol, intCurrentY, 150, 100)
            lblText.FontSize = 10

            'Add label to MyPage
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'CLCODE --------------------------
            lblText = New Label(objClient.ClCode, intCLCodeCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'DATE -----------------------------
            lblText = New Label(objInvoice.InvoiceDate, intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'INV # ----------------------------
            lblText = New Label(objInvoice.InvoiceNumber, intInvCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'PAGE # ---------------------------
            lblText = New Label(objInvoice.CurrentPage & "/" & objInvoice.PageCount, intPageCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'second line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'company name ----------------------
            lblText = New Label(objClient.ClientName, intATTNCol, intCurrentY, 550, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'FED ID ----------------------------
            'lblText = New Label("FED. ID " & Left(m_objCompany.FedID, 6) & "xxxx", intDateCol, intCurrentY, 100, 100)
            'lblText.FontSize = 10
            'MyPage.Elements.Add(lblText)
            'lblText = Nothing

            intCurrentY = intCurrentY + 12

            'NET DAYS
            lblText = New Label("Net " & objInvoice.LateDays & " days", intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'DEPT ------------------------------
            If Not objClient.Dept = vbNullString Then
                lblText = New Label(objClient.Dept, intATTNCol, intCurrentY, 550, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing
                intCurrentY = intCurrentY + 12
            End If


            'client address ----------------------
            Dim strAddress As String = objClient.Address1
            If Not Trim(objClient.Address2) = vbNullString Then strAddress = strAddress & ", " & objClient.Address2
            lblText = New Label(strAddress, intATTNCol, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'fourth line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'client city, state, zip ----------------------
            lblText = New Label(objClient.City & ", " & objClient.State & " " & objClient.Zip, intATTNCol, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'Account Exec -----------------------------------
            lblText = New Label("Account Executive: " & objClient.AccountExec, intDateCol, intCurrentY, 150, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'INVOICE AREA HEADINGS -------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 40

            lblText = New Label("WE DATE", intLowerCol1, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Employee Name", intLowerCol2, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("ST Hours/Rate", intLowerCol3, intCurrentY, 100, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("OT Hours/Rate", intLowerCol4, intCurrentY, 100, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("DT Hours/Rate", intLowerCol5, intCurrentY, 100, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Amt Due", intLowerCol6, intCurrentY, 100, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            PrintMultiJobPageHeaders = intCurrentY

        End Function

        Public Function PrintMultiJobTotalPageHeaders(ByVal MyPage As ceTe.DynamicPDF.Page, ByVal objClient As ClientInfo, ByVal objInvoice As InvoiceInfo) As Integer

            intCurrentY = 120

            'TOP PART OF INVOICE ---------------------------------------------------------------------------------------
            lblText = New Label("ATTN: " & objClient.Contact, intATTNCol, intCurrentY, 150, 100)
            lblText.FontSize = 10

            'Add label to MyPage
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'CLCODE --------------------------
            lblText = New Label(objClient.ClCode, intCLCodeCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'DATE -----------------------------
            lblText = New Label(objInvoice.InvoiceDate, intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'INV # ----------------------------
            lblText = New Label(objInvoice.InvoiceNumber, intInvCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'PAGE # ---------------------------
            lblText = New Label("1/" & objInvoice.PageCount, intPageCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'second line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'company name ----------------------
            lblText = New Label(objClient.ClientName, intATTNCol, intCurrentY, 550, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'FED ID ----------------------------
            'lblText = New Label("FED. ID " & Left(m_objCompany.FedID, 6) & "xxxx", intDateCol, intCurrentY, 100, 100)
            'lblText.FontSize = 10
            'MyPage.Elements.Add(lblText)
            'lblText = Nothing

            intCurrentY = intCurrentY + 12

            'NET DAYS
            lblText = New Label("Net " & objInvoice.LateDays & " days", intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'DEPT ------------------------------
            If Not objClient.Dept = vbNullString Then
                lblText = New Label(objClient.Dept, intATTNCol, intCurrentY, 550, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing
                intCurrentY = intCurrentY + 12
            End If


            'client address ----------------------
            Dim strAddress As String = objClient.Address1
            If Not Trim(objClient.Address2) = vbNullString Then strAddress = strAddress & ", " & objClient.Address2
            lblText = New Label(strAddress, intATTNCol, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'fourth line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'client city, state, zip ----------------------
            lblText = New Label(objClient.City & ", " & objClient.State & " " & objClient.Zip, intATTNCol, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'Account Exec -----------------------------------
            lblText = New Label("Account Executive: " & objClient.AccountExec, intDateCol, intCurrentY, 150, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing



            PrintMultiJobTotalPageHeaders = intCurrentY

        End Function


        Public Sub PrintInvoices(ByVal lngStartNo As Long, ByVal lngEndNo As Long, ByVal blSendEmail As Boolean, Optional ByVal blExportFile As Boolean = False)

            'home workstation key
            'ceTe.DynamicPDF.Document.AddLicense("GEN50NPSCIBLCEgnIxARzvJaAPIY9i6BoStci/fgzy42b4qgzCnS8zbnnNZWG7RdaPZqDEaA2ZcXfp7RFXr1tQrx9pf891Jjpcig")

            'SCHEPP key
            'ceTe.DynamicPDF.Document.AddLicense("GEN50NPSCIBLCEZzlxUndwCT2PBz+MNMdF9kH127Kkhb8Ih6ciWGOt2H6tUAyu34JadrHB27mT3Aix7hviZNEH/onMPC+PM8aGMA")

            'LEE key
            'ceTe.DynamicPDF.Document.AddLicense("GEN50NPSCIBLCEu82ZUSocn1p3hE0ViaZQSaW9zSz45JiNi5CZfq0OHRtpVPJmNjFg9dKY3FpgMrYjcIy31XytjsTRfdIkCJgJNA")
            'Dim MyDocument As ceTe.DynamicPDF.Document = New ceTe.DynamicPDF.Document


            'class library uses the server license key to create the document

            Dim MyDocument As ceTe.DynamicPDF.Document = CreateDocument()
            Dim MyPage As ceTe.DynamicPDF.Page = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)


            MyPage.Elements.Add(New BackgroundImage(m_strImagePath))


            Dim lngFEDID As String = vbNullString
            Dim dbIntPercent As Double = 0
            Dim dLateDate As Date
            Dim intLateDays As Integer = 0
            Dim dbLateFee As Double = 0
            Dim lngAdjStartNum As Long = 0

            Dim lngCurrInvNum As Long = 0
            Dim lngNewInvNum As Long = 0
            Dim intPrintCount As Integer = 0
            Dim intPageCount As Integer = 0, intCurrPageCount As Integer = 0
            Dim dbThisInvoiceTotal As Decimal = 0
            Dim dbThisInvMiscBill As Decimal = 0
            Dim dbTemp As Decimal = 0
            Dim strClientName As String = vbNullString
            Dim strLastClientName As String = vbNullString

            Dim blPrintTotal As Boolean = False

            Dim strInv As String

            Dim objC As ClientInfo = New ClientInfo
            Dim objI As InvoiceInfo = New InvoiceInfo

            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim rs As SqlDataReader = objDB.GetTempInvoices(lngStartNo, lngEndNo)

            m_blPrintingTemp = True

            Dim blMultiJobPrinted As Boolean = False

            m_objCompany = LoadCompanyInfo()

            intPrintCount = 0

            Dim intJobCount As Integer = 0

            While rs.Read

                If rs!onlinetimeslips Is System.DBNull.Value Then
                    m_blOnlineTimeSlips = False
                Else
                    If rs!onlinetimeslips = 1 Then
                        m_blOnlineTimeSlips = True
                    Else
                        m_blOnlineTimeSlips = False
                    End If
                End If

                If rs!hideinvoicepage2 Is System.DBNull.Value Then
                    m_blHidePage2 = False
                Else
                    If rs!hideinvoicepage2 = 1 Then
                        m_blHidePage2 = True
                    Else
                        m_blHidePage2 = False
                    End If
                End If

                'INVOICES CONTAINING MORE THAN 2 JOBS WILL NOW BE CONDENSED

                'this is actually the page count
                intJobCount = GetInvoiceJobCount(rs("INVNUMBER"))

                lngNewInvNum = rs("INVNUMBER")
                strClientName = rs("CLNAME")

                'If lngNewInvNum = 42225 Then
                '    Dim i As String
                '    i = "poo"
                'End If


                If intJobCount > 1 Then

                    'after this module was all set up, CGH decided he wanted the invoices compressed, but ONLY if
                    'the invoices had more than one page. The following shenanigans are inserted to accommodate that request
                    If Not lngNewInvNum = lngCurrInvNum Then
                        If Not lngCurrInvNum = 0 Then
                            'FIRST, PRINT TOTAL FOR LAST INVOICE -------------------------------
                            If blPrintTotal Then
                                PrintInvoiceTotal(MyPage, objC.ChargeInt, objI)

                                'draw previous invoice -------------------------------------------
                                MyDocument.Pages.Insert(0, MyPage)

                                'insert timeslip(s)
                                PrintTimeSlips(objI, MyDocument)

                                strInv = CStr(lngCurrInvNum)

                                SavePDF(MyDocument, objI, strLastClientName, strInv, blSendEmail, blExportFile, objC.ClCode)
                            End If
                        End If
                        PrintMultiPageInvoice(rs("INVNUMBER"), blSendEmail, blExportFile)
                        blPrintTotal = False
                        blMultiJobPrinted = True
                    End If
                Else

                    'record (invoice number) changes here

                    blPrintTotal = True

                    If lngCurrInvNum = 0 Or blMultiJobPrinted Then
                        'PRINT FIRST INVOICE
                        blMultiJobPrinted = False

                        MyDocument = Nothing
                        MyDocument = CreateDocument()

                        MyPage = Nothing

                        MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F) ' new page in current document
                        MyPage.Elements.Add(New BackgroundImage(m_strImagePath))

                        ' headers -----------------------------------------------
                        objI = LoadInvoiceObject(rs)
                        objC = LoadClientObject(rs, False)
                        intCurrentY = PrintPageHeaders(MyPage, objC, objI)

                        intCurrentY = intCurrentY + 20

                        'first job ----------------------------------------------
                        PrintJob(rs, MyPage)
                        objI.InvoiceTotal = objI.InvoiceTotal + rs("OrigInvAmt") + rs("MiscBill")
                        objI.PageSubTotal = rs("OrigInvAmt") + rs("MiscBill")
                        intPrintCount = 1


                    ElseIf lngCurrInvNum = rs("INVNUMBER") Then

                        'same invoice, different job
                        If intPrintCount = 1 Then

                            'print second job on same page
                            PrintJob(rs, MyPage)
                            objI.InvoiceTotal = objI.InvoiceTotal + rs("OrigInvAmt") + rs("MiscBill")
                            objI.PageSubTotal = objI.PageSubTotal + rs("OrigInvAmt") + rs("MiscBill")
                            intPrintCount = 0

                        Else

                            'start new page for same invoice (2 jobs per page)
                            'MyDocument.Pages.Add(MyPage) 'add finished page to currentdocument

                            intCurrentY = intCurrentY + 20

                            'add page subtotal
                            lblText = New Label("PAGE SUB TOTAL", intLowerCol4, intCurrentY, 200, 100)
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                            lblText = New Label(Format(objI.PageSubTotal, "currency"), intLowerCol6, intCurrentY, 100, 100)
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                            objI.PageSubTotal = 0

                            MyDocument.Pages.Insert(0, MyPage)

                            MyPage = Nothing

                            MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F) ' new page in current document
                            MyPage.Elements.Add(New BackgroundImage(m_strImagePath))

                            objI.CurrentPage = objI.CurrentPage - 1
                            intCurrentY = PrintPageHeaders(MyPage, objC, objI)

                            intCurrentY = intCurrentY + 20

                            'print first job on page -----------------
                            PrintJob(rs, MyPage)
                            objI.InvoiceTotal = objI.InvoiceTotal + rs("OrigInvAmt") + rs("MiscBill")
                            objI.PageSubTotal = rs("OrigInvAmt") + rs("MiscBill")

                            intPrintCount = intPrintCount + 1


                        End If



                    Else

                        'START A NEW INVOICE

                        'FIRST, PRINT TOTAL FOR LAST INVOICE -------------------------------
                        If blPrintTotal Then
                            PrintInvoiceTotal(MyPage, objC.ChargeInt, objI)

                            'draw previous invoice -------------------------------------------
                            MyDocument.Pages.Insert(0, MyPage)

                            'insert timeslip(s)
                            PrintTimeSlips(objI, MyDocument)

                            strInv = CStr(lngCurrInvNum)

                            SavePDF(MyDocument, objI, strLastClientName, strInv, blSendEmail, blExportFile, objC.ClCode)
                        End If


                        'START THE NEW INVOICE -------------------------------------------
                        MyDocument = Nothing
                        MyDocument = CreateDocument()

                        MyPage = Nothing

                        MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F) ' new page in current document
                        MyPage.Elements.Add(New BackgroundImage(m_strImagePath))

                        'load new client object getting billing address
                        'from bill address ID in job record

                        'print headers on new page -------------------
                        objI = Nothing
                        objC = Nothing

                        objI = LoadInvoiceObject(rs)
                        objC = LoadClientObject(rs)
                        intCurrentY = PrintPageHeaders(MyPage, objC, objI)

                        intCurrentY = intCurrentY + 10

                        'PRINT FIRST JOB ON NEW INVOICE
                        PrintJob(rs, MyPage)
                        objI.InvoiceTotal = objI.InvoiceTotal + rs("OrigInvAmt") + rs("MiscBill")
                        objI.PageSubTotal = rs("OrigInvAmt") + rs("MiscBill")
                        intPrintCount = 1


                        intCurrentY = intCurrentY + 20

                    End If




                End If


                lngNewInvNum = rs("INVNUMBER")

                strLastClientName = rs("CLNAME")

                lngCurrInvNum = lngNewInvNum


            End While

            If rs.HasRows Then
                If blPrintTotal Then
                    'assume we're at the end and print the last invoice
                    PrintInvoiceTotal(MyPage, objC.ChargeInt, objI)

                    MyDocument.Pages.Insert(0, MyPage)


                    'insert timeslip(s)
                    PrintTimeSlips(objI, MyDocument)

                    strInv = CStr(lngCurrInvNum)
                    SavePDF(MyDocument, objI, strLastClientName, strInv, blSendEmail, blExportFile, objC.ClCode)
                End If


            End If

            PrintTempAdjustmentInvoices(lngStartNo, lngEndNo, blSendEmail)



        End Sub

        Public Sub PrintMultiPageInvoice(ByVal lngInvNumber As Long, ByVal blSendEmail As Boolean, Optional ByVal blExportFile As Boolean = False)

            Dim MyDocument As ceTe.DynamicPDF.Document = New ceTe.DynamicPDF.Document

            Dim MyPage As ceTe.DynamicPDF.Page = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)
            MyPage.Elements.Add(New BackgroundImage(m_strMultiJobImagePath))


            Dim lngFEDID As String = vbNullString
            Dim dbIntPercent As Double = 0
            Dim dLateDate As Date
            Dim intLateDays As Integer = 0
            Dim dbLateFee As Double = 0
            Dim lngAdjStartNum As Long = 0

            Dim lngCurrInvNum As Long = 0
            Dim lngNewInvNum As Long = 0
            Dim intPrintCount As Integer = 0
            Dim intPageCount As Integer = 0, intCurrPageCount As Integer = 0
            Dim dbThisInvoiceTotal As Decimal = 0
            Dim dbThisInvMiscBill As Decimal = 0
            Dim dbTemp As Decimal = 0
            Dim strClientName As String = vbNullString
            Dim strLastClientName As String = vbNullString
            Dim strAssignedTo As String = vbNullString
            Dim strLastAssignedTo As String = vbNullString

            Dim strDocumentName As String = vbNullString
            Dim strFile As String = vbNullString
            Dim strInv As String

            Dim objC As ClientInfo = New ClientInfo
            Dim objI As InvoiceInfo = New InvoiceInfo


            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim rs As SqlDataReader = objDB.GetTempInvoices(lngInvNumber, lngInvNumber)

            m_objCompany = LoadCompanyInfo()

            intPrintCount = 0

            Dim intJobCount As Integer = 0

            While rs.Read
                'record (invoice number) changes here

                lngNewInvNum = rs("INVNUMBER")
                strClientName = rs("CLNAME")
                strAssignedTo = rs("ASSIGNEDTO")

                If lngCurrInvNum = 0 Then
                    'PRINT FIRST INVOICE

                    'headers -----------------------------------------------
                    objI = LoadInvoiceObject(rs, False, True)
                    objC = LoadClientObject(rs, False)
                    intCurrentY = PrintMultiJobPageHeaders(MyPage, objC, objI)

                    intCurrentY = intCurrentY + 20

                    'first job ----------------------------------------------
                    PrintJob(rs, MyPage)
                    objI.InvoiceTotal = objI.InvoiceTotal + rs("OrigInvAmt") + rs("MiscBill")
                    objI.PageSubTotal = rs("OrigInvAmt") + rs("MiscBill")
                    intPrintCount = 1


                ElseIf lngCurrInvNum = rs("INVNUMBER") Then

                    'same invoice, different job
                    If intPrintCount < 4 Then

                        'print second/third job on same page
                        PrintJob(rs, MyPage)
                        objI.InvoiceTotal = objI.InvoiceTotal + rs("OrigInvAmt") + rs("MiscBill")
                        objI.PageSubTotal = objI.PageSubTotal + rs("OrigInvAmt") + rs("MiscBill")
                        intPrintCount = intPrintCount + 1

                    Else

                        'start new page for same invoice (4 jobs per page)

                        intCurrentY = intCurrentY + 20

                        'add page subtotal
                        lblText = New Label("PAGE SUB TOTAL", intLowerCol4, intCurrentY, 200, 100)
                        lblText.Align = TextAlign.Right
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        lblText = New Label(Format(objI.PageSubTotal, "currency"), intLowerCol6, intCurrentY, 100, 100)
                        lblText.Align = TextAlign.Right
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        intCurrentY = intCurrentY + 50

                        lblText = New Label(Format("Continued . . ."), intLowerCol4, 650, 100, 100)
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing


                        objI.PageSubTotal = 0


                        'MyDocument.Pages.Insert(0, MyPage)
                        MyDocument.Pages.Add(MyPage)


                        MyPage = Nothing

                        MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F) ' new page in current document
                        MyPage.Elements.Add(New BackgroundImage(m_strMultiJobImagePath))


                        objI.CurrentPage = objI.CurrentPage + 1
                        intCurrentY = PrintMultiJobPageHeaders(MyPage, objC, objI)

                        intCurrentY = intCurrentY + 20

                        'print first job on page -----------------
                        PrintJob(rs, MyPage)
                        objI.InvoiceTotal = objI.InvoiceTotal + rs("OrigInvAmt") + rs("MiscBill")
                        objI.PageSubTotal = rs("OrigInvAmt") + rs("MiscBill")

                        intPrintCount = 1


                    End If

                End If


                lngNewInvNum = rs("INVNUMBER")

                strLastClientName = rs("CLNAME")

                lngCurrInvNum = lngNewInvNum

            End While

            'print total ---------------------------------------------------------------------
            If rs.HasRows Then

                'add sub-total to final page
                intCurrentY = intCurrentY + 20

                'add page subtotal
                lblText = New Label("PAGE SUB TOTAL", intLowerCol4, intCurrentY, 200, 100)
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label(Format(objI.PageSubTotal, "currency"), intLowerCol6, intCurrentY, 100, 100)
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 50

                lblText = New Label(Format("Continued . . ."), intLowerCol4, 650, 100, 100)
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                objI.PageSubTotal = 0


                'draw previous invoice -------------------------------------------
                MyDocument.Pages.Add(MyPage)

                PrintMultiJobInvoiceTotal(MyDocument, objC.ChargeInt, objI, objC)


                'insert timeslip(s)
                PrintTimeSlips(objI, MyDocument)

                strInv = CStr(lngCurrInvNum)

                SavePDF(MyDocument, objI, strLastClientName, strInv, blSendEmail, blExportFile, objC.ClCode)


            End If

        End Sub

        Private Sub PrintJob(ByVal rs As SqlDataReader, ByVal MyPage As ceTe.DynamicPDF.Page)

            Dim dcJobTotal As Decimal = 0
            Static strLastAssignedTo As String

            intCurrentY = intCurrentY + 20

            If Not strLastAssignedTo = rs("ASSIGNEDTO") Then
                lblText = New Label("EMPLOYEES ASSIGNED TO: " & UCase(rs("ASSIGNEDTO")), intLowerCol1, intCurrentY, 300, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing
                intCurrentY = intCurrentY + 20
            End If

            strLastAssignedTo = rs("ASSIGNEDTO")

            'WE DATE ----------------------
            lblText = New Label(rs("WEDate"), intLowerCol1, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'Employee Name
            lblText = New Label(rs("FNAME") & " " & rs("LNAME"), intLowerCol2, intCurrentY, 100, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Left
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'ST HOURS/RATE
            lblText = New Label(Format(rs("BILLST"), "###.00") & "  " & Format(rs("STBILLRATE"), "currency"), intLowerCol3, intCurrentY, 100, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'OT HOURS/RATE
            If Not rs("PAYOT") = 0 Then
                lblText = New Label(Format(rs("BILLOT"), "###.00") & "  " & Format(rs("OTBillRate"), "currency"), intLowerCol4, intCurrentY, 100, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If

            If Not rs("PAYDT") = 0 Then
                lblText = New Label(Format(rs("BILLDT"), "###.00") & "  " & Format(rs("DTBillRate"), "currency"), intLowerCol5, intCurrentY, 100, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If

            'Amt Due
            If Not rs("OrigInvAmt") = 0 Then
                lblText = New Label(Format(rs("OrigInvAmt"), "currency"), intLowerCol6, intCurrentY, 100, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing

            End If

            

            'dbThisInvoiceTotal = dbThisInvoiceTotal + rs!originvamt
            'dbThisInvMiscBill = dbThisInvMiscBill + RidNULL(rs!miscbill)



            'Jobnumber
            intCurrentY = intCurrentY + 15
            lblText = New Label("PRL #: " & rs("jobnumber"), intLowerCol1, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing



            'PO NUMBER

            Dim objC As ClientInfo = LoadClientObject(rs)

            If Not Trim(objC.PONumber) = "" Then

                intCurrentY = intCurrentY + 15
                lblText = New Label("PO/Dept #: " & objC.PONumber, intLowerCol1, intCurrentY, 500, 100)
                lblText.FontSize = 10
                        lblText.Align = TextAlign.Left
                        MyPage.Elements.Add(lblText)
                lblText = Nothing

            End If




            'Position
            intCurrentY = intCurrentY + 15
            lblText = New Label(rs("typeassign"), intLowerCol1, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            If Not rs("miscbill") = 0 Then

                intCurrentY = intCurrentY + 15
                lblText = New Label("Misc. Billing: " & Trim(rs!reason), intLowerCol4, intCurrentY, 300, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label(Format(rs("miscbill"), "currency"), intLowerCol6, intCurrentY, 100, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing

            End If



            'Assigned To
            'intCurrentY = intCurrentY + 15
            'lblText = New Label("Assigned to: " & rs("AssignedTo"), intLowerCol1, intCurrentY, 300, 100)
            'lblText.FontSize = 10
            'MyPage.Elements.Add(lblText)
            'lblText = Nothing



            intCurrentY = intCurrentY + 30

        End Sub

       
        Private Sub PrintInvoiceTotal(ByVal MyPage As ceTe.DynamicPDF.Page, ByVal blChargeInt As Boolean, ByVal objI As InvoiceInfo)

            intCurrentY = intCurrentY + 5

            'add page subtotal
            'lblText = New Label("PAGE SUB TOTAL", intLowerCol4, intCurrentY, 200, 100)
            'lblText.Align = TextAlign.Right
            'MyPage.Elements.Add(lblText)
            'lblText = Nothing

            'lblText = New Label(Format(objI.PageSubTotal, "currency"), intLowerCol6, intCurrentY, 100, 100)
            'lblText.Align = TextAlign.Right
            'MyPage.Elements.Add(lblText)
            'lblText = Nothing

            objI.PageSubTotal = 0

            intCurrentY = intCurrentY + 20

            If Not m_blOnlineTimeSlips Then
                lblText = New Label("***** TIME SLIPS ENCLOSED *****", intLowerCol1, intCurrentY, 200, 100)
            Else
                lblText = New Label("Time entered and approved via web", intLowerCol1, intCurrentY, 200, 100)
            End If

            MyPage.Elements.Add(lblText)
            lblText = Nothing


            lblText = New Label("GRAND TOTAL INVOICE AMOUNT", intLowerCol4, intCurrentY, 200, 100)
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label(Format(objI.InvoiceTotal, "currency"), intLowerCol6, intCurrentY, 100, 100)
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 20

            '--------------------------------------------------------------------------------------------------------------------------------

            'GLOBAL MESSAGE
            If Not m_objCompany.GlobalStatementMessage Is Nothing Then lblText = New Label(m_objCompany.GlobalStatementMessage, intLowerCol1, intCurrentY, 1000, 100)
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 20

            lblText = New Label("Please direct any questions pertaining to this invoice to your account executive shown above.", intLowerCol1, intCurrentY, 500, 300)
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 20

            lblText = New Label("** This invoice represents a payroll related matter. Please Process Promptly.", intLowerCol1, intCurrentY, 500, 300)
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'figure out whether or not to charge interest ---------------------------------------------------------------------------------
            If blChargeInt Then
                If m_objCompany.InterestStartDate <> "1/1/1900" And m_objCompany.InterestStartDate <> "12:00:00 AM" And objI.InvoiceDate >= m_objCompany.InterestStartDate Then

                    Dim dLateDate As Date = DateAdd("d", objI.LateDays, objI.InvoiceDate)

                    If Not m_objCompany.DailyInterestPercent = 0 Then

                        Dim dbLateFee As Decimal = (objI.InvoiceTotal) * m_objCompany.DailyInterestPercent


                        Dim strPayLate As String = "If paid after " & dLateDate & " interest will be compounded daily on the unpaid "
                        strPayLate = strPayLate & " balance at the rate of " & _
                            CStr(m_objCompany.AnnualInterestPercent * 100) & "% annually  (" & m_objCompany.DailyInterestPercent & "% Daily)"


                        If intCurrentY < 650 Then
                            intCurrentY = intCurrentY + 20
                        Else
                            intCurrentY = intCurrentY + 15
                        End If


                        'DISPLAY INTEREST MESSAGE
                        lblText = New Label(strPayLate, intLowerCol1, intCurrentY, 500, 300)
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing


                    End If
                    End If
                End If



        End Sub

        Private Sub PrintMultiJobInvoiceTotal(ByVal MyDocument As ceTe.DynamicPDF.Document, ByVal blChargeInt As Boolean, ByVal objI As InvoiceInfo, ByVal objC As ClientInfo)

            Dim myPage As ceTe.DynamicPDF.Page = New ceTe.DynamicPDF.Page
            myPage.Elements.Add(New BackgroundImage(m_strMultiJobImagePath))

            intCurrentY = PrintMultiJobTotalPageHeaders(myPage, objC, objI)

            intCurrentY = intCurrentY + 50


            lblText = New Label("GRAND TOTAL INVOICE AMOUNT", intLowerCol1, intCurrentY, 200, 100)
            lblText.Align = TextAlign.Right
            myPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label(Format(objI.InvoiceTotal, "currency"), intLowerCol6, intCurrentY, 100, 100)
            lblText.Align = TextAlign.Right
            myPage.Elements.Add(lblText)
            lblText = Nothing


            intCurrentY = intCurrentY + 100

            lblText = New Label("***** TIME SLIPS ENCLOSED *****", intLowerCol1, intCurrentY, 200, 100)
            myPage.Elements.Add(lblText)
            lblText = Nothing





            intCurrentY = intCurrentY + 100

            '--------------------------------------------------------------------------------------------------------------------------------

            'GLOBAL MESSAGE
            If Not m_objCompany.GlobalStatementMessage Is Nothing Then lblText = New Label(m_objCompany.GlobalStatementMessage, intLowerCol1, intCurrentY, 1000, 100)
            myPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 30

            lblText = New Label("Please direct any questions pertaining to this invoice to your account executive shown above.", intLowerCol1, intCurrentY, 500, 300)
            myPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 30

            lblText = New Label("** This invoice represents a payroll related matter. Please Process Promptly.", intLowerCol1, intCurrentY, 500, 300)
            myPage.Elements.Add(lblText)
            lblText = Nothing

            'figure out whether or not to charge interest ---------------------------------------------------------------------------------
            If blChargeInt Then
                If m_objCompany.InterestStartDate <> "1/1/1900" And m_objCompany.InterestStartDate <> "12:00:00 AM" And objI.InvoiceDate >= m_objCompany.InterestStartDate Then

                    Dim dLateDate As Date = DateAdd("d", objI.LateDays, objI.InvoiceDate)

                    If Not m_objCompany.DailyInterestPercent = 0 Then

                        Dim dbLateFee As Decimal = (objI.InvoiceTotal) * m_objCompany.DailyInterestPercent


                        Dim strPayLate As String = "If paid after " & dLateDate & " interest will be compounded daily on the unpaid "
                        strPayLate = strPayLate & " balance at the rate of " & _
                            CStr(m_objCompany.AnnualInterestPercent * 100) & "% annually  (" & m_objCompany.DailyInterestPercent & "% Daily)"


                        If intCurrentY < 650 Then
                            intCurrentY = intCurrentY + 20
                        Else
                            intCurrentY = intCurrentY + 15
                        End If


                        'DISPLAY INTEREST MESSAGE
                        lblText = New Label(strPayLate, intLowerCol1, intCurrentY, 500, 300)
                        myPage.Elements.Add(lblText)
                        lblText = Nothing


                    End If
                End If
            End If

            MyDocument.Pages.Insert(0, myPage)

        End Sub

        Private Sub PrintTimeSlips(ByVal objI As InvoiceInfo, ByRef MyDocument As ceTe.DynamicPDF.Document)

            'get timeslips for this invoice
            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim reader As SqlDataReader = objDB.GetInvoiceJobIDs(objI.InvoiceNumber)
            Dim objTS As GGBackOffice.TimeSlipInfo = New TimeSlipInfo
            Dim strTemp As String

            Dim intSlipsPerPage As Integer = 2
            Dim intSlipsPrinted As Integer = 0
            Dim MyPage As ceTe.DynamicPDF.Page = New Page
            MyPage.Elements.Add(New BackgroundImage(m_strTSImagePath))

            lblText = New Label("TIME SLIPS", 10, 75, 200, 100)
            lblText.FontSize = 14
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            objI.CurrentPage = objI.PageCount - (m_intTSCount)

            'PAGE # ---------------------------
            lblText = New Label("Page " & objI.CurrentPage & "/" & objI.PageCount, 500, 75, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            If reader Is Nothing Then Exit Sub

            While reader.Read

                objTS = New TimeSlipInfo


                intCurrentY = intCurrentY + 50
                objTS = objTS.GetSavedTimeSlip(reader("pk_id"))

                'print time slip, sending objTS
                If Not objTS Is Nothing Then

                    'count no more than 3 timeslips per page and send the page
                    If intSlipsPrinted < intSlipsPerPage Then
                        PrintTimeSlip(MyPage, objTS, objI, intSlipsPrinted)
                        intSlipsPrinted = intSlipsPrinted + 1
                    Else
                        intCurrentY = intCurrentY + 40

                        lblText = New Label(Format("Continued . . ."), intLowerCol4, 650, 100, 100)
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        MyDocument.Pages.Add(MyPage)
                        MyPage = Nothing
                        objI.CurrentPage = objI.CurrentPage + 1
                        intSlipsPrinted = 0
                        MyPage = New ceTe.DynamicPDF.Page
                        MyPage.Elements.Add(New BackgroundImage(m_strTSImagePath))

                        lblText = New Label("TIME SLIPS", 10, 75, 200, 100)
                        lblText.FontSize = 14
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        'PAGE # ---------------------------
                        lblText = New Label("Page " & objI.CurrentPage & "/" & objI.PageCount, 500, 75, 100, 100)
                        lblText.FontSize = 10
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        PrintTimeSlip(MyPage, objTS, objI, intSlipsPrinted)
                        intSlipsPrinted = 1
                    End If

                End If

                objTS = Nothing


            End While


            If intSlipsPrinted > 0 Then

                lblText = New Label(Format("Continued . . ."), intLowerCol4, 650, 100, 100)
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                MyDocument.Pages.Add(MyPage)
            End If

            If Not m_blHidePage2 Then PrintTimeSlipBackPage(MyDocument)

        End Sub

        Private Sub PrintTimeSlip(ByRef MyPage As ceTe.DynamicPDF.Page, ByVal objTS As GGBackOffice.TimeSlipInfo, ByVal objI As InvoiceInfo, ByVal intSlipsPrinted As Integer)
            Dim objDay As GGBackOffice.TimeSlipInfo.GGTimeSlipDay = New GGBackOffice.TimeSlipInfo.GGTimeSlipDay

            'days could be lower or upper case
            Dim hsh As Hashtable = Collections.Specialized.CollectionsUtil.CreateCaseInsensitiveHashtable
            Dim lstDates As SortedList = Collections.Specialized.CollectionsUtil.CreateCaseInsensitiveSortedList


            'load dates into SortedList
            Dim weDate As Date = objTS.WEDate
            lstDates.Add("Sunday", weDate)
            lstDates.Add("Saturday", DateAdd(DateInterval.Day, -1, weDate))
            lstDates.Add("Friday", DateAdd(DateInterval.Day, -2, weDate))
            lstDates.Add("Thursday", DateAdd(DateInterval.Day, -3, weDate))
            lstDates.Add("Wednesday", DateAdd(DateInterval.Day, -4, weDate))
            lstDates.Add("Tuesday", DateAdd(DateInterval.Day, -5, weDate))
            lstDates.Add("Monday", DateAdd(DateInterval.Day, -6, weDate))


            'columns
            Dim intDayCol As Integer = 10
            Dim intDateCol As Integer = 100
            Dim intTimeInCol As Integer = 175
            Dim intTimeOutCol As Integer = 225
            Dim intLunchCol As Integer = 300
            Dim intTotalCol As Integer = 320
            Dim ccTotalHours As Decimal



            Dim dThisDate As Date
            Dim x As Integer = 0

            If intSlipsPrinted = 0 Then
                intCurrentY = 100
            Else
                intCurrentY = intCurrentY + 30
            End If



            lblText = New Label("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", intDayCol, intCurrentY, 800, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing
            intCurrentY = intCurrentY + 20

            'print headings
            lblText = New Label("EMPLOYEE: " & objTS.EmployeeName, intDayCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("DEPT: " & objTS.Dept, intTimeInCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("JOB TITLE: " & objTS.JobTitle, intLunchCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            intCurrentY = intCurrentY + 20

            'Client + clcode, Invnumber & Date
            lblText = New Label("Client: " & objTS.ClientName & "  ACCT #" & objTS.ClientNumber, intDayCol, intCurrentY, 400, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Invoice: " & objI.InvoiceNumber & "  Date: " & objI.InvoiceDate, intLunchCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("W/E: " & objTS.WEDate, intTotalCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 30

            lblText = New Label("DAY", intDayCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("DATE", intDateCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("TIME IN", intTimeInCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("TIME OUT", intTimeOutCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("LUNCH", intLunchCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("TOTAL HOURS", intTotalCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 10


            lblText = New Label("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", intDayCol, intCurrentY, 800, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            For Each objDay In objTS.QDays

                intCurrentY = intCurrentY + 15


                'Select Case x

                '    Case 0
                '        objDay = objTS.HashDays.Item("Monday")
                '    Case 1
                '        objDay = objTS.HashDays.Item("Tuesday")
                '    Case 2
                '        objDay = objTS.HashDays.Item("Wednesday")
                '    Case 3
                '        objDay = objTS.HashDays.Item("Thursday")
                '    Case 4
                '        objDay = objTS.HashDays.Item("Friday")
                '    Case 5
                '        objDay = objTS.HashDays.Item("Saturday")
                '    Case 6
                '        objDay = objTS.HashDays.Item("Sunday")

                'End Select


                lblText = New Label(UCase(objDay.Day), intDayCol, intCurrentY, 200, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                dThisDate = lstDates.Item(objDay.Day)
                lblText = New Label(dThisDate, intDateCol, intCurrentY, 200, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label(objDay.TimeIn, intTimeInCol, intCurrentY, 200, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label(objDay.TimeOut, intTimeOutCol, intCurrentY, 200, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label(objDay.Lunch, intLunchCol, intCurrentY, 200, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label(Format(objDay.TotalHours, "#.00"), intTotalCol, intCurrentY, 200, 100)
                ccTotalHours = ccTotalHours + objDay.TotalHours
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing

            Next

            intCurrentY = intCurrentY + 20


            'GRAND TOTAL
            lblText = New Label("TOTAL HOURS: " & Format(ccTotalHours, "#.00"), intTotalCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 30

            lblText = New Label("APPROVED BY: " & objTS.ApprovedBy, intDayCol, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("APPROVED ON: " & Format(objTS.ApprovedOn, "short date"), intTimeOutCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 10


            lblText = New Label("-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", intDayCol, intCurrentY, 800, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


        End Sub
        Private Sub PrintTimeSlipBackPage(ByRef MyDocument As ceTe.DynamicPDF.Document)

            Dim lblHead As ceTe.DynamicPDF.PageElements.Label

            Dim SR As New StreamReader(m_strTimeSlipTextPath)
            Dim txt As String = SR.ReadToEnd
            SR.Close()

            Dim MyPage As ceTe.DynamicPDF.Page = New Page

            'title -------------------------------------------------------------------------------------------------
            lblHead = New ceTe.DynamicPDF.PageElements.Label("CLIENT INFORMATION / AGREEMENT", 125, 10, 400, 100, Font.HelveticaBold)
            MyPage.Elements.Add(lblHead)
            lblHead = Nothing

            'agreement text
            lblText = New Label(txt, 5, 30, 500, 1000, Font.Helvetica, 10, TextAlign.Justify, RgbColor.Black)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'footer ------------------------------------------------------------------------------------------------
            lblHead = New ceTe.DynamicPDF.PageElements.Label("Rev 9/06 GATEWAY GROUP PERSONNEL", 125, 400, 400, 100, Font.HelveticaBold)
            MyPage.Elements.Add(lblHead)

            MyDocument.Pages.Add(MyPage)

        End Sub


#End Region

#Region "Shared Functions"

        Public Function CreateDocument2() As ceTe.DynamicPDF.Document

            Dim strKey As String = ""

            'new license key obtained 8/25/17
            strKey = "GEN60NPDPMICPPUQgGPqXqGQ0punmq3ZVJtC3Qb2visVeZBBIlc+8A2k7IF7DQ542IBTLAkKTcdeJTqZCmjMoaiDS6xJ7302X6Hw"

            ceTe.DynamicPDF.Document.AddLicense(strKey)

            Dim MyDocument As ceTe.DynamicPDF.Document = New ceTe.DynamicPDF.Document
            Return MyDocument

        End Function



        Public Function CreateDocument() As ceTe.DynamicPDF.Document

            'Dim strKey As String = vbNullString

            'Select Case UCase(My.Computer.Name)

            '    Case "YVETTE"
            '        strKey = "GEN50NPSCIBLCEgnIxARzvJaAPIY9i6BoStci/fgzy42b4qgzCnS8zbnnNZWG7RdaPZqDEaA2ZcXfp7RFXr1tQrx9pf891Jjpcig"
            '    Case "TEMPSPARE"
            '        strKey = "GEN50NPSCIBLCENifAkSQIuABHygkumKBVM1QcoZDrtS2L1pWIcsdnttkmnqeYLMRTFXO7hv0JDAqhcTWRCYRi70YO3tK5bZ2JBA"
            '    Case "PATRICIANEW"
            '        strKey = "GEN50NPSCIBLCEWeyeCTvwFFcNWnaCIgMQczcUgCidDsTnTW5zQQbt7wfFnjgNEdXdsMkiMl6KtiEI8NtqidNQ1+qUQeqOIeekOg"
            '    Case "LEE"
            '        strKey = "GEN50NPSCIBLCEu82ZUSocn1p3hE0ViaZQSaW9zSz45JiNi5CZfq0OHRtpVPJmNjFg9dKY3FpgMrYjcIy31XytjsTRfdIkCJgJNA"
            '    Case "TLW"
            '        strKey = "GEN50NPSCIBLCEuDWrguypmrw8kjoicx0PyKhpbqPg/eOBRdvyqPhdXJU0vI5aTMMn52SDygmtQvc+QgfYIAt/HZk8kUwYrU6UtA"
            '    Case "MAILSCAN2"
            '        strKey = "GEN50NPSCIBLCEMZVSN3k4XNMW58Hrc+FSqia9Bx8BExY7WMmCn42OLvBufL32AhRVs+hm8zhe/DoxosAG6JRpW76fUC/vtZns0g"
            '    Case "DEVWIN7"
            '        strKey = "GEN50NPSCIBLCEwIIgENV+wWLYPncy41mG9mKuo2upr+CTFtLDNGUNlRSSQE0NhzGrJtaNPa4vh0DnhIY8YP8ikutwKtP1umretg"

            'End Select

            ''try new license key
            'strKey = "GEN60NPDPMICPPUQgGPqXqGQ0punmq3ZVJtC3Qb2visVeZBBIlc+8A2k7IF7DQ542IBTLAkKTcdeJTqZCmjMoaiDS6xJ7302X6Hw"

            'Dim objGG As BackOfficeNETClassLibrary.GG.BAKNET.Classes.DocumentCreator = New BackOfficeNETClassLibrary.GG.BAKNET.Classes.DocumentCreator
            'Dim MyDocument As ceTe.DynamicPDF.Document = objGG.CreatePDFDocument(strKey)



            'getting rid of the BackOfficeNETClassLibrary and using new key that will hopefully work on Francille
            Dim strKey As String = ""

            'new license key obtained 8/25/17
            strKey = "GEN60NPDPMICPPUQgGPqXqGQ0punmq3ZVJtC3Qb2visVeZBBIlc+8A2k7IF7DQ542IBTLAkKTcdeJTqZCmjMoaiDS6xJ7302X6Hw"

            ceTe.DynamicPDF.Document.AddLicense(strKey)

            Dim MyDocument As ceTe.DynamicPDF.Document = New ceTe.DynamicPDF.Document

            MyDocument.Creator = "Gateway Group Personnel"
            MyDocument.Author = "Gateway Group Personnel"
            MyDocument.Title = "Invoice"


            Return MyDocument

        End Function

        Private Function LoadClientObject(ByVal rs As SqlDataReader, Optional ByVal blIsAdjust As Boolean = False, Optional ByVal blIsPerm As Boolean = False) As ClientInfo

            Dim objC As ClientInfo = New ClientInfo
            Dim objDb As GGDatabaseController = New GGDatabaseController
            Dim rsBill As SqlDataReader
            'determine whether to use address from billing address table or address in client record

            If rs("billaddressID") Is System.DBNull.Value Then

                objC.Address1 = rs("InvAdd1")
                objC.Address2 = rs("InvAdd2")
                objC.City = rs("InvCity")
                objC.State = rs("InvState")
                objC.Zip = rs("InvZip")
                objC.Contact = rs("InvContact")

                If Not rs("invdept") Is System.DBNull.Value Then
                    objC.Dept = rs("invdept")
                Else
                    objC.Dept = vbNullString
                End If

                If Not rs("billemail") Is System.DBNull.Value Then objC.BillEmail = rs("billemail")

                If m_blPrintingTemp Then
                    objC.PONumber = RidNull(rs("thispo"))
                Else
                    objC.PONumber = RidNull(rs("permpo"))
                End If


            Else

                    If rs("billaddressID") > 0 Then
                    'there's a billing address ID in the job record
                    rsBill = objDb.GetGABillingAddress(rs("billaddressID"))

                    If Not rsBill.HasRows Then
                        'no billing address found for billing address ID
                        objC.Address1 = rs("InvAdd1")
                        objC.Address2 = rs("InvAdd2")
                        objC.City = rs("InvCity")
                        objC.State = rs("InvState")
                        objC.Zip = rs("InvZip")
                        objC.Contact = rs("invcontact")

                        If Not rs("invdept") Is System.DBNull.Value Then
                            objC.Dept = rs("invdept")
                        Else
                            objC.Dept = vbNullString
                        End If

                        objC.BillEmail = rs("billemail")

                        If m_blPrintingTemp Then
                            objC.PONumber = RidNull(rs("thispo"))
                        Else
                            objC.PONumber = RidNull(rs("permpo"))
                        End If

                    Else

                        'billing address found
                        rsBill.Read()

                        objC.ClientName = rs("CLNAME")
                        objC.Address1 = rsBill("ADDRESS1")
                        objC.Address2 = rsBill("Address2")
                        objC.City = rsBill("City")
                        objC.State = rsBill("State")
                        objC.Zip = rsBill("Zip")

                        objC.Contact = rsBill("Attn")

                        objC.PONumber = rsBill("PONumber")

                        If Not rsBill("dept") Is System.DBNull.Value Then
                            objC.Dept = rsBill("dept")
                        Else
                            objC.Dept = vbNullString
                        End If

                        If Not rsBill("email") Is System.DBNull.Value Then objC.BillEmail = rsBill("email")

                        'only use billing address PO if a PO has not been entered in the Job Order
                        If RidNull(rs!thispo) = vbNullString Then

                            objC.PONumber = RidNull(rsBill("ponumber"))

                        Else
                            If m_blPrintingTemp Then
                                objC.PONumber = RidNull(rs("thispo"))
                            Else
                                objC.PONumber = RidNull(rs("permpo"))
                            End If
                        End If



                        rsBill.Close()
                        rsBill = Nothing

                    End If

                Else
                    'no billing address ID in job record
                    objC.Address1 = rs("InvAdd1")
                    objC.Address2 = rs("InvAdd2")
                    objC.City = rs("InvCity")
                    objC.State = rs("InvState")
                    objC.Zip = rs("InvZip")
                    objC.Contact = rs("invcontact")

                    If Not rs("invdept") Is System.DBNull.Value Then
                        objC.Dept = rs("invdept")
                    Else
                        objC.Dept = vbNullString
                    End If

                    objC.BillEmail = rs("billemail")

                    If m_blPrintingTemp Then
                        objC.PONumber = RidNull(rs("thispo"))
                    Else
                        objC.PONumber = RidNull(rs("permpo"))
                    End If

                End If


            End If

            objC.ClientName = rs("CLNAME")


            If rs("charge_int") = "Y" Then
                objC.ChargeInt = True
            Else
                objC.ChargeInt = False
            End If

            objC.ClCode = rs!clcode

            If blIsPerm Then
                objC.AccountExec = rs("Consultant")
                objC.Contact = rs("ATTN")
            ElseIf Not blIsAdjust Then
                Try
                    objC.AccountExec = rs("AcctExec")
                Catch
                    objC.AccountExec = ""
                End Try
            End If

            objDb = Nothing

            Return objC



        End Function

        Private Function LoadInvoiceObject(ByVal rs As SqlDataReader, Optional ByVal blIsPermAdjust As Boolean = False, Optional ByVal blIsMultiJob As Boolean = False) As InvoiceInfo
            Dim objI As InvoiceInfo = New InvoiceInfo

            If Not blIsPermAdjust Then
                With objI
                    .InvoiceDate = rs("INVDATE")
                    .InvoiceNumber = rs("INVNUMBER")

                    If Not blIsMultiJob Then
                        .PageCount = GetPageCount(rs("INVNUMBER"))
                    Else
                        .PageCount = GetMultiJobPageCount(rs("INVNUMBER"))
                    End If


                    If Not blIsMultiJob Then
                        .CurrentPage = 1
                    Else
                        .CurrentPage = 2
                    End If
                    .PrintCopy = False    'need to throw up a dialog or send in params
                    .LateDays = rs("LATEDAYS")
                End With

            Else
                With objI
                    .CurrentPage = 1
                    .InvoiceDate = rs("ADJINVDATE")
                    .InvoiceNumber = rs("ADJINVNUMBER")
                    .PageCount = 1
                    .CurrentPage = 1
                    .PrintCopy = False    'need to throw up a dialog or send in params
                    .LateDays = False

                End With
            End If

            Return objI
        End Function

        Private Function LoadCompanyInfo() As CompanyInfo

            Dim objCompany As CompanyInfo = New CompanyInfo
            Dim objDB As New GGDatabaseController
            Dim rs As SqlDataReader = objDB.GetCompanyInformation
            Dim strInterest As String = vbNullString
            Dim dbDailyInt As Decimal = 0
            Dim dbIntPercent As Decimal = 0

            rs.Read()

            If Not rs("Z_INTPERC") = 0 Then
                dbIntPercent = rs!Z_intperc
                dbIntPercent = dbIntPercent / 100

                dbDailyInt = dbIntPercent * 100
                dbDailyInt = dbDailyInt / 365
                dbDailyInt = Math.Round(dbDailyInt, 4)
            Else
                dbIntPercent = 0
            End If


            With objCompany
                .FedID = rs("z_fedid")
                .GlobalStatementMessage = rs("gbinvoicemsg")
                .InterestStartDate = rs("InterestStart")
                .DailyInterestPercent = dbDailyInt
                .AnnualInterestPercent = dbIntPercent
            End With

            rs = Nothing
            objDB = Nothing

            Return objCompany


        End Function
        Public Function GetInvoiceJobCount(ByVal lngInvNumber As Long) As Integer

            'this will have to run a sproc that counts all jobs on the invoice and divides by 2
            Dim intJobCount As Integer = 0
            GetInvoiceJobCount = 0

            Dim objDB As GGDatabaseController = New GGDatabaseController
            intJobCount = objDB.GetInvoicePageCount(lngInvNumber)

            Return intJobCount

        End Function

        Private Function GetMultiJobPageCount(ByVal lngInvNumber As Long) As Integer
            Dim intPageCount As Integer = 0
            Dim strTSCount As Decimal = 0, intTSCount As Integer = 0, dcTSCount As Decimal
            Dim strArgs() As String
            Dim intJobCount As Integer = 0
            Dim dcJobCount As Decimal

            Dim objDB As GGDatabaseController = New GGDatabaseController
            intJobCount = objDB.GetInvoiceJobCount(lngInvNumber)
            'intTSCount = objDB.GetTimeSlipCount(lngInvNumber)
            objDB = Nothing

            If intJobCount > 2 Then
                dcJobCount = intJobCount / 4
                strArgs = Split(dcJobCount, ".")

                intPageCount = intPageCount + strArgs(0)
              
                If UBound(strArgs) = 1 Then
                    intPageCount = intPageCount + 1
                End If

            Else
                intPageCount = intPageCount + 1
            End If



            'if more than 2 jobs, time slips will be condensed to fit 3 to a page
            If intTSCount > 2 Then
                dcTSCount = intTSCount / 2
                strTSCount = CStr(dcTSCount)
                strArgs = Split(strTSCount, ".")

                intPageCount = intPageCount + strArgs(0)
                m_intTSCount = strArgs(0)

                If UBound(strArgs) = 1 Then
                    intPageCount = intPageCount + 1
                    m_intTSCount = m_intTSCount + 1
                End If

            Else
                intPageCount = intPageCount + 1
            End If

            'add 1 for time slip back page text
            intPageCount = intPageCount + 1

            'add 1 for separate total page 
            intPageCount = intPageCount + 1

            Return intPageCount
        End Function

        Private Function GetPageCount(ByVal lngInvNumber As Long) As Integer

            'this will have to run a sproc that counts all jobs on the invoice and divides by 2
            Dim intPageCount As Integer = 0
            Dim strTSCount As Decimal = 0, intTSCount As Integer = 0, dcTSCount As Decimal
            Dim strArgs() As String
            GetPageCount = 0

            'Dim objDB As GGDatabaseController = New GGDatabaseController
            'intPageCount = objDB.GetInvoicePageCount(lngInvNumber)


            'intTSCount = objDB.GetTimeSlipCount(lngInvNumber)
            'm_intTSCount = intTSCount

            'objDB = Nothing

            ''if more than 2 jobs, time slips will be condensed to fit 3 to a page
            'If intTSCount > 2 Then
            '    dcTSCount = intTSCount / 3
            '    strTSCount = CStr(dcTSCount)
            '    strArgs = Split(strTSCount, ".")

            '    intPageCount = intPageCount + strArgs(0)
            '    m_intTSCount = strArgs(0)

            '    If UBound(strArgs) = 1 Then
            '        intPageCount = intPageCount + 1
            '        m_intTSCount = m_intTSCount + 1
            '    End If

            'Else
            '    'intPageCount = intPageCount + 1
            'End If

            'add 1 for time slip back page text
            intPageCount = intPageCount + 1

            Return intPageCount


        End Function

#End Region


#Region "Temp Adjustments"

        Public Function PrintTempAdjustmentInvoices(ByVal lngStartNo As Long, ByVal lngEndNo As Long, ByVal blSendEmail As Boolean)

            Dim MyDocument As ceTe.DynamicPDF.Document = CreateDocument()

            Dim MyPage As ceTe.DynamicPDF.Page = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)

            MyPage.Elements.Add(New BackgroundImage(m_strImagePath))


            Dim lngFEDID As String = vbNullString
            Dim dbIntPercent As Double = 0
            Dim intLateDays As Integer = 0
            Dim dbLateFee As Double = 0
            Dim lngAdjStartNum As Long = 0

            Dim lngCurrInvNum As Long = 0
            Dim lngNewInvNum As Long = 0
            Dim intPrintCount As Integer = 0
            Dim intPageCount As Integer = 0, intCurrPageCount As Integer = 0
            Dim dbThisInvoiceTotal As Decimal = 0
            Dim dbThisInvMiscBill As Decimal = 0
            Dim dbTemp As Decimal = 0
            Dim strClientName As String = vbNullString
            Dim strLastClientName As String = vbNullString

            Dim strDocumentName As String = vbNullString
            Dim strFile As String = vbNullString


            Dim objC As ClientInfo = New ClientInfo
            Dim objI As InvoiceInfo = New InvoiceInfo

            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim rs As SqlDataReader = objDB.GetTempAdjustmentInvoices(lngStartNo, lngEndNo)

            m_objCompany = LoadCompanyInfo()

            intPrintCount = 0



            While rs.Read

                strClientName = rs("CLNAME")

                If lngCurrInvNum = 0 Then
                    'PRINT FIRST INVOICE

                    ' headers -----------------------------------------------
                    objI = LoadInvoiceObject(rs)
                    objC = LoadClientObject(rs, True)
                    intCurrentY = PrintAdjustmentPageHeaders(MyPage, objC, objI, rs)

                    intCurrentY = intCurrentY + 10

                    'first job ----------------------------------------------
                    PrintAdjustmentJob(rs, MyPage)
                    objI.InvoiceTotal = objI.InvoiceTotal + rs("InvAdjust")
                    intPrintCount = 1


                ElseIf lngCurrInvNum = rs("INVNUMBER") Then

                    'same invoice, different job
                    If intPrintCount = 1 Then

                        'print second job on same page
                        PrintAdjustmentJob(rs, MyPage)
                        objI.InvoiceTotal = objI.InvoiceTotal + rs("InvAdjust")
                        intPrintCount = 0

                    Else

                        'start new page for same invoice (2 jobs per page)
                        MyDocument.Pages.Add(MyPage) 'add finished page to currentdocument
                        MyPage = Nothing

                        MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F) ' new page in current document
                        MyPage.Elements.Add(New BackgroundImage(m_strImagePath))

                        objI.CurrentPage = objI.CurrentPage + 1
                        intCurrentY = PrintPageHeaders(MyPage, objC, objI)

                        intCurrentY = intCurrentY + 20

                        'print first job on page -----------------
                        PrintAdjustmentJob(rs, MyPage)
                        objI.InvoiceTotal = objI.InvoiceTotal + rs("InvAdjust")


                        intPrintCount = intPrintCount + 1


                    End If



                Else

                    'START A NEW INVOICE

                    'FIRST, PRINT TOTAL FOR LAST INVOICE -------------------------------
                    PrintAdjustmentInvoiceTotal(MyPage, objI, intCurrentY)

                    'draw previous invoice -------------------------------------------
                    MyDocument.Pages.Add(MyPage)

                    Dim strInv As String = CStr(lngCurrInvNum)

                    strDocumentName = strLastClientName & strInv
                    strFile = strDocumentName.Replace("\", " ")
                    strFile = strFile.Replace("/", "")
                    strFile = strFile.Replace(".", "")
                    strFile = strFile.Replace("-", "")

                    MyDocument.Draw(m_strPDFPath & strFile & ".pdf")

                    If blSendEmail Then
                        Dim objM As BAKEmailManager = New BAKEmailManager
                        Dim strEmailInfo As String = vbNullString

                        'From
                        strEmailInfo = "yvette@yswconsulting.com++"

                        'To - need to use objC.BillAddress
                        strEmailInfo = strEmailInfo & "yvette@yswconsulting.com++"

                        'Subject
                        strEmailInfo = strEmailInfo & "Your Invoice From Gateway Group++"

                        'Body ---------------------------------------------------------
                        strEmailInfo = strEmailInfo & "We thank you for your business.++"

                        'Mail Server
                        strEmailInfo = strEmailInfo & "mail.yswconsulting.com++"

                        'Attachment 
                        strEmailInfo = strEmailInfo & m_strPDFPath & strFile & ".pdf"

                        objM.SendEmail(strEmailInfo)

                        objM = Nothing

                    End If


                    'START THE NEW INVOICE -------------------------------------------
                    MyDocument = Nothing
                    MyDocument = CreateDocument()

                    MyPage = Nothing

                    MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F) ' new page in current document
                    MyPage.Elements.Add(New BackgroundImage(m_strImagePath))

                    'load new client object getting billing address
                    'from bill address ID in job record

                    'print headers on new page -------------------
                    objI = Nothing
                    objC = Nothing

                    objI = LoadInvoiceObject(rs)
                    objC = LoadClientObject(rs, True)
                    intCurrentY = PrintAdjustmentPageHeaders(MyPage, objC, objI, rs)

                    intCurrentY = intCurrentY + 10

                    'PRINT FIRST JOB ON NEW INVOICE
                    PrintAdjustmentJob(rs, MyPage)
                    objI.InvoiceTotal = objI.InvoiceTotal + rs("InvAdjust")
                    intPrintCount = 1


                    intCurrentY = intCurrentY + 20

                End If


                lngNewInvNum = rs("INVNUMBER")

                strLastClientName = rs("CLNAME")

                lngCurrInvNum = lngNewInvNum

            End While

            If rs.HasRows Then

                'assume we're at the end and print the last invoice
                PrintAdjustmentInvoiceTotal(MyPage, objI, intCurrentY)

                MyDocument.Pages.Add(MyPage)

                Dim strInv As String = CStr(lngCurrInvNum)

                strDocumentName = strClientName & strInv
                strFile = strDocumentName.Replace("\", " ")
                strFile = strFile.Replace("/", "")
                strFile = strFile.Replace(".", "")
                strFile = strFile.Replace("-", "")

                MyDocument.Draw(m_strPDFPath & strFile & ".pdf")

                If blSendEmail Then
                    Dim objM As BAKEmailManager = New BAKEmailManager
                    Dim strEmailInfo As String = vbNullString

                    'From
                    strEmailInfo = "yvette@yswconsulting.com++"

                    'To - need to use objC.BillAddress
                    strEmailInfo = strEmailInfo & "yvette@yswconsulting.com++"

                    'Subject
                    strEmailInfo = strEmailInfo & "Your Invoice From Gateway Group++"

                    'Body ---------------------------------------------------------
                    strEmailInfo = strEmailInfo & "We thank you for your business.++"

                    'Mail Server
                    strEmailInfo = strEmailInfo & "mail.yswconsulting.com++"

                    'Attachment 
                    strEmailInfo = strEmailInfo & m_strPDFPath & strFile & ".pdf"

                    objM.SendEmail(strEmailInfo)

                    objM = Nothing

                End If

            End If


            rs.Close()

            objDB = Nothing
            objI = Nothing


        End Function

        Private Function PrintAdjustmentPageHeaders(ByVal MyPage As ceTe.DynamicPDF.Page, ByVal objClient As ClientInfo, ByVal objInvoice As InvoiceInfo, ByVal rs As SqlDataReader) As Integer


            intCurrentY = 160

            'TOP PART OF INVOICE ---------------------------------------------------------------------------------------
            lblText = New Label("ATTN: " & objClient.Contact, intATTNCol, intCurrentY, 150, 100)
            lblText.FontSize = 10

            'Add label to MyPage
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'CLCODE --------------------------
            lblText = New Label(objClient.ClCode, intCLCodeCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'DATE -----------------------------
            lblText = New Label(objInvoice.InvoiceDate, intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'INV # ----------------------------
            lblText = New Label(objInvoice.InvoiceNumber, intInvCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'PAGE # ---------------------------
            lblText = New Label(objInvoice.CurrentPage & "/" & objInvoice.PageCount, intPageCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'second line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'company name ----------------------
            lblText = New Label(objClient.ClientName, intATTNCol, intCurrentY, 550, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'FED ID ----------------------------
            lblText = New Label("FED. ID " & Left(m_objCompany.FedID, 6) & "xxxx", intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'third line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'client address ----------------------
            Dim strAddress As String = objClient.Address1
            If Not Trim(objClient.Address2) = vbNullString Then strAddress = strAddress & ", " & objClient.Address2
            lblText = New Label(strAddress, intATTNCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'fourth line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'client city, state, zip ----------------------
            lblText = New Label(objClient.City & ", " & objClient.State & objClient.Zip, intATTNCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'employee name --------------------------------

            intCurrentY = intCurrentY + 20

            Dim strPO As String = vbNullString

            'get original job
            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim rsOrig As SqlDataReader = objDB.GetJobRecord(rs("jobnumber"))

            rsOrig.Read()

            Dim strName As String = rsOrig("FNAME")
            strName = strName & " " & rsOrig("LNAME")

            'employee name
            lblText = New Label(strName, adjCol1, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'PO Number and PRL number ----------------------------------------------------

            'Add ponumber of original job. If there is no PO in the job, get PO from client
            'record (if any) and print that

            If rsOrig.HasRows Then
                strPO = rsOrig("thispo")

                If strPO = vbNullString Then
                    'get PO from bill address table
                    If rsOrig("billaddressID") > 0 Then
                        strPO = objClient.PONumber

                    Else
                        strPO = vbNullString
                    End If
                End If

            End If



            intCurrentY = intCurrentY + 15

            lblText = New Label("PO Number: " & strPO & " | PRL number: " & rsOrig("jobnumber"), adjCol1, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            rsOrig.Close()
            objDB = Nothing

            'JOB AREA HEADERS ----------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 50

            'first line of headings --------------------------------------------------------
            lblText = New Label("Inv/CM", adjCol1, intCurrentY, 75, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Inv/CM", adjCol2, intCurrentY, 75, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("ST", adjCol3, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("ST", adjCol4, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("OT", adjCol5, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("OT", adjCol6, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("DT", adjCol7, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("DT", adjCol8, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Adjusted", adjCol10, intCurrentY, 50, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'second line of headings -------------------------------------------------------
            intCurrentY = intCurrentY + 12

            lblText = New Label("Number", adjCol1, intCurrentY, 75, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Date", adjCol2, intCurrentY, 75, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Hours", adjCol3, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Rate", adjCol4, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Hours", adjCol5, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Rate", adjCol6, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            lblText = New Label("Hours", adjCol7, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Rate", adjCol8, intCurrentY, 35, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Amount", adjCol9, intCurrentY, 50, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            lblText = New Label("Amount", adjCol10, intCurrentY, 50, 100)
            lblText.Align = TextAlign.Right
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            PrintAdjustmentPageHeaders = intCurrentY

        End Function

        Private Sub PrintAdjustmentInvoiceTotal(ByVal MyPage As ceTe.DynamicPDF.Page, ByVal objI As InvoiceInfo, ByVal intCurrentY As Integer)

            Dim strTotalLine As String

            If objI.InvoiceTotal < 0 Then
                strTotalLine = "CREDIT MEMO TOTAL"
            Else
                strTotalLine = "INVOICE TOTAL"
            End If

            intCurrentY = intCurrentY + 30

            lblText = New Label(strTotalLine, adjCol7, intCurrentY, 150, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            lblText = New Label(Format(objI.InvoiceTotal, "#,###.00"), adjCol10, intCurrentY, 50, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing



        End Sub

        Private Sub PrintAdjustmentJob(ByVal rs As SqlDataReader, ByVal MyPage As ceTe.DynamicPDF.Page)

            intCurrentY = intCurrentY + 30

            Dim strPO As String = vbNullString
            Dim dbTemp As Decimal = 0

            'get original job
            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim rsOrig As SqlDataReader = objDB.GetJobRecord(rs!jobnumber)
            objDB = Nothing

            rsOrig.Read()

            'Original Job Info -------------------------------------------------------------------------
            lblText = New Label(rsOrig!InvNumber, adjCol1, intCurrentY, 75, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label(Format(rsOrig!InvDate, "MM/dd/yyyy"), adjCol2, intCurrentY, 75, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            If Not rsOrig!billst = 0 Then
                lblText = New Label(Format(rsOrig!billst, "##0.00"), adjCol3, intCurrentY, 35, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If


            lblText = New Label(Format(rsOrig!stbillrate, "##0.00"), adjCol4, intCurrentY, 35, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            If Not rsOrig!billot = 0 Then
                lblText = New Label(Format(rsOrig!billot, "##0.00"), adjCol5, intCurrentY, 35, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If

            lblText = New Label(Format(rsOrig!otbillrate, "##0.00"), adjCol6, intCurrentY, 35, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            If Not rsOrig!billdt = 0 Then
                lblText = New Label(Format(rsOrig!billdt, "##0.00"), adjCol7, intCurrentY, 35, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If

            lblText = New Label(Format(rsOrig!dtbillrate, "##0.00"), adjCol8, intCurrentY, 35, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label(Format(rsOrig!originvamt, "##0.00"), adjCol9, intCurrentY, 50, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing



            intCurrentY = intCurrentY + 20

            'Adjustment Info -----------------------------------------------------------------------
            lblText = New Label(rs!InvNumber, adjCol1, intCurrentY, 75, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label(Format(rs!InvDate, "MM/dd/yyyy"), adjCol2, intCurrentY, 75, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            dbTemp = rs!billst
            dbTemp = dbTemp + (rsOrig!billst)

            If Not dbTemp = 0 Then
                lblText = New Label(Format(dbTemp, "##0.00"), adjCol3, intCurrentY, 35, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If

            dbTemp = rs!stbillrate
            dbTemp = dbTemp + (rsOrig!stbillrate)

            If Not dbTemp = 0 Then
                lblText = New Label(Format(dbTemp, "##0.00"), adjCol4, intCurrentY, 35, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If


            dbTemp = rs!billot
            dbTemp = dbTemp + (rsOrig!billot)

            If Not dbTemp = 0 Then
                lblText = New Label(Format(dbTemp, "##0.00"), adjCol5, intCurrentY, 35, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If

            dbTemp = rs!otbillrate
            dbTemp = dbTemp + (rsOrig!otbillrate)

            If Not dbTemp = 0 Then
                lblText = New Label(Format(dbTemp, "##0.00"), adjCol6, intCurrentY, 35, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If


            dbTemp = rs!billdt
            dbTemp = dbTemp + (rsOrig!billdt)

            If Not dbTemp = 0 Then
                lblText = New Label(Format(dbTemp, "##0.00"), adjCol7, intCurrentY, 35, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If

            dbTemp = rs!dtbillrate
            dbTemp = dbTemp + (rsOrig!dtbillrate)

            lblText = New Label(Format(dbTemp, "##0.00"), adjCol8, intCurrentY, 35, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            dbTemp = rs!invadjust
            dbTemp = dbTemp + (rsOrig!originvamt)

            lblText = New Label(Format(dbTemp, "##0.00"), adjCol9, intCurrentY, 50, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label(Format(rs!invadjust, "##0.00"), adjCol10, intCurrentY, 50, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            rsOrig.Close()

        End Sub

#End Region

#Region "Perm Invoices"

        Private Function PrintPermPageHeaders(ByVal MyPage As ceTe.DynamicPDF.Page, ByVal objClient As ClientInfo, ByVal objInvoice As InvoiceInfo) As Integer
            intCurrentY = 160

            'TOP PART OF INVOICE ---------------------------------------------------------------------------------------
            lblText = New Label("ATTN: " & objClient.Contact, intATTNCol, intCurrentY, 400, 100)
            lblText.FontSize = 10

            'Add label to MyPage
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'CLCODE --------------------------
            lblText = New Label(objClient.ClCode, intCLCodeCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'DATE -----------------------------
            lblText = New Label(objInvoice.InvoiceDate, intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'INV # ----------------------------
            lblText = New Label(objInvoice.InvoiceNumber, intInvCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'PAGE # ---------------------------
            lblText = New Label("1/1", intPageCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'second line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'company name ----------------------
            lblText = New Label(objClient.ClientName, intATTNCol, intCurrentY, 550, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'FED ID ----------------------------
            'lblText = New Label("FED. ID " & Left(m_objCompany.FedID, 6) & "xxxx", intDateCol, intCurrentY, 100, 100)
            'lblText.FontSize = 10
            'MyPage.Elements.Add(lblText)
            'lblText = Nothing

            'third line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'client address ----------------------
            Dim strAddress As String = objClient.Address1
            If Not Trim(objClient.Address2) = vbNullString Then strAddress = strAddress & ", " & objClient.Address2
            lblText = New Label(strAddress, intATTNCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'fourth line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'client city, state, zip ----------------------
            lblText = New Label(objClient.City & ", " & Trim(objClient.State) & " " & objClient.Zip, intATTNCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'Account Exec -----------------------------------
            lblText = New Label("Consultant: " & objClient.AccountExec, intDateCol, intCurrentY, 150, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing



            PrintPermPageHeaders = intCurrentY


        End Function

        Public Function PrintPermInvoices(ByVal lngStartNo As Long, ByVal lngEndNo As Long, ByVal blSendEmail As Boolean)

            Dim MyDocument As ceTe.DynamicPDF.Document = CreateDocument()

            Dim MyPage As ceTe.DynamicPDF.Page = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)

            MyPage.Elements.Add(New BackgroundImage(m_strImagePath))


            Dim lngFEDID As String = vbNullString
            Dim dbIntPercent As Double = 0
            Dim dLateDate As Date
            Dim intLateDays As Integer = 0
            Dim dbLateFee As Double = 0
            Dim lngAdjStartNum As Long = 0

            Dim lngCurrInvNum As Long = 0
            Dim lngNewInvNum As Long = 0
            Dim intPrintCount As Integer = 0
            Dim intPageCount As Integer = 0, intCurrPageCount As Integer = 0
            Dim dbThisInvoiceTotal As Decimal = 0
            Dim dbThisInvMiscBill As Decimal = 0
            Dim dbTemp As Decimal = 0
            Dim strClientName As String = vbNullString
            Dim strLastClientName As String = vbNullString

            Dim strDocumentName As String = vbNullString
            Dim strFile As String = vbNullString


            Dim objC As ClientInfo = New ClientInfo
            Dim objI As InvoiceInfo = New InvoiceInfo

            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim rs As SqlDataReader = objDB.GetPermInvoices(lngStartNo, lngEndNo)

            m_objCompany = LoadCompanyInfo()

            intPrintCount = 0



            While rs.Read

                'NEED TO HANDLE PRINTING ONE INVOICE AT A TIME -- NO CHANGE IN INVOICE NUMBER


                'record (invoice number) changes here

                lngNewInvNum = rs("INVNUMBER")
                strClientName = rs("CLNAME")

                If lngCurrInvNum = 0 Then
                    'PRINT FIRST INVOICE

                    ' headers -----------------------------------------------
                    objI = LoadInvoiceObject(rs)
                    objC = LoadClientObject(rs, False, True)
                    intCurrentY = PrintPermPageHeaders(MyPage, objC, objI)

                    intCurrentY = intCurrentY + 10

                    'first job ----------------------------------------------
                    PrintPermJob(rs, MyPage, objI, objC.ChargeInt)
                    objI.InvoiceTotal = objI.InvoiceTotal + rs("amount")
                    intPrintCount = 1


                ElseIf lngCurrInvNum = rs("INVNUMBER") Then

                    'same invoice, different job
                    If intPrintCount = 1 Then

                        'print second job on same page
                        PrintPermJob(rs, MyPage, objI, objC.ChargeInt)
                        objI.InvoiceTotal = objI.InvoiceTotal + rs("OrigInvAmt") + rs("MiscBill")
                        intPrintCount = 0

                    Else

                        'start new page for same invoice (2 jobs per page)
                        MyDocument.Pages.Add(MyPage) 'add finished page to currentdocument

                        MyPage = Nothing

                        MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F) ' new page in current document
                        MyPage.Elements.Add(New BackgroundImage(m_strImagePath))

                        objI.CurrentPage = objI.CurrentPage + 1
                        intCurrentY = PrintPermPageHeaders(MyPage, objC, objI)

                        intCurrentY = intCurrentY + 20

                        'print first job on page -----------------
                        PrintPermJob(rs, MyPage, objI, objC.ChargeInt)
                        objI.InvoiceTotal = objI.InvoiceTotal + rs("OrigInvAmt") + rs("MiscBill")


                        intPrintCount = intPrintCount + 1


                    End If



                Else

                    'START A NEW INVOICE

                    'draw previous invoice -------------------------------------------
                    MyDocument.Pages.Add(MyPage)

                    Dim strInv As String = CStr(lngCurrInvNum)

                    strDocumentName = strLastClientName & strInv
                    strFile = strDocumentName.Replace("\", " ")
                    strFile = strFile.Replace("/", "")
                    strFile = strFile.Replace(".", "")
                    strFile = strFile.Replace("-", "")

                    MyDocument.Draw(m_strPDFPath & strFile & ".pdf")

                    If blSendEmail Then
                        Dim objM As BAKEmailManager = New BAKEmailManager
                        Dim strEmailInfo As String = vbNullString

                        'From
                        strEmailInfo = "yvette@yswconsulting.com++"

                        'To - need to use objC.BillAddress
                        strEmailInfo = strEmailInfo & "yvette@yswconsulting.com++"

                        'Subject
                        strEmailInfo = strEmailInfo & "Your Invoice From Gateway Group++"

                        'Body ---------------------------------------------------------
                        strEmailInfo = strEmailInfo & "We thank you for your business.++"

                        'Mail Server
                        strEmailInfo = strEmailInfo & "mail.yswconsulting.com++"

                        'Attachment 
                        strEmailInfo = strEmailInfo & m_strPDFPath & strFile & ".pdf"

                        objM.SendEmail(strEmailInfo)

                        objM = Nothing

                    End If


                    'START THE NEW INVOICE -------------------------------------------
                    MyDocument = Nothing
                    MyDocument = CreateDocument()

                    MyPage = Nothing

                    MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F) ' new page in current document
                    MyPage.Elements.Add(New BackgroundImage(m_strImagePath))

                    'load new client object getting billing address
                    'from bill address ID in job record

                    'print headers on new page -------------------
                    objI = Nothing
                    objC = Nothing

                    objI = LoadInvoiceObject(rs)
                    objC = LoadClientObject(rs, False, True)
                    intCurrentY = PrintPermPageHeaders(MyPage, objC, objI)

                    intCurrentY = intCurrentY + 10

                    'PRINT FIRST JOB ON NEW INVOICE
                    PrintPermJob(rs, MyPage, objI, objC.ChargeInt)
                    objI.InvoiceTotal = objI.InvoiceTotal + rs("Amount")
                    intPrintCount = 1


                    intCurrentY = intCurrentY + 20

                End If


                lngNewInvNum = rs("INVNUMBER")

                strLastClientName = rs("CLNAME")

                lngCurrInvNum = lngNewInvNum

            End While

            If rs.HasRows Then

                'assume we're at the end 
                MyDocument.Pages.Add(MyPage)

                Dim strInv As String = CStr(lngCurrInvNum)

                strDocumentName = strClientName & strInv
                strFile = strDocumentName.Replace("\", " ")
                strFile = strFile.Replace("/", "")
                strFile = strFile.Replace(".", "")
                strFile = strFile.Replace("-", "")

                MyDocument.Draw(m_strPDFPath & strFile & ".pdf")

                If blSendEmail Then
                    Dim objM As BAKEmailManager = New BAKEmailManager
                    Dim strEmailInfo As String = vbNullString

                    'From
                    strEmailInfo = "yvette@yswconsulting.com++"

                    'To - need to use objC.BillAddress
                    strEmailInfo = strEmailInfo & "yvette@yswconsulting.com++"

                    'Subject
                    strEmailInfo = strEmailInfo & "Your Invoice From Gateway Group++"

                    'Body ---------------------------------------------------------
                    strEmailInfo = strEmailInfo & "We thank you for your business.++"

                    'Mail Server
                    strEmailInfo = strEmailInfo & "mail.yswconsulting.com++"

                    'Attachment 
                    strEmailInfo = strEmailInfo & m_strPDFPath & strFile & ".pdf"

                    objM.SendEmail(strEmailInfo)

                    objM = Nothing

                End If

            End If


            PrintPermAdjustmentInvoices(lngStartNo, lngEndNo, blSendEmail)

        End Function

        Private Function PrintPermJob(ByVal rs As SqlDataReader, ByVal MyPage As ceTe.DynamicPDF.Page, ByVal objI As InvoiceInfo, ByVal blChargeInt As Boolean)
            Dim dcJobTotal As Decimal = 0
            Dim intCol1 As Integer = 25
            Dim intCol2 As Integer = 150


            intCurrentY = intCurrentY + 80

            Dim strLine As String = "Recruitment consulting fee for placement of " & rs("emplname")
            If Not rs!socsec Is System.DBNull.Value Then strLine = strLine & " " & Left(rs!socsec, 7) & "-xxxx"

            If Not TypeOf (rs("permpo")) Is DBNull Then
                If Not rs("permpo") = vbNullString Then strLine = strLine & "   |   PO # " & rs("permpo")
            End If


            'Employee Name
            lblText = New Label(strLine, intATTNCol, intCurrentY, 450, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Left
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 50

            lblText = New Label("Start Date", intCol1, intCurrentY, 105, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label(rs!startdate, intCol2, intCurrentY, 75, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 20

            If rs("cramount") <> 0 Then

                lblText = New Label("Consulting Fee", intCol1, intCurrentY, 105, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label("$" & Format(rs!grossamt, "#,###.00"), intCol2, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20

                lblText = New Label("Credit Amt", intCol1, intCurrentY, 105, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label("$" & Format(rs!cramount, "#,###.00"), intCol2, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rs!creditnote, 250, intCurrentY, 300, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Left
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20


            End If

            lblText = New Label("Total Fee Due", intCol1, intCurrentY, 105, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("$" & Format(rs!amount, "#,###.00"), intCol2, intCurrentY, 75, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 20

            lblText = New Label("Description/Comments", intCol1, intCurrentY, 105, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label(rs!comment1, intCol2, intCurrentY, 400, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Left
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 20

            lblText = New Label(rs!comment2, intCol2, intCurrentY, 400, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Left
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 20

            lblText = New Label("Terms", intCol1, intCurrentY, 105, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label(rs!terms, intCol2, intCurrentY, 250, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Left
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 20

            lblText = New Label("Guarantee", intCol1, intCurrentY, 105, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label(rs!guarantee1, intCol2, intCurrentY, 250, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Left
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 50

            '--------------------------------------------------------------------------------------------------------------------------------

            'GLOBAL MESSAGE
            If Not m_objCompany.GlobalStatementMessage Is Nothing Then lblText = New Label(m_objCompany.GlobalStatementMessage, intLowerCol1, intCurrentY, 1000, 100)
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            intCurrentY = intCurrentY + 30

            lblText = New Label("Please direct any questions pertaining to this invoice to your account representative shown above.", intLowerCol1, intCurrentY, 500, 300)
            MyPage.Elements.Add(lblText)
            lblText = Nothing



            If blChargeInt Then
                If m_objCompany.InterestStartDate <> "1/1/1900" And m_objCompany.InterestStartDate <> "12:00:00 AM" And objI.InvoiceDate >= m_objCompany.InterestStartDate Then

                    Dim dLateDate As Date = DateAdd("d", objI.LateDays, rs!startdate)

                    If Not m_objCompany.DailyInterestPercent = 0 Then

                        Dim dbLateFee As Decimal = (objI.InvoiceTotal) * m_objCompany.DailyInterestPercent


                        Dim strPayLate As String = "If paid after " & dLateDate & " interest will be compounded daily on the unpaid "
                        strPayLate = strPayLate & " balance at the rate of " & _
                            CStr(m_objCompany.AnnualInterestPercent * 100) & "% annually  (" & m_objCompany.DailyInterestPercent & "% Daily)"

                        intCurrentY = intCurrentY + 150

                        'DISPLAY INTEREST MESSAGE
                        lblText = New Label(strPayLate, intLowerCol1, intCurrentY, 500, 300)
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing



                    End If
                End If
            End If






        End Function


#End Region

#Region "Perm Adjustment Invoices"

        Private Function PrintPermAdjustmentInvoices(ByVal lngStartNo As Long, ByVal lngEndNo As Long, ByVal blSendEmail As Boolean)

            Dim intCol1 As Integer = 25
            Dim intCol2 As Integer = 150
            Dim strDocumentName As String = vbNullString
            Dim strFile As String = vbNullString


            Dim objC As ClientInfo = New ClientInfo
            Dim objI As InvoiceInfo = New InvoiceInfo

            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim rs As SqlDataReader = objDB.GetPermAdjustmentInvoices(lngStartNo, lngEndNo)

            m_objCompany = LoadCompanyInfo()



            While rs.Read

                Dim MyDocument As ceTe.DynamicPDF.Document = CreateDocument()


                Dim MyPage As ceTe.DynamicPDF.Page = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)

                MyPage.Elements.Add(New BackgroundImage(m_strImagePath))

                'PRINT DETAIL

                ' headers -----------------------------------------------
                objI = LoadInvoiceObject(rs, True)
                objC = LoadClientObject(rs, False, True)
                intCurrentY = PrintPermPageHeaders(MyPage, objC, objI)

                intCurrentY = intCurrentY + 10

                'ADJUSTMENT DETAIL ---------------------------------------------------------
                intCurrentY = intCurrentY + 80

                If rs!grossadj < 0 Then
                    lblText = New Label("CREDIT MEMO", intCol1, intCurrentY, 105, 100)
                    lblText.FontSize = 10
                    lblText.Align = TextAlign.Left
                    MyPage.Elements.Add(lblText)
                    lblText = Nothing
                End If

                intCurrentY = intCurrentY + 20

                lblText = New Label("Adjustment to invoice #" & rs!originvnumber, intCol1, intCurrentY, 250, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Left
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20

                lblText = New Label("Total Adjustment", intCol1, intCurrentY, 105, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label("$" & Format(rs!grossadj, "#,###.00"), intCol2, intCurrentY, 300, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Left
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20

                lblText = New Label("Description/Comments", intCol1, intCurrentY, 105, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label(rs!comments, intCol2, intCurrentY, 300, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Left
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20

                lblText = New Label(rs!comments2, intCol2, intCurrentY, 300, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Left
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                MyDocument.Pages.Add(MyPage)

                Dim strInv As String = CStr(rs!adjinvnumber)


                strDocumentName = rs!ClName & strInv
                strFile = strDocumentName.Replace("\", " ")
                strFile = strFile.Replace("/", "")
                strFile = strFile.Replace(".", "")
                strFile = strFile.Replace("-", "")

                MyDocument.Draw(m_strPDFPath & strFile & ".pdf")

                If blSendEmail Then
                    Dim objM As BAKEmailManager = New BAKEmailManager
                    Dim strEmailInfo As String = vbNullString

                    'From
                    strEmailInfo = "yvette@yswconsulting.com++"

                    'To - need to use objC.BillAddress
                    strEmailInfo = strEmailInfo & "yvette@yswconsulting.com++"

                    'Subject
                    strEmailInfo = strEmailInfo & "Your Invoice From Gateway Group++"

                    'Body ---------------------------------------------------------
                    strEmailInfo = strEmailInfo & "We thank you for your business.++"

                    'Mail Server
                    strEmailInfo = strEmailInfo & "mail.yswconsulting.com++"

                    'Attachment 
                    strEmailInfo = strEmailInfo & m_strPDFPath & strFile & ".pdf"

                    objM.SendEmail(strEmailInfo)

                    objM = Nothing

                End If


            End While


        End Function


#End Region

#Region "Temp Statements"

        Public Sub PrintTempStatements(ByVal strClientNo As String, ByVal blPrintCurrent As Boolean)

            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim objC As ClientInfo = New ClientInfo
            Dim rsClients As SqlDataReader = objDB.GetTempStatementClients(strClientNo, True)

            Dim strFile As String = vbNullString
            Dim strDocumentName As String = vbNullString

            If Directory.Exists(m_strPDFPath) Then
            Else

                Directory.CreateDirectory(m_strPDFPath)
            End If


            While rsClients.Read
                'check for invoices
                If TempInvoicesExist(rsClients("clcode"), blPrintCurrent) Then

                    Dim MyDocument As ceTe.DynamicPDF.Document = CreateDocument()
                    Dim MyPage As ceTe.DynamicPDF.Page = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)

                    MyPage.Elements.Add(New BackgroundImage(m_strSImagePath))

                    objC = LoadStatementClient(rsClients)

                    PrintStatementHeaders(MyPage, objC)
                    PrintStatementColumnHeaders(MyPage)

                    If PrintTempStatementInvoices(objC, blPrintCurrent, MyPage, MyDocument) Then

                        PrintStatementFooter(MyPage, rsClients, objC)

                        MyDocument.Pages.Add(MyPage)

                        strDocumentName = objC.ClientName & objC.ClCode & "Statement"
                        strFile = strDocumentName.Replace("\", " ")
                        strFile = strFile.Replace("/", "")
                        strFile = strFile.Replace(".", "")
                        strFile = strFile.Replace("-", "")

                        MyDocument.Draw(m_strPDFPath & strFile & ".pdf")

                        MyPage = Nothing
                        MyDocument = Nothing

                    End If


                End If


            End While


        End Sub

        Private Function PrintStatementHeaders(ByVal MyPage As ceTe.DynamicPDF.Page, ByVal objClient As ClientInfo, Optional ByVal blIsPerm As Boolean = False) As Integer

            intCurrentY = 120

            'TOP PART OF INVOICE ---------------------------------------------------------------------------------------
            lblText = New Label("ATTN: " & objClient.Contact, intATTNCol, intCurrentY, 400, 100)
            lblText.FontSize = 10

            'Add label to MyPage
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'second line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'company name ----------------------
            lblText = New Label(objClient.ClientName, intATTNCol, intCurrentY, 550, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'third line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'DEPT ------------------------------
            If Not objClient.Dept = vbNullString Then
                lblText = New Label(objClient.Dept, intATTNCol, intCurrentY, 550, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing
                intCurrentY = intCurrentY + 12
            End If




            'client address ----------------------
            Dim strAddress As String = objClient.Address1
            If Not Trim(objClient.Address2) = vbNullString Then strAddress = strAddress & ", " & objClient.Address2
            lblText = New Label(strAddress, intATTNCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'fourth line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'client city, state, zip ----------------------
            lblText = New Label(objClient.City & ", " & objClient.State & objClient.Zip, intATTNCol, intCurrentY, 200, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 50

            lblText = New Label("STATEMENT as of " & Now.Date, 150, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 8

            lblText = New Label("----------------------------------------", 150, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 15

            lblText = New Label("CLIENT " & objClient.ClCode, 177, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            intCurrentY = intCurrentY + 20

            If Not blIsPerm Then
                If Not objClient.LateDays = 0 Then
                    lblText = New Label("TERMS ARE NET " & objClient.LateDays & " DAYS", 155, intCurrentY, 300, 100)
                    lblText.FontSize = 10
                    MyPage.Elements.Add(lblText)
                    lblText = Nothing
                End If
            End If


            intCurrentY = intCurrentY + 30

            lblText = New Label("Our records indicate the following outstanding invoices as of the above statement date", intATTNCol, intCurrentY, 600, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            PrintStatementHeaders = intCurrentY

        End Function

        Private Function PrintTempStatementInvoices(ByVal objC As ClientInfo, ByVal blPrintCurrent As Boolean, ByRef MyPage As ceTe.DynamicPDF.Page, ByRef MyDocument As ceTe.DynamicPDF.Document) As Boolean
            Dim d1630 As Date
            Dim d3160 As Date
            Dim d6190 As Date
            Dim d91 As Date

            Dim dDueDate As Date
            Dim dcInvTotal As Decimal = 0
            Dim dbTotal As Decimal, dbTempTotal As Decimal
            Dim dbCurrTotal As Double, db1630Total As Double, db3160Total As Double, db6190Total As Double, db91Total As Double
            Dim dbInterest As Decimal, dbInterestTotal As Decimal, dbUnpaidInterestTotal As Decimal
            Dim blUnpaidInvsExist As Boolean
            Dim blChargeInt As Boolean

            Dim dcAmount As Decimal = 0
            Dim intLateDays As Integer

            Dim objDB As GGDatabaseController = New GGDatabaseController

            PrintTempStatementInvoices = False

            d1630 = DateAdd("d", -16, Now)
            d3160 = DateAdd("d", -31, Now)
            d6190 = DateAdd("d", -61, Now)
            d91 = DateAdd("d", -91, Now)

            intCurrentY = intCurrentY + 20

            Dim rsInvoices As SqlDataReader = objDB.GetTempStatementInvoices(objC.ClCode, blPrintCurrent, d1630)

            While rsInvoices.Read

                'If rsInvoices("invnumber") = 43549 Then
                '    MsgBox("BREAK")
                'End If


                If intCurrentY >= m_intEndOfPage Then
                    MyDocument.Pages.Add(MyPage)
                    intCurrentY = 50
                    MyPage = Nothing
                    MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)
                    PrintStatementColumnHeaders(MyPage)
                End If

                dDueDate = DateAdd("d", rsInvoices!latedays, rsInvoices!invdate)


                If Now >= d1630 Or blPrintCurrent Then

                    If Not rsInvoices!invtotal Is System.DBNull.Value Then
                        dbTotal = rsInvoices!invtotal
                    ElseIf Not rsInvoices!adjamt Is System.DBNull.Value Then
                        dbTotal = rsInvoices!adjamt
                    End If

                    'subtract payments received
                    dcAmount = objDB.GetInvoicePayments(rsInvoices("InvNumber"))

                    dbTotal = dbTotal - dcAmount


                    'credit memos
                    dcAmount = objDB.GetInvoiceCreditMemos(rsInvoices("InvNumber"))

                    If Not dbTotal = 0 Then dbTotal = dbTotal - dcAmount

                    If Not dbTotal = 0 Then

                        intCurrentY = intCurrentY + 15


                        'disabling interest 3/14/18
                        'If rsInvoices("latedays") = 0 Then
                        '    intLateDays = objC.LateDays
                        'Else
                        '    intLateDays = rsInvoices("latedays")
                        'End If

                        'If objC.ChargeInt Then dbInterest = GetInvoiceInterest(rsInvoices!invdate, dbTotal, intLateDays)

                        ''add interest to invoice balance
                        'dbInterest = Round(dbInterest, 2)
                        'If Now > dDueDate Then dbTotal = dbTotal + dbInterest
                        'dbInterestTotal = dbInterestTotal + dbInterest


                        blUnpaidInvsExist = True

                        'print invnumber & date & amt ----------------------------------------------------
                        lblText = New Label(rsInvoices("InvNumber"), intSInvNumCol, intCurrentY, 75, 100)
                        lblText.FontSize = 10
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        lblText = New Label(rsInvoices("InvDate"), intSInvDateCol, intCurrentY, 75, 100)
                        lblText.FontSize = 10
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        If Not rsInvoices("InvTotal") Is System.DBNull.Value Then
                            dcInvTotal = rsInvoices("InvTotal")
                        Else
                            dcInvTotal = rsInvoices("AdjAmt")
                        End If


                        lblText = New Label(Format(dcInvTotal, "#,###.00"), intSOrigInvCol, intCurrentY, 75, 100)
                        lblText.FontSize = 10
                        lblText.Align = TextAlign.Right
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing


                        'current
                        If rsInvoices!invdate >= d1630 Then


                            If blPrintCurrent Then
                                dbCurrTotal = dbCurrTotal + dbTotal

                                'print dbtotal
                                lblText = New Label(Format(dbTotal, "#,###.00"), intSCurrentCol, intCurrentY, 75, 100)
                                lblText.FontSize = 10
                                lblText.Align = TextAlign.Right
                                MyPage.Elements.Add(lblText)
                                lblText = Nothing

                            End If

                            '16-30 days past due
                        ElseIf rsInvoices!invdate <= d1630 And rsInvoices!invdate > d3160 Then

                            db1630Total = db1630Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intS1630Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing


                            '31-60 days past due
                        ElseIf rsInvoices!invdate <= d3160 And rsInvoices!invdate > d6190 Then

                            db3160Total = db3160Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intS3160Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                            '61-90 days past due
                        ElseIf rsInvoices!invdate <= d6190 And rsInvoices!invdate > d91 Then

                            db6190Total = db6190Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intS6190Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                            'more than 90 days
                        ElseIf rsInvoices!invdate <= d91 Then

                            db91Total = db91Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intSOver90Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                        End If


                        'interest
                        'If dbInterest > 0 Then

                        '    'print dbinterest
                        '    lblText = New Label(Format(dbInterest, "#,###.00"), intSInterestCol, intCurrentY, 75, 100)
                        '    lblText.FontSize = 10
                        '    lblText.Align = TextAlign.Right
                        '    MyPage.Elements.Add(lblText)
                        '    lblText = Nothing

                        'End If

                        dbInterest = 0


                    End If


                    dbTempTotal = dbTempTotal + dbTotal



                End If


            End While

            'OTHER (MISC) INVOICES --------------------------------------------------------------------------------------------
            Dim rsOtherInvoices As SqlDataReader = objDB.GetTempStatementOtherInvoices(objC.ClCode, blPrintCurrent, d1630)

            intCurrentY = intCurrentY + 20

            While rsOtherInvoices.Read

                If intCurrentY >= m_intEndOfPage Then
                    MyDocument.Pages.Add(MyPage)
                    intCurrentY = 50
                    MyPage = Nothing
                    MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)
                    PrintStatementColumnHeaders(MyPage)
                End If

                dDueDate = DateAdd("d", 15, rsOtherInvoices!invdate)


                If Now >= d1630 Or blPrintCurrent Then

                    If Not rsOtherInvoices!invtotal Is System.DBNull.Value Then
                        dbTotal = rsOtherInvoices!invtotal
                    End If

                    'subtract payments received
                    dcAmount = objDB.GetInvoicePayments(rsOtherInvoices("InvNumber"))

                    dbTotal = dbTotal - dcAmount


                    'credit memos
                    dcAmount = objDB.GetInvoiceCreditMemos(rsOtherInvoices("InvNumber"))

                    If Not dbTotal = 0 Then dbTotal = dbTotal - dcAmount

                    If Not dbTotal = 0 Then

                        intCurrentY = intCurrentY + 15



                        blUnpaidInvsExist = True

                        'print invnumber & date & amt ----------------------------------------------------
                        lblText = New Label(rsOtherInvoices("InvNumber"), intSInvNumCol, intCurrentY, 75, 100)
                        lblText.FontSize = 10
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        lblText = New Label(rsOtherInvoices("InvDate"), intSInvDateCol, intCurrentY, 75, 100)
                        lblText.FontSize = 10
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        If Not rsOtherInvoices("InvTotal") Is System.DBNull.Value Then
                            dcInvTotal = rsOtherInvoices("InvTotal")
                        End If


                        lblText = New Label(Format(dcInvTotal, "#,###.00"), intSOrigInvCol, intCurrentY, 75, 100)
                        lblText.FontSize = 10
                        lblText.Align = TextAlign.Right
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing


                        'current
                        If rsOtherInvoices!invdate >= d1630 Then


                            If blPrintCurrent Then
                                dbCurrTotal = dbCurrTotal + dbTotal

                                'print dbtotal
                                lblText = New Label(Format(dbTotal, "#,###.00"), intSCurrentCol, intCurrentY, 75, 100)
                                lblText.FontSize = 10
                                lblText.Align = TextAlign.Right
                                MyPage.Elements.Add(lblText)
                                lblText = Nothing

                            End If

                            '16-30 days past due
                        ElseIf rsOtherInvoices!invdate <= d1630 And rsOtherInvoices!invdate > d3160 Then

                            db1630Total = db1630Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intS1630Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing


                            '31-60 days past due
                        ElseIf rsOtherInvoices!invdate <= d3160 And rsOtherInvoices!invdate > d6190 Then

                            db3160Total = db3160Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intS3160Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                            '61-90 days past due
                        ElseIf rsOtherInvoices!invdate <= d6190 And rsOtherInvoices!invdate > d91 Then

                            db6190Total = db6190Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intS6190Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                            'more than 90 days
                        ElseIf rsOtherInvoices!invdate <= d91 Then

                            db91Total = db91Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intSOver90Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                        End If



                    End If


                    dbTempTotal = dbTempTotal + dbTotal



                End If


            End While


            intCurrentY = intCurrentY + 20

            'unpaid interest ---------------------------------------------------------------------------------------------------
            'turning this off per Angela 3/12/18
            'If objC.ChargeInt Then
            '    Dim dCutOff As Date
            '    Dim intExpDays As Integer

            '    intExpDays = objDB.GetInterestExpirationDays * -1
            '    dCutOff = DateAdd("d", intExpDays, Now)

            '    Dim rs As SqlDataReader
            '    rs = objDB.GetStatementUnpaidInterestItems(objC.ClCode, dCutOff)

            '    While rs.Read

            '        If intCurrentY >= m_intEndOfPage Then
            '            MyDocument.Pages.Add(MyPage)
            '            intCurrentY = 50
            '            MyPage = Nothing
            '            MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)
            '            PrintStatementColumnHeaders(MyPage)
            '        End If

            '        dbInterest = rs!interest - rs!paid

            '        dbInterestTotal = dbInterestTotal + dbInterest
            '        dbUnpaidInterestTotal = dbUnpaidInterestTotal + dbInterest

            '        If dbInterest > 0 Then

            '            intCurrentY = intCurrentY + 15

            '            lblText = New Label(rs("InvNumber"), intSInvNumCol, intCurrentY, 75, 100)
            '            lblText.FontSize = 10
            '            MyPage.Elements.Add(lblText)
            '            lblText = Nothing

            '            lblText = New Label("Unpaid Interest", intSInvDateCol, intCurrentY, 75, 100)
            '            lblText.FontSize = 10
            '            MyPage.Elements.Add(lblText)
            '            lblText = Nothing

            '            lblText = New Label(Format(dbInterest, "#,###,##0.00"), intSInterestCol, intCurrentY, 75, 100)
            '            lblText.FontSize = 10
            '            lblText.Align = TextAlign.Right
            '            MyPage.Elements.Add(lblText)
            '            lblText = Nothing

            '        End If


            '    End While

            '    rs.Close()

            'End If

            If dbTempTotal = 0 Then Exit Function

            PrintTempStatementInvoices = True

            intCurrentY = intCurrentY + 20

            'print statement totals
            If blPrintCurrent Then
                If Not dbCurrTotal = 0 Then
                    lblText = New Label(Format(dbCurrTotal, "#,###.00"), intSCurrentCol, intCurrentY, 75, 100)
                    lblText.FontSize = 10
                    lblText.Align = TextAlign.Right
                    MyPage.Elements.Add(lblText)
                    lblText = Nothing
                End If
            End If

            'If dbInterestTotal > 0 Then
            '    lblText = New Label(Format(dbInterestTotal, "#,###.00"), intSInterestCol, intCurrentY, 75, 100)
            '    lblText.FontSize = 10
            '    lblText.Align = TextAlign.Right
            '    MyPage.Elements.Add(lblText)
            '    lblText = Nothing
            'End If

            If Not db1630Total = 0 Then

                lblText = New Label(Format(db1630Total, "#,###.00"), intS1630Col, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If


            If Not db3160Total = 0 Then
                lblText = New Label(Format(db3160Total, "#,###.00"), intS3160Col, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If


            If Not db6190Total = 0 Then
                lblText = New Label(Format(db6190Total, "#,###.00"), intS6190Col, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If


            If Not db91Total = 0 Then
                lblText = New Label(Format(db91Total, "#,###.00"), intSOver90Col, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If

            intCurrentY = intCurrentY + 30

            'grand total
            dbTotal = dbCurrTotal + db1630Total + db3160Total + db6190Total + db91Total + +dbUnpaidInterestTotal

            If blPrintCurrent Then
                lblText = New Label("Statement Total: " & Format(dbTotal, "#,###.00"), intATTNCol, intCurrentY, 200, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            Else
                lblText = New Label("Statement Total: " & Format(dbTotal, "#,###.00"), intATTNCol, intCurrentY, 200, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If




        End Function

        Private Sub PrintStatementFooter(ByVal myPage As ceTe.DynamicPDF.Page, ByVal rs As SqlDataReader, ByVal objC As ClientInfo)
            Dim strArgs() As String
            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim strMessages As String = objDB.GetGlobalStatementMessages
            strArgs = Split(strMessages, ",")
            Dim strPayLate As String = vbNullString
            Dim dbIntPercent As Double = 0
            Dim dbDailyInt As Double = 0
            Dim dInterestStartDate As Date

            intCurrentY = intCurrentY + 30



            'if client doesn't have custom msg, print global custom msg
            If rs("statementmsg1") Is System.DBNull.Value Then


                For x As Integer = LBound(strArgs) To UBound(strArgs)
                    lblText = New Label(strArgs(x), intATTNCol, intCurrentY, 600, 100)
                    lblText.FontSize = 10
                    myPage.Elements.Add(lblText)
                    lblText = Nothing

                    If Not Trim(strArgs(x)) = vbNullString Then intCurrentY = intCurrentY + 15

                Next

            Else

                lblText = New Label(rs("statementmsg1"), intATTNCol, intCurrentY, 600, 100)
                lblText.FontSize = 10
                myPage.Elements.Add(lblText)
                lblText = Nothing



                If Not rs("statementmsg2") Is System.DBNull.Value Then

                    intCurrentY = intCurrentY + 15
                    lblText = New Label(rs("statementmsg2"), intATTNCol, intCurrentY, 600, 100)
                    lblText.FontSize = 10
                    myPage.Elements.Add(lblText)
                    lblText = Nothing

                End If




                If rs("statementmsg3") Is System.DBNull.Value Then

                    intCurrentY = intCurrentY + 15

                    lblText = New Label(rs("statementmsg3"), intATTNCol, intCurrentY, 600, 100)
                    lblText.FontSize = 10
                    myPage.Elements.Add(lblText)
                    lblText = Nothing

                End If


            End If

            intCurrentY = intCurrentY + 20

            lblText = New Label("Gateway Group Personnel", intATTNCol, intCurrentY, 600, 100)
            lblText.FontSize = 12
            myPage.Elements.Add(lblText)
            lblText = Nothing

            'disabling 3/14/18 per Angela
            'print statement of interest charges
            'If objC.ChargeInt Then


            '    dbIntPercent = objDB.GetInterestPercent
            '    dInterestStartDate = objDB.GetInterestStartDate

            '    If Not dbIntPercent = 0 Then
            '        dbIntPercent = dbIntPercent / 100

            '        dbDailyInt = dbIntPercent * 100
            '        dbDailyInt = dbDailyInt / 365
            '        dbDailyInt = Round(dbDailyInt, 4)

            '        If dInterestStartDate <> "1/1/1900" And dInterestStartDate <> "12:00:00 AM" Then


            '            strPayLate = "Beginning " & Format(dInterestStartDate, "short date") & " interest is compounded daily on past due balances at the rate of " & _
            '        CStr(dbIntPercent * 100) & "% annually  (" & dbDailyInt & "% Daily)"

            '        End If

            '        intCurrentY = intCurrentY + 30

            '        lblText = New Label(strPayLate, intATTNCol, intCurrentY, 600, 100)
            '        lblText.FontSize = 10
            '        myPage.Elements.Add(lblText)
            '        lblText = Nothing

            '    End If

            'End If





        End Sub


        Private Function PrintStatementColumnHeaders(ByVal myPage As ceTe.DynamicPDF.Page) As Integer

            'headings for statement invoices ---------------------------------------------------------------------------
            intCurrentY = intCurrentY + 20

            lblText = New Label("Inv Num", intSInvNumCol, intCurrentY, 75, 100)
            lblText.FontSize = 10
            myPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Date", intSInvDateCol, intCurrentY, 75, 100)
            lblText.FontSize = 10
            myPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Orig. Inv", intSOrigInvCol, intCurrentY, 75, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            myPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Current", intSCurrentCol, intCurrentY, 75, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            myPage.Elements.Add(lblText)
            lblText = Nothing

            'lblText = New Label("Interest", intSInterestCol, intCurrentY, 75, 100)
            'lblText.FontSize = 10
            'lblText.Align = TextAlign.Right
            'myPage.Elements.Add(lblText)
            'lblText = Nothing

            lblText = New Label("16-30 Days", intS1630Col, intCurrentY, 75, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            myPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("31-60 Days", intS3160Col, intCurrentY, 75, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            myPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("61-90 Days", intS6190Col, intCurrentY, 75, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            myPage.Elements.Add(lblText)
            lblText = Nothing

            lblText = New Label("Over 90", intSOver90Col, intCurrentY, 75, 100)
            lblText.FontSize = 10
            lblText.Align = TextAlign.Right
            myPage.Elements.Add(lblText)
            lblText = Nothing


        End Function


        Private Function LoadStatementClient(ByVal rs As SqlDataReader) As ClientInfo

            Dim objC As ClientInfo = New ClientInfo

            objC.Address1 = rs("InvAdd1")
            objC.Address2 = rs("InvAdd2")
            objC.City = rs("InvCity")
            objC.State = rs("InvState")
            objC.Zip = rs("InvZip")
            objC.Contact = rs("invcontact")

            If Not rs("invdept") Is System.DBNull.Value Then
                objC.Dept = rs("invdept")
            Else
                objC.Dept = vbNullString
            End If


            objC.Contact = rs("invcontact")

            objC.ClientName = rs("CLNAME")


            If rs("charge_int") = "Y" Then
                objC.ChargeInt = True
            Else
                objC.ChargeInt = False
            End If

            objC.ClCode = rs!clcode

            objC.LateDays = rs!latedays

            Return objC



        End Function

        Private Function GetInvoiceInterest(ByVal dInvDate As Date, ByVal dbInvBal As Decimal, ByVal intClientLateDays As Integer, Optional ByVal strCLCode As String = "") As Decimal
            On Error GoTo errhandler

            Dim dbIntPercent As Double, dbDailyInt As Double
            Dim intDaysLate As Integer
            Dim dLateDate As Date
            Dim dbNewBal As Decimal
            Dim dbInterest As Decimal

            Dim dInterestStart As Date
            Dim objDB As GGDatabaseController = New GGDatabaseController

            dInterestStart = objDB.GetInterestStartDate


            If dInterestStart > dInvDate Then
                objDB = Nothing
                Exit Function
            End If

            dbIntPercent = objDB.GetInterestPercent

            If Not dbIntPercent = 0 Then
                dbIntPercent = dbIntPercent / 100

                dbDailyInt = dbIntPercent
                dbDailyInt = dbDailyInt / 365
                dbDailyInt = Round(dbDailyInt, 7)
            End If


            If dbDailyInt = 0 Then Exit Function

            'if late days are zero, get late days from client record
            If intClientLateDays = 0 Then Exit Function


            dLateDate = DateAdd("d", intClientLateDays, dInvDate)

            'days currently late - difference between Now and LateDate
            intDaysLate = DateDiff("d", dLateDate, Now)

            If intDaysLate <= 0 Then Exit Function

            'latecharges
            dbNewBal = dbInvBal * (1 + dbDailyInt) ^ intDaysLate

            dbInterest = dbNewBal - dbInvBal

            GetInvoiceInterest = Round(dbInterest, 2)

            Exit Function

errhandler:
            MsgBox("An error occurred while attempting to calculate interest on an invoice. " & vbCrLf & vbCrLf & _
                "The message from the system is -- " & Err.Description, vbOKOnly + vbInformation)



        End Function

        Private Function TempInvoicesExist(ByVal strClCode As String, ByVal blPrintCurrent As Boolean) As Boolean

            Dim dcTempTotal As Decimal
            Dim dcPaymentTotal As Decimal
            Dim dcBalance As Decimal
            Dim d1630 As Date
            d1630 = DateAdd("d", -16, Now)

            Dim objDB As GGDatabaseController = New GGDatabaseController


            TempInvoicesExist = False


            'CHECK FOR UNPAID TEMP INVOICES

            'unpaid invoices
            dcTempTotal = objDB.GetSumUnpaidTempInvoices(strClCode, blPrintCurrent)

            'payments received
            dcPaymentTotal = objDB.GetSumTempInvoicePayments(strClCode)

            dcBalance = dcTempTotal - dcPaymentTotal


            If dcBalance > 0 Then TempInvoicesExist = True

        End Function

        Private Function PermInvoicesExist(ByVal strClCode As String, ByVal blPrintCurrent As Boolean) As Boolean

            Dim dcPermTotal As Decimal
            Dim dcPaymentTotal As Decimal
            Dim dcBalance As Decimal
            Dim d1630 As Date
            d1630 = DateAdd("d", -16, Now)

            Dim objDB As GGDatabaseController = New GGDatabaseController


            PermInvoicesExist = False


            'CHECK FOR UNPAID PERM INVOICES

            'unpaid invoices
            dcPermTotal = objDB.GetSumUnpaidPermInvoices(strClCode, blPrintCurrent)

            'payments received
            dcPaymentTotal = objDB.GetSumPermInvoicePayments(strClCode)

            dcBalance = dcPermTotal - dcPaymentTotal


            If dcBalance > 0 Then PermInvoicesExist = True

        End Function

#End Region

#Region "Perm Statements"

        Public Sub PrintPermStatements(ByVal strClientNo As String, ByVal blPrintCurrent As Boolean)

            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim objC As ClientInfo = New ClientInfo
            Dim rsClients As SqlDataReader = objDB.GetPermStatementClients(strClientNo, True)

            Dim strFile As String = vbNullString
            Dim strDocumentName As String = vbNullString

            If Directory.Exists(m_strPDFPath) Then

            Else
                Directory.CreateDirectory(m_strPDFPath)
            End If

            'm_strImagePath = m_strAppDirectory & "\StatementBackground.jpg"

            While rsClients.Read
                'check for invoices
                If PermInvoicesExist(rsClients("clcode"), blPrintCurrent) Then

                    Dim MyDocument As ceTe.DynamicPDF.Document = CreateDocument()
                    Dim MyPage As ceTe.DynamicPDF.Page = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)

                    MyPage.Elements.Add(New BackgroundImage(m_strSImagePath))

                    objC = LoadStatementClient(rsClients)

                    PrintStatementHeaders(MyPage, objC)
                    PrintStatementColumnHeaders(MyPage)

                    If PrintPermStatementInvoices(objC, blPrintCurrent, MyPage, MyDocument) Then

                        PrintStatementFooter(MyPage, rsClients, objC)

                        MyDocument.Pages.Add(MyPage)

                        strDocumentName = objC.ClientName & objC.ClCode & "Statement"
                        strFile = strDocumentName.Replace("\", " ")
                        strFile = strFile.Replace("/", "")
                        strFile = strFile.Replace(".", "")
                        strFile = strFile.Replace("-", "")

                        MyDocument.Draw(m_strPDFPath & strFile & ".pdf")

                        MyPage = Nothing
                        MyDocument = Nothing

                    End If


                End If


            End While


        End Sub

        Private Function PrintPermStatementInvoices(ByVal objC As ClientInfo, ByVal blPrintCurrent As Boolean, ByRef MyPage As ceTe.DynamicPDF.Page, ByRef MyDocument As ceTe.DynamicPDF.Document) As Boolean
            Dim d1630 As Date
            Dim d3160 As Date
            Dim d6190 As Date
            Dim d91 As Date

            Dim dDueDate As Date
            Dim dcInvTotal As Decimal = 0
            Dim dbTotal As Decimal, dbTempTotal As Decimal, dbPermTotal As Decimal
            Dim dbCurrTotal As Double, db1630Total As Double, db3160Total As Double, db6190Total As Double, db91Total As Double
            Dim dbInterest As Decimal, dbInterestTotal As Decimal, dbUnpaidInterestTotal As Decimal
            Dim blUnpaidInvsExist As Boolean
            Dim blChargeInt As Boolean

            Dim dcAmount As Decimal = 0
            Dim intLateDays As Integer
            Dim strTemp As String

            Dim objDB As GGDatabaseController = New GGDatabaseController

            PrintPermStatementInvoices = False

            d1630 = DateAdd("d", -16, Now)
            d3160 = DateAdd("d", -31, Now)
            d6190 = DateAdd("d", -61, Now)
            d91 = DateAdd("d", -91, Now)

            intCurrentY = intCurrentY + 20

            Dim rsInvoices As SqlDataReader = objDB.GetPermStatementInvoices(objC.ClCode, blPrintCurrent, d1630)

            While rsInvoices.Read

                If intCurrentY >= m_intEndOfPage Then
                    MyDocument.Pages.Add(MyPage)
                    intCurrentY = 50
                    MyPage = Nothing
                    MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)
                    PrintStatementColumnHeaders(MyPage)
                End If

                dDueDate = DateAdd("d", rsInvoices!latedays, rsInvoices!invdate)


                If Now >= d1630 Or blPrintCurrent Then

                    If Not rsInvoices!invtotal Is System.DBNull.Value Then
                        dbTotal = rsInvoices!invtotal
                    End If

                    Dim strInv As String = rsInvoices("InvNumber")

                    'subtract payments received
                    dcAmount = objDB.GetInvoicePayments(rsInvoices("InvNumber"))

                    dbTotal = dbTotal - dcAmount


                    'credit memos
                    dcAmount = objDB.GetInvoiceCreditMemos(rsInvoices("InvNumber"))

                    If Not dbTotal = 0 Then dbTotal = dbTotal - dcAmount

                    If Not dbTotal = 0 Then

                        intCurrentY = intCurrentY + 25


                        'disabling all interest 3/14/18
                        'If rsInvoices("latedays") = 0 Then
                        '    intLateDays = objC.LateDays
                        'Else
                        '    intLateDays = rsInvoices("latedays")
                        'End If

                        'If rsInvoices!chargeint = "Y" Then dbInterest = GetInvoiceInterest(rsInvoices!startdate, dbTotal, intLateDays)

                        'add interest to invoice balance
                        'dbInterest = Round(dbInterest, 2)
                        'If Now > dDueDate Then dbTotal = dbTotal + dbInterest
                        'dbInterestTotal = dbInterestTotal + dbInterest

                        blUnpaidInvsExist = True

                        'print invnumber & date & amt ----------------------------------------------------
                        lblText = New Label(rsInvoices("InvNumber"), intSInvNumCol, intCurrentY, 75, 100)
                        lblText.FontSize = 10
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        lblText = New Label(rsInvoices("InvDate"), intSInvDateCol, intCurrentY, 75, 100)
                        lblText.FontSize = 10
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing

                        If Not rsInvoices("InvTotal") Is System.DBNull.Value Then
                            dcInvTotal = rsInvoices("InvTotal")
                        End If


                        lblText = New Label(Format(dcInvTotal, "#,###.00"), intSOrigInvCol, intCurrentY, 75, 100)
                        lblText.FontSize = 10
                        lblText.Align = TextAlign.Right
                        MyPage.Elements.Add(lblText)
                        lblText = Nothing


                        'current
                        If rsInvoices!invdate >= d1630 Then


                            If blPrintCurrent Then
                                dbCurrTotal = dbCurrTotal + dbTotal

                                'print dbtotal
                                lblText = New Label(Format(dbTotal, "#,###.00"), intSCurrentCol, intCurrentY, 75, 100)
                                lblText.FontSize = 10
                                lblText.Align = TextAlign.Right
                                MyPage.Elements.Add(lblText)
                                lblText = Nothing

                            End If

                            '16-30 days past due
                        ElseIf rsInvoices!invdate <= d1630 And rsInvoices!invdate > d3160 Then

                            db1630Total = db1630Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intS1630Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing


                            '31-60 days past due
                        ElseIf rsInvoices!invdate <= d3160 And rsInvoices!invdate > d6190 Then

                            db3160Total = db3160Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intS3160Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                            '61-90 days past due
                        ElseIf rsInvoices!invdate <= d6190 And rsInvoices!invdate > d91 Then

                            db6190Total = db6190Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intS6190Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                            'more than 90 days
                        ElseIf rsInvoices!invdate <= d91 Then

                            db91Total = db91Total + dbTotal

                            'print dbtotal
                            lblText = New Label(Format(dbTotal, "#,###.00"), intSOver90Col, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Right
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                        End If


                        'interest -- disabling 3/14/18
                        'If dbInterest > 0 Then

                        '    'print dbinterest
                        '    lblText = New Label(Format(dbInterest, "#,###.00"), intSInterestCol, intCurrentY, 75, 100)
                        '    lblText.FontSize = 10
                        '    lblText.Align = TextAlign.Right
                        '    MyPage.Elements.Add(lblText)
                        '    lblText = Nothing

                        'End If

                        dbInterest = 0

                        If Not rsInvoices!startdate = "1/1/1900" Then
                            'line 2 ----------------------------------------------------------------------------------------
                            intCurrentY = intCurrentY + 15

                            lblText = New Label("START:", intSInvNumCol, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Left
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                            lblText = New Label(rsInvoices!startdate, intSInvDateCol, intCurrentY, 75, 100)
                            lblText.FontSize = 10
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                            'disabling interest 3/14/18

                            'If Not dbInterest = 0 Then
                            '    lblText = New Label(Format(dbInterest, "#,###.00"), intSInvDateCol, intCurrentY, 75, 100)
                            '    lblText.FontSize = 10
                            '    lblText.Align = TextAlign.Right
                            '    MyPage.Elements.Add(lblText)
                            '    lblText = Nothing
                            'End If


                            'line 3 ----------------------------------------------------------------------------------------
                            intCurrentY = intCurrentY + 15

                            strTemp = rsInvoices!terms

                            lblText = New Label(strTemp, intSInvNumCol, intCurrentY, 300, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Left
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing


                            'line 4 ----------------------------------------------------------------------------------------
                            intCurrentY = intCurrentY + 15

                            strTemp = rsInvoices!emplname & " -- " & rsInvoices!comment1

                            lblText = New Label(strTemp, intSInvNumCol, intCurrentY, 600, 100)
                            lblText.FontSize = 10
                            lblText.Align = TextAlign.Left
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing

                        End If

                    End If


                    dbTempTotal = dbTempTotal + dbTotal



                End If


            End While

            intCurrentY = intCurrentY + 20

            'unpaid interest ---------------------------------------------------------------------------------------------------
            'disabling this 3/12/18 per Angela
            'If objC.ChargeInt Then
            '    Dim dCutOff As Date
            '    Dim intExpDays As Integer

            '    intExpDays = objDB.GetInterestExpirationDays * -1
            '    dCutOff = DateAdd("d", intExpDays, Now)

            '    Dim rs As SqlDataReader
            '    rs = objDB.GetStatementUnpaidInterestItems(objC.ClCode, dCutOff, True)

            '    While rs.Read

            '        If intCurrentY >= m_intEndOfPage Then
            '            MyDocument.Pages.Add(MyPage)
            '            intCurrentY = 50
            '            MyPage = Nothing
            '            MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)
            '            PrintStatementColumnHeaders(MyPage)
            '        End If

            '        dbInterest = rs!interest - rs!paid

            '        dbInterestTotal = dbInterestTotal + dbInterest
            '        dbUnpaidInterestTotal = dbUnpaidInterestTotal + dbInterest

            '        If dbInterest > 0 Then

            '            intCurrentY = intCurrentY + 15

            '            lblText = New Label(rs("InvNumber"), intSInvNumCol, intCurrentY, 75, 100)
            '            lblText.FontSize = 10
            '            MyPage.Elements.Add(lblText)
            '            lblText = Nothing

            '            lblText = New Label("Unpaid Interest", intSInvDateCol, intCurrentY, 75, 100)
            '            lblText.FontSize = 10
            '            MyPage.Elements.Add(lblText)
            '            lblText = Nothing

            '            lblText = New Label(Format(dbInterest, "#,###,##0.00"), intSInterestCol, intCurrentY, 75, 100)
            '            lblText.FontSize = 10
            '            lblText.Align = TextAlign.Right
            '            MyPage.Elements.Add(lblText)
            '            lblText = Nothing

            '        End If


            '    End While

            '    rs.Close()

            'End If

            If dbTotal + dbInterestTotal = 0 Then Exit Function

            PrintPermStatementInvoices = True

            intCurrentY = intCurrentY + 20

            'print statement totals
            If blPrintCurrent Then
                If Not dbCurrTotal = 0 Then
                    lblText = New Label(Format(dbCurrTotal, "#,###.00"), intSCurrentCol, intCurrentY, 75, 100)
                    lblText.FontSize = 10
                    lblText.Align = TextAlign.Right
                    MyPage.Elements.Add(lblText)
                    lblText = Nothing
                End If
            End If

            If dbInterestTotal > 0 Then
                lblText = New Label(Format(dbInterestTotal, "#,###.00"), intSInterestCol, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If

            If Not db1630Total = 0 Then

                lblText = New Label(Format(db1630Total, "#,###.00"), intS1630Col, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If


            If Not db3160Total = 0 Then
                lblText = New Label(Format(db3160Total, "#,###.00"), intS3160Col, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If


            If Not db6190Total = 0 Then
                lblText = New Label(Format(db6190Total, "#,###.00"), intS6190Col, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If


            If Not db91Total = 0 Then
                lblText = New Label(Format(db91Total, "#,###.00"), intSOver90Col, intCurrentY, 75, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Right
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If

            intCurrentY = intCurrentY + 30

            'grand total
            dbTotal = dbCurrTotal + db1630Total + db3160Total + db6190Total + db91Total + dbUnpaidInterestTotal

            If blPrintCurrent Then
                lblText = New Label("Statement Total: " & Format(dbTotal, "#,###.00"), intATTNCol, intCurrentY, 200, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            Else
                lblText = New Label("Statement Total: " & Format(dbTotal, "#,###.00"), intATTNCol, intCurrentY, 200, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing
            End If




        End Function



#End Region

#Region "Other Invoices"

        Public Sub PrintOtherInvoices(ByVal lngStartNo As Long, ByVal lngEndNo As Long, ByVal blSendEmail As Boolean, Optional ByVal blExportFile As Boolean = False)

            Dim MyDocument As ceTe.DynamicPDF.Document
            Dim MyPage As ceTe.DynamicPDF.Page
            Dim objI As BackOfficeNETModule.InvoiceInfo = New BackOfficeNETModule.InvoiceInfo
            Dim objC As BackOfficeNETModule.GGBackOffice.ClientInfo = New BackOfficeNETModule.GGBackOffice.ClientInfo

            m_objCompany = LoadCompanyInfo()

            Dim objDB As GGDatabaseController = New GGDatabaseController
            Dim rs As SqlDataReader = objDB.GetOtherInvoices(lngStartNo, lngEndNo)

            While rs.Read

                MyDocument = CreateDocument()

                MyPage = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F) ' new page in current document
                MyPage.Elements.Add(New BackgroundImage(m_strImagePath))

                ' headers -----------------------------------------------
                objI = LoadInvoiceObject(rs)
                objC = LoadClientObject(rs, False)

                intCurrentY = PrintOtherInvoiceHeaders(MyPage, objC, objI)

                intCurrentY = intCurrentY + 100

                lblText = New Label("Description", intLowerCol1, intCurrentY, 100, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 50


                'REASON ----------------------
                lblText = New Label("Misc. Billing:", intLowerCol1, intCurrentY, 100, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label(rs("REASON"), intLowerCol2, intCurrentY, 500, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Left
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20

                'REASON 2
                lblText = New Label(rs("REASON2"), intLowerCol2, intCurrentY, 500, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Left
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20



                'REASON 3
                lblText = New Label(rs("REASON3"), intLowerCol2, intCurrentY, 500, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Left
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20

                'Amount Due ----------------------
                lblText = New Label("Total Due:", intLowerCol1, intCurrentY, 100, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                lblText = New Label(Format(rs("AMOUNT"), "#,###.00"), intLowerCol2, intCurrentY, 100, 100)
                lblText.FontSize = 10
                lblText.Align = TextAlign.Left
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                If objC.ChargeInt Then

                    If m_objCompany.InterestStartDate <> "1/1/1900" And m_objCompany.InterestStartDate <> "12:00:00 AM" And objI.InvoiceDate >= m_objCompany.InterestStartDate Then

                        Dim dLateDate As Date = DateAdd("d", objI.LateDays, objI.InvoiceDate)

                        If Not m_objCompany.DailyInterestPercent = 0 Then

                            Dim dbLateFee As Decimal = (objI.InvoiceTotal) * m_objCompany.DailyInterestPercent


                            Dim strPayLate As String = "If paid after " & dLateDate & " interest will be compounded daily on the unpaid "
                            strPayLate = strPayLate & " balance at the rate of " & _
                                CStr(m_objCompany.AnnualInterestPercent * 100) & "% annually  (" & m_objCompany.DailyInterestPercent & "% Daily)"


                            If intCurrentY < 650 Then
                                intCurrentY = intCurrentY + 20
                            Else
                                intCurrentY = intCurrentY + 15
                            End If


                            'DISPLAY INTEREST MESSAGE
                            lblText = New Label(strPayLate, intLowerCol1, intCurrentY, 500, 300)
                            MyPage.Elements.Add(lblText)
                            lblText = Nothing


                        End If
                    End If
                End If




                MyDocument.Pages.Add(MyPage)


                SavePDF(MyDocument, objI, objC.ClientName & "Misc", objI.InvoiceNumber, blSendEmail, blExportFile, objC.ClCode)

                MyPage = Nothing
                MyDocument = Nothing


            End While

        End Sub

        Public Function PrintOtherInvoiceHeaders(ByVal MyPage As ceTe.DynamicPDF.Page, ByVal objClient As ClientInfo, ByVal objInvoice As InvoiceInfo) As Integer

            intCurrentY = 160

            'TOP PART OF INVOICE ---------------------------------------------------------------------------------------
            lblText = New Label("ATTN: " & objClient.Contact, intATTNCol, intCurrentY, 150, 100)
            lblText.FontSize = 10

            'Add label to MyPage
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'CLCODE --------------------------
            lblText = New Label(objClient.ClCode, intCLCodeCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'DATE -----------------------------
            lblText = New Label(objInvoice.InvoiceDate, intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'INV # ----------------------------
            lblText = New Label(objInvoice.InvoiceNumber, intInvCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'PAGE # ---------------------------
            lblText = New Label(objInvoice.CurrentPage & "/" & objInvoice.PageCount, intPageCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'second line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'company name ----------------------
            lblText = New Label(objClient.ClientName, intATTNCol, intCurrentY, 550, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'FED ID ----------------------------
            'lblText = New Label("FED. ID " & Left(m_objCompany.FedID, 6) & "xxxx", intDateCol, intCurrentY, 100, 100)
            'lblText.FontSize = 10
            'MyPage.Elements.Add(lblText)
            'lblText = Nothing

            intCurrentY = intCurrentY + 12


            'NET DAYS
            lblText = New Label("Net " & objInvoice.LateDays & " days", intDateCol, intCurrentY, 100, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing


            'DEPT ------------------------------
            If Not objClient.Dept = vbNullString Then
                lblText = New Label(objClient.Dept, intATTNCol, intCurrentY, 550, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing
                intCurrentY = intCurrentY + 12
            End If


            'client address ----------------------
            Dim strAddress As String = objClient.Address1
            If Not Trim(objClient.Address2) = vbNullString Then strAddress = strAddress & ", " & objClient.Address2
            lblText = New Label(strAddress, intATTNCol, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            'fourth line ------------------------------------------------------------------------------------
            intCurrentY = intCurrentY + 12

            'client city, state, zip ----------------------
            lblText = New Label(objClient.City & ", " & objClient.State & objClient.Zip, intATTNCol, intCurrentY, 300, 100)
            lblText.FontSize = 10
            MyPage.Elements.Add(lblText)
            lblText = Nothing

            PrintOtherInvoiceHeaders = intCurrentY

        End Function

#End Region

        Public Sub New()
            If m_strAppDirectory = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release" Then
                m_strImagePath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release\background.jpg"
                m_strMultiJobImagePath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release\multijobbackground.jpg"
                m_strSImagePath = m_strAppDirectory & "\StatementBackground.jpg"
                m_strTSImagePath = m_strAppDirectory & "\TSBackground.jpg"
                m_strPDFPath = "C:\PDFInvoices\"
                m_strTimeSlipTextPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release\TimeSlipText.txt"

            ElseIf m_strAppDirectory = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug" Then
                m_strImagePath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug\background.jpg"
                m_strMultiJobImagePath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Release\multijobbackground.jpg"
                m_strSImagePath = m_strAppDirectory & "\StatementBackground.jpg"
                m_strPDFPath = "C:\PDFInvoices\"
                m_strTSImagePath = m_strAppDirectory & "\TSBackground.jpg"
                m_strTimeSlipTextPath = "C:\Clients\Payroll\BackOfficeNETModule\BackOfficeNETModule\bin\Debug\TimeSlipText.txt"
            Else
                m_strImagePath = "\\" & My.Settings.ServerName & "\GGP\apps\BackOfficeNETModule\PDFManager\background.jpg"
                m_strMultiJobImagePath = "\\" & My.Settings.ServerName & "\GGP\apps\BackOfficeNETModule\PDFManager\multijobbackground.jpg"
                m_strSImagePath = "\\" & My.Settings.ServerName & "\GGP\apps\BackOfficeNETModule\PDFManager\StatementBackground.jpg"
                m_strPDFPath = "\\" & My.Settings.ServerName & "\GGP\apps\BackOfficeNETModule\PDFManager\PDFs\"
                m_strTSImagePath = "\\" & My.Settings.ServerName & "\GGP\apps\BackOfficeNETModule\PDFManager\TSBackground.jpg"
                m_strTimeSlipTextPath = "\\" & My.Settings.ServerName & "\GGP\apps\BackOfficeNETModule\PDFManager\TimeSlipText.txt"
                m_strBCHeaderPath = " \\" & My.Settings.ServerName & "\GGP\apps\BackOfficeNETModule\PDFManager\GGLogoWithAddress.jpg"
            End If
        End Sub

        Private Sub ExportPDFSToWeb(ByVal strFileName As String, ByVal strClCode As String)

            If Not Directory.Exists("\\mailscan2\Invoices\" & strClCode) Then
                Directory.CreateDirectory("\\mailscan2\Invoices\" & strClCode)
            End If

            File.Copy("\\SCHEPP\GGP\APPS\BACKOFFICENETMODULE\PDFManager\PDFS\" & strFileName, "\\mailscan2\Invoices\" & strClCode & "\" & strFileName, True)

        End Sub


#Region "Other PDFs"


        Public Sub PrintBackgroundCheck(ByVal lngEmpID As Long)

            Dim objDB As GGDatabaseController = New GGDatabaseController

            Dim rsCheck As SqlDataReader = objDB.GetBackgroundCheckInfo(lngEmpID)



            Dim strFile As String = vbNullString
            Dim strDocumentName As String = vbNullString

            If Directory.Exists(m_strPDFPath) Then

            Else
                Directory.CreateDirectory(m_strPDFPath)
            End If



            While rsCheck.Read



                Dim MyDocument As ceTe.DynamicPDF.Document = CreateDocument()
                Dim MyPage As ceTe.DynamicPDF.Page = New ceTe.DynamicPDF.Page(PageSize.Letter, PageOrientation.Portrait, 54.0F)

                MyPage.Elements.Add(New Image(m_strBCHeaderPath, 0, 0))

                Dim strContact As String = ""

                Select Case (rsCheck!contactemployer)
                    Case "Y"
                        strContact = "YES"

                    Case "N"
                        strContact = "NO"

                    Case "NA"
                        strContact = "NOT EMPLOYED"

                    Case "P"
                        strContact = "POST HIRE ONLY"

                    Case Else
                        strContact = "NO RESPONSE"

                End Select



                intCurrentY = 120

                'TOP PART 
                lblText = New Label("ACKNOWLEDGMENT AND AUTHORIZATION FOR BACKGROUND CHECK", intLowerCol1, intCurrentY, 1000, 100)
                lblText.FontSize = 14

                'Add label to MyPage
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 50

                lblText = New Label("May Employer Be Contacted:", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(strContact, intLowerCol3, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20

                lblText = New Label("Electronic Signature:", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("Signature"), intLowerCol3, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 20


                lblText = New Label("Date Signed:", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("DateSigned"), intLowerCol3, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 50


                lblText = New Label("Name:", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("FirstName") & " " & rsCheck("MiddleName") & " " & rsCheck("LastName"), intLowerCol3, intCurrentY, 500, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 20


                lblText = New Label("Address:", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("CurrentAddress") & ", " & rsCheck("City") & " , " & rsCheck("state") & " " & rsCheck("zipcode"), intLowerCol3, intCurrentY, 500, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20

                lblText = New Label("County: ", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("county"), intLowerCol3, intCurrentY, 500, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing

                intCurrentY = intCurrentY + 20

                lblText = New Label("Length at Address: ", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("lengthataddress"), intLowerCol3, intCurrentY, 500, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 20


                intCurrentY = intCurrentY + 12

                lblText = New Label("DL #: ", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("DriversLic") & "  State Issued: " & rsCheck("stateissued"), intLowerCol3, intCurrentY, 500, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 20


                lblText = New Label("Date of Birth: ", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("DOB"), intLowerCol3, intCurrentY, 500, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 20



                lblText = New Label("SSN: ", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("SSN"), intLowerCol3, intCurrentY, 500, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 20

                lblText = New Label("Previous Addresses: ", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("previousaddress"), intLowerCol3, intCurrentY, 400, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 20


                lblText = New Label("Other Names: ", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("othernames"), intLowerCol3, intCurrentY, 400, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 20

                lblText = New Label("Education: ", intLowerCol1, intCurrentY, 150, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                lblText = New Label(rsCheck("education"), intLowerCol3, intCurrentY, 400, 100)
                lblText.FontSize = 10
                MyPage.Elements.Add(lblText)
                lblText = Nothing


                intCurrentY = intCurrentY + 20





                MyDocument.Pages.Add(MyPage)

                strDocumentName = "BackgroundCheck_" & lngEmpID.ToString

                strFile = strDocumentName.Replace("\", " ")

                strFile = strFile.Replace("/", "")
                strFile = strFile.Replace(".", "")

                strFile = strFile.Replace("-", "")

                MyDocument.Draw(m_strPDFPath & strFile & ".pdf")


                MyPage = Nothing
                MyDocument = Nothing






            End While

            rsCheck.Close()

        End Sub



#End Region


#Region "Handle Null"
        Public Function RidNull(ByVal n As DBNull)
            Return 0
        End Function
        Public Function RidNull(ByVal int As Integer)
            If int.ToString Is System.DBNull.Value Then
                Return 0
            Else
                Return int
            End If
        End Function

        Public Function RidNull(ByVal lng As Long)
            If lng.ToString Is System.DBNull.Value Then
                Return 0
            Else
                Return lng
            End If
        End Function

        Public Function RidNull(ByVal str As String)
            If str.ToString Is System.DBNull.Value Then
                Return vbNullString
            Else
                Return str
            End If
        End Function

        Public Function RidNull(ByVal dec As Decimal)
            If dec.ToString Is System.DBNull.Value Then
                Return 0
            Else
                Return dec
            End If
        End Function

#End Region


    End Class



End Namespace

