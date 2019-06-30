Namespace GGBackOffice


    Public Class ClientInfo
        Private m_strClientName As String = vbNullString
        Private m_strClCode As String = vbNullString
        Private m_strContact As String = vbNullString
        Private m_strAddress1 As String = vbNullString
        Private m_strAddress2 As String = vbNullString
        Private m_strCity As String = vbNullString
        Private m_strState As String = vbNullString
        Private m_strZip As String = vbNullString
        Private m_strDept As String = vbNullString
        Private m_strBillEmail As String = vbNullString


        Private m_blChargeInt As Boolean = False
        Private m_intLateDays As Integer = 0

        Private m_strInterestMessage As String = vbNullString
        Private m_strGlobalMessage As String = vbNullString
        Private m_strAccountExec As String = vbNullString
        Private m_strPONumber As String = vbNullString

        Private m_intJobCount As Integer

        Private m_lstJobs As New List(Of JobInfo)



        Public Property BillEmail() As String
            Get
                BillEmail = m_strBillEmail
            End Get
            Set(ByVal value As String)
                m_strBillEmail = value
            End Set
        End Property

        Public Property LateDays() As Integer
            Get
                LateDays = m_intLateDays
            End Get
            Set(ByVal value As Integer)
                m_intLateDays = value
            End Set
        End Property

        Public Property PONumber() As String
            Get
                PONumber = m_strPONumber
            End Get
            Set(ByVal value As String)
                m_strPONumber = value
            End Set
        End Property

        Public Property AccountExec() As String
            Get
                AccountExec = m_strAccountExec
            End Get
            Set(ByVal value As String)
                m_strAccountExec = value
            End Set
        End Property

        Public Property GlobalMessage() As String
            Get
                GlobalMessage = m_strGlobalMessage
            End Get
            Set(ByVal value As String)
                m_strGlobalMessage = value
            End Set
        End Property

        Public Property InterestMessage() As String
            Get
                InterestMessage = m_strInterestMessage
            End Get
            Set(ByVal value As String)
                m_strInterestMessage = value
            End Set
        End Property

        Public Property Dept() As String
            Get
                Dept = m_strDept
            End Get
            Set(ByVal value As String)
                m_strDept = value
            End Set
        End Property

        Public Property ChargeInt() As Boolean
            Get
                ChargeInt = m_blChargeInt
            End Get
            Set(ByVal value As Boolean)
                m_blChargeInt = value
            End Set
        End Property

        Public Property InvoiceCount() As Integer
            Get
                InvoiceCount = m_lstJobs.Count
            End Get
            Set(ByVal value As Integer)
                m_intJobCount = value
            End Set
        End Property

        Public Property Zip() As String
            Get
                Zip = m_strZip
            End Get
            Set(ByVal value As String)
                m_strZip = value
            End Set
        End Property
        Public Property State() As String
            Get
                State = m_strState
            End Get
            Set(ByVal value As String)
                m_strState = value
            End Set
        End Property
        Public Property City() As String
            Get
                City = m_strCity
            End Get
            Set(ByVal value As String)
                m_strCity = value
            End Set
        End Property
        Public Property Address2() As String
            Get
                Address2 = m_strAddress2
            End Get
            Set(ByVal value As String)
                m_strAddress2 = value
            End Set
        End Property
        Public Property Address1() As String
            Get
                Address1 = m_strAddress1
            End Get
            Set(ByVal value As String)
                m_strAddress1 = value
            End Set
        End Property
        Public Property Contact() As String
            Get
                Contact = m_strContact
            End Get
            Set(ByVal value As String)
                m_strContact = value
            End Set
        End Property

        Public Property ClCode() As String
            Get
                ClCode = m_strClCode
            End Get
            Set(ByVal value As String)
                m_strClCode = value
            End Set
        End Property

        Public Property ClientName() As String
            Get
                ClientName = m_strClientName
            End Get
            Set(ByVal value As String)
                m_strClientName = value
            End Set
        End Property

        Public Property JobsList() As List(Of JobInfo)
            Get
                JobsList = m_lstJobs
            End Get
            Set(ByVal value As List(Of JobInfo))
                m_lstJobs = value
            End Set
        End Property


        Public Sub AddInvoice(ByVal objJob As JobInfo)
            m_lstJobs.Add(objJob)
        End Sub

    End Class


End Namespace

