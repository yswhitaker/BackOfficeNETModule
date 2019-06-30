Public Class JobInfo
    Private m_dDate As Date
    Private m_lngInvNumber As String

    Private m_dWEDate As Date
    Private m_strEmployeeName As String = vbNullString

    Private m_strSTHours As String
    Private m_strSTRate As String
    Private m_strOTHours As String
    Private m_strOTRate As String
    Private m_strDTHours As String
    Private m_strDTRate As String

    Private m_strInvoiceTotal As String
    Private m_lngJobNumber As String
    Private m_strPosition As String
    Private m_strAssignedTo As String
    Private m_strNetDays As String

    Private m_strPONumber As String

    Public Property PONumber() As String
        Get
            PONumber = m_strPONumber
        End Get
        Set(ByVal value As String)
            m_strPONumber = value
        End Set
    End Property

    Public Property EmployeeName() As String
        Get
            EmployeeName = m_strEmployeeName
        End Get
        Set(ByVal value As String)
            m_strEmployeeName = value
        End Set
    End Property

    Public Property NetDays() As String
        Get
            NetDays = m_strNetDays
        End Get
        Set(ByVal value As String)
            m_strNetDays = value
        End Set
    End Property
    Public Property AssignedTo() As String
        Get
            AssignedTo = m_strAssignedTo
        End Get
        Set(ByVal value As String)
            m_strAssignedTo = value
        End Set
    End Property
    Public Property Position() As String
        Get
            Position = m_strPosition
        End Get
        Set(ByVal value As String)
            m_strPosition = value
        End Set
    End Property
    Public Property JobNumber() As String
        Get
            JobNumber = m_lngJobNumber
        End Get
        Set(ByVal value As String)
            m_lngJobNumber = value
        End Set
    End Property
    Public Property InvoiceTotal() As String
        Get
            InvoiceTotal = m_strInvoiceTotal
        End Get
        Set(ByVal value As String)
            m_strInvoiceTotal = value
        End Set
    End Property
    Public Property DTRate() As String
        Get
            DTRate = m_strDTRate
        End Get
        Set(ByVal value As String)
            m_strDTRate = value
        End Set
    End Property

    Public Property DTHours() As String
        Get
            DTHours = m_strDTHours
        End Get
        Set(ByVal value As String)
            m_strDTHours = value
        End Set
    End Property

    Public Property OTRate() As String
        Get
            OTRate = m_strOTRate
        End Get
        Set(ByVal value As String)
            m_strOTRate = value
        End Set
    End Property
    Public Property OTHours() As String
        Get
            OTHours = m_strOTHours
        End Get
        Set(ByVal value As String)
            m_strOTHours = value
        End Set
    End Property
    Public Property STRate() As String
        Get
            STRate = m_strSTRate
        End Get
        Set(ByVal value As String)
            m_strSTRate = value
        End Set
    End Property
    Public Property STHours() As String
        Get
            STHours = m_strSTHours
        End Get
        Set(ByVal value As String)
            m_strSTHours = value
        End Set
    End Property

    Public Property WEDate() As Date
        Get
            WEDate = m_dWEDate
        End Get
        Set(ByVal value As Date)
            m_dWEDate = value
        End Set
    End Property


    Public Property InvNumber() As String
        Get
            InvNumber = m_lngInvNumber
        End Get
        Set(ByVal value As String)
            m_lngInvNumber = value
        End Set
    End Property

    Public Property InvoiceDate() As Date
        Get
            InvoiceDate = m_dDate
        End Get
        Set(ByVal value As Date)
            m_dDate = value
        End Set
    End Property

End Class


