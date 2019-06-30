Public Class InvoiceInfo
#Region "Variable Declarations"


    Private m_strAccountExec As String = vbNullString
    Private m_intPageCount As Integer = 0
    Private m_intCurrentPage As Integer = 0
    Private m_blPrintCopy As Boolean = False
    Private m_dcInvoiceTotal As Decimal = 0
    Private m_dcPageSubTotal As Decimal = 0
    Private m_dcMiscBillTotal As Decimal = 0
    Private m_intLateDays As Integer = 0

    Private m_lngInvoiceNumber As Long = 0
    Private m_dInvoiceDate As Date

#End Region

#Region "Property Declarations"

    Public Property PageSubTotal() As Decimal
        Get
            Return m_dcPageSubTotal
        End Get
        Set(ByVal value As Decimal)
            m_dcPageSubTotal = value
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

    Public Property InvoiceDate() As Date
        Get
            Return m_dInvoiceDate
        End Get
        Set(ByVal value As Date)
            m_dInvoiceDate = value
        End Set
    End Property

    Public Property InvoiceNumber() As Long
        Get
            Return m_lngInvoiceNumber
        End Get
        Set(ByVal value As Long)
            m_lngInvoiceNumber = value
        End Set
    End Property

    Public Property AccountExec() As String
        Get
            Return m_strAccountExec
        End Get
        Set(ByVal value As String)
            m_strAccountExec = value
        End Set
    End Property


    Public Property PageCount() As Integer
        Get
            Return m_intPageCount
        End Get
        Set(ByVal value As Integer)
            m_intPageCount = value
        End Set
    End Property


    Public Property CurrentPage() As Integer
        Get
            Return m_intCurrentPage
        End Get
        Set(ByVal value As Integer)
            m_intCurrentPage = value
        End Set
    End Property


    Public Property PrintCopy() As Boolean
        Get
            Return m_blPrintCopy
        End Get
        Set(ByVal value As Boolean)
            m_blPrintCopy = value
        End Set
    End Property


    Public Property InvoiceTotal() As Decimal
        Get
            Return m_dcInvoiceTotal
        End Get
        Set(ByVal value As Decimal)
            m_dcInvoiceTotal = value
        End Set
    End Property

    Public Property MiscBillTotal() As Decimal
        Get
            Return m_dcMiscBillTotal
        End Get
        Set(ByVal value As Decimal)
            m_dcMiscBillTotal = value
        End Set
    End Property


#End Region

End Class
