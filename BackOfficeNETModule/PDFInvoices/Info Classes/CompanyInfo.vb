Public Class CompanyInfo

#Region "Variable Declarations"
    Private m_strGlobalStatementMessage As String = vbNullString
    Private m_dcInterestPercent As Decimal = 0
    Private m_dcAnnualInterestPercent As Decimal = 0
    Private m_strFedID As String = vbNullString
    Private m_dInterestStartDate As Date
#End Region

#Region "Property Declarations"

    Public Property AnnualInterestPercent()
        Get
            AnnualInterestPercent = m_dcAnnualInterestPercent
        End Get
        Set(ByVal value)
            m_dcAnnualInterestPercent = value
        End Set
    End Property


    Public Property DailyInterestPercent() As Decimal
        Get
            Return m_dcInterestPercent
        End Get
        Set(ByVal value As Decimal)
            m_dcInterestPercent = value
        End Set
    End Property



    Public Property InterestStartDate() As Date
        Get
            Return m_dInterestStartDate
        End Get
        Set(ByVal value As Date)
            m_dInterestStartDate = value
        End Set
    End Property


    Public Property GlobalStatementMessage() As String
        Get
            Return m_strGlobalStatementMessage
        End Get
        Set(ByVal value As String)
            m_strGlobalStatementMessage = value
        End Set
    End Property


    Public Property FedID() As String
        Get
            Return m_strFedID
        End Get
        Set(ByVal value As String)
            m_strFedID = value
        End Set
    End Property


#End Region

End Class
