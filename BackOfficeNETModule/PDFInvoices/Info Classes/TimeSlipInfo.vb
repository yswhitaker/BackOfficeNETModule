Imports System.Data.SqlClient
Imports System.Collections.Specialized
Imports System.Collections.Generic


Namespace GGBackOffice

    Public Class TimeSlipInfo
        ' local property declarations
        Private _ModuleId As Integer
        Private _ItemId As Integer

        Dim m_lngBackOfficeID As Long = 0
        Dim m_lngEmployeeID As Long = 0
        Dim m_strReportTo As String = vbNullString
        Dim m_strEmployeeName As String = vbNullString
        Dim m_strDept As String = vbNullString
        Dim m_strJobTitle As String = vbNullString
        Dim m_blReturning As Boolean = False
        Dim m_dNextAvailable As Date
        Dim m_strClientName As String = vbNullString
        Dim m_strClientNumber As String = vbNullString
        Dim m_strTSManager As String = vbNullString
        Dim m_dWEDate As Date

        Dim m_strApprovedBy As String = vbNullString
        Dim m_dApprovedOn As Date

        Dim m_ldDays As ListDictionary

        Dim hshDays As Hashtable = CollectionsUtil.CreateCaseInsensitiveHashtable

        Dim quDays As Queue(Of GGTimeSlipDay)


        ' initialization
        Public Sub New()
            m_ldDays = New ListDictionary
        End Sub
        Public Property ApprovedOn() As Date
            Get
                ApprovedOn = m_dApprovedOn
            End Get
            Set(ByVal value As Date)
                m_dApprovedOn = value
            End Set
        End Property

        Public Property ApprovedBy() As String
            Get
                ApprovedBy = m_strApprovedBy
            End Get
            Set(ByVal value As String)
                m_strApprovedBy = value
            End Set
        End Property

        Public Property EmployeeName() As String
            Get
                EmployeeName = Me.m_strEmployeeName
            End Get
            Set(ByVal value As String)
                Me.m_strEmployeeName = value
            End Set
        End Property


        Public Property WEDate() As Date
            Get
                WEDate = Me.m_dWEDate
            End Get
            Set(ByVal value As Date)
                Me.m_dWEDate = value
            End Set
        End Property

        Public Property TSManager() As String
            Get
                TSManager = Me.m_strTSManager
            End Get
            Set(ByVal value As String)
                Me.m_strTSManager = value
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

        Public Property ClientNumber() As String
            Get
                ClientNumber = m_strClientNumber
            End Get
            Set(ByVal value As String)
                m_strClientNumber = value
            End Set
        End Property


        Public Property NextAvailable() As Date
            Get
                NextAvailable = m_dNextAvailable
            End Get
            Set(ByVal value As Date)
                m_dNextAvailable = value
            End Set
        End Property

        Public Property Returning() As Boolean
            Get
                Returning = m_blReturning
            End Get
            Set(ByVal value As Boolean)
                m_blReturning = value
            End Set
        End Property

        Public Property JobTitle() As String
            Get
                JobTitle = m_strJobTitle
            End Get
            Set(ByVal value As String)
                m_strJobTitle = value
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

        Public Property ReportTo() As String
            Get
                ReportTo = m_strReportTo
            End Get
            Set(ByVal value As String)
                m_strReportTo = value
            End Set
        End Property

        Public Property EmployeeID() As Long
            Get
                EmployeeID = m_lngEmployeeID
            End Get
            Set(ByVal value As Long)
                m_lngEmployeeID = value
            End Set
        End Property

        Public Property BackOfficeID() As Long
            Get
                BackOfficeID = m_lngBackOfficeID
            End Get
            Set(ByVal value As Long)
                m_lngBackOfficeID = value
            End Set
        End Property

        Public Property Days() As ListDictionary
            Get
                Days = m_ldDays
            End Get
            Set(ByVal value As ListDictionary)
                m_ldDays = value
            End Set
        End Property

        Public Property HashDays() As Hashtable
            Get
                HashDays = Me.hshDays
            End Get
            Set(ByVal value As Hashtable)
                Me.hshDays = value
            End Set
        End Property

        Public Property QDays() As Queue(Of GGTimeSlipDay)
            Get
                QDays = Me.quDays
            End Get
            Set(ByVal value As Queue(Of GGBackOffice.TimeSlipInfo.GGTimeSlipDay))
                Me.quDays = value
            End Set
        End Property

        Public Property ModuleId() As Integer
            Get
                Return _ModuleId
            End Get
            Set(ByVal Value As Integer)
                _ModuleId = Value
            End Set
        End Property

        Public Property ItemId() As Integer
            Get
                Return _ItemId
            End Get
            Set(ByVal Value As Integer)
                _ItemId = Value
            End Set
        End Property

        Class GGTimeSlipDay

            Dim m_strDay As String
            Dim m_dTimeIn As String
            Dim m_dTimeOut As String
            Dim m_strLunch As String
            Dim m_dcTotalHours As Decimal


            Public Property TotalHours() As Decimal
                Get
                    TotalHours = m_dcTotalHours
                End Get
                Set(ByVal value As Decimal)
                    m_dcTotalHours = value
                End Set
            End Property

            Public Property Lunch() As String
                Get
                    Lunch = m_strLunch
                End Get
                Set(ByVal value As String)
                    m_strLunch = value
                End Set
            End Property

            Public Property TimeOut() As String
                Get
                    TimeOut = m_dTimeOut
                End Get
                Set(ByVal value As String)
                    m_dTimeOut = value
                End Set
            End Property

            Public Property TimeIn() As String
                Get
                    TimeIn = m_dTimeIn
                End Get
                Set(ByVal value As String)
                    m_dTimeIn = value
                End Set
            End Property

            Public Property Day() As String
                Get
                    Day = m_strDay
                End Get
                Set(ByVal value As String)
                    m_strDay = value
                End Set
            End Property


        End Class


        Public Function GetSavedTimeSlip(ByVal lngAssignmentID As Long) As GGBackOffice.TimeSlipInfo

            Dim objT As TimeSlipInfo
            Dim objDay As TimeSlipInfo.GGTimeSlipDay = New TimeSlipInfo.GGTimeSlipDay
            Dim ldDays As ListDictionary = New ListDictionary
            Dim qDays As New Queue(Of GGBackOffice.TimeSlipInfo.GGTimeSlipDay)
            Dim HSH As Hashtable = CollectionsUtil.CreateCaseInsensitiveHashtable

            Dim cmd2 As SqlCommand = New SqlCommand
            Dim cmdDays As SqlCommand = New SqlCommand
            Dim reader As SqlDataReader
            Dim rdDays As SqlDataReader

            Dim param2 As SqlParameter = New SqlParameter

            Dim lngTimeSlipID As Long = 0

            Dim objDB As GGBackOffice.GGDatabaseController = New GGDatabaseController


            Try
                reader = objDB.GetTimeSlipByAssignmentID(lngAssignmentID)

                If reader Is Nothing Then Exit Function



                While reader.Read

                    objT = New TimeSlipInfo

                    lngTimeSlipID = reader("pk_id")

                    With objT
                        If Not reader("EmployeeName") Is System.DBNull.Value Then .EmployeeName = reader("EmployeeName")
                        .BackOfficeID = reader("AssignmentID")
                        If Not reader("Dept") Is System.DBNull.Value Then .Dept = reader("Dept")
                        .EmployeeID = reader("EmployeeID")
                        If Not reader("JobTitle") Is System.DBNull.Value Then .JobTitle = reader("JobTitle")
                        .Returning = reader("Returning")
                        If Not reader("ReportTo") Is System.DBNull.Value Then .ReportTo = reader("ReportTo")
                        .WEDate = reader("wedate")
                        If Not reader("clname") Is System.DBNull.Value Then .ClientName = reader("clname")
                        If Not reader("clcode") Is System.DBNull.Value Then .ClientNumber = reader("clcode")
                        If Not reader("ReportTo") Is System.DBNull.Value Then .TSManager = reader("ReportTo")
                        If Not reader("ApprovedBy") Is System.DBNull.Value Then .ApprovedBy = reader("ApprovedBy")
                        If Not reader("ApprovedOn") Is System.DBNull.Value Then .ApprovedOn = reader("ApprovedOn")
                    End With


                End While

                reader.Close()

                'get 


                rdDays = objDB.GetTimeSlipHours(lngTimeSlipID)

                If rdDays Is Nothing Then Exit Function

                'loop through data reader, loading a new day class and adding to listdictionary
                While rdDays.Read

                    objDay = New GGBackOffice.TimeSlipInfo.GGTimeSlipDay

                    With objDay
                        .Day = rdDays("Day")
                        .Lunch = rdDays("Lunch")
                        .TimeIn = rdDays("TimeIn")
                        .TimeOut = rdDays("TimeOut")
                        .TotalHours = rdDays("TotalHours")
                    End With

                    qDays.Enqueue(objDay)
                    ldDays.Add(rdDays("Day"), objDay)
                    HSH.Add(rdDays("Day"), objDay)

                    objDay = Nothing


                End While


                If Not objT Is Nothing Then
                    objT.Days = ldDays
                    objT.HashDays = HSH
                    objT.qDays = qDays
                End If


                reader.Close()


            Catch


            Finally


            End Try




            Return objT


        End Function


    End Class





End Namespace