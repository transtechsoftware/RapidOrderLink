Imports System.Data.SqlClient


Public Class Location

    Public TRCDBName As String = "UN_TRACKING"
    Public TRCDBUser As String = "Unison" '"tpctrk"
    Public TRCDBPass As String = "unison" '"top"
    Public TRCTblPath As String = TRCDBName & ".dbo."


    ' Members used to query the Unison Database
    Private m_oDataAdapter As SqlDataAdapter
    Private m_oDataSet As DataSet
    Private m_strSelect As String

    ' Members used for reporting the objects status
    Private m_bIsEmpty As Boolean
    Public ReadOnly Property IsEmpty() As Boolean
        Get
            Return m_bIsEmpty
        End Get
    End Property

    ' Memebers that map to corresponding data fields
    Private m_strCustomerID As New VarChar(10)
    Public Property CustomerID() As String
        Get
            Return m_strCustomerID.Value
        End Get
        Set(ByVal Value As String)
            m_strCustomerID.Value = Value
        End Set
    End Property

    Private m_strLocationID As New VarChar(10)
    Public Property LocationID() As String
        Get
            Return m_strLocationID.Value
        End Get
        Set(ByVal Value As String)
            m_strLocationID.Value = Value
        End Set
    End Property

    Private m_strNAME As New VarChar(70)
    Public Property Name() As String
        Get
            Return m_strNAME.Value
        End Get
        Set(ByVal Value As String)
            m_strNAME.Value = Value
        End Set
    End Property

    Private m_strAddress1 As New VarChar(50)
    Public Property Address1() As String
        Get
            Return m_strAddress1.Value
        End Get
        Set(ByVal Value As String)
            m_strAddress1.Value = Value
        End Set
    End Property

    Private m_strAddress2 As New VarChar(30)
    Public Property Address2() As String
        Get
            Return m_strAddress2.Value
        End Get
        Set(ByVal Value As String)
            m_strAddress2.Value = Value
        End Set
    End Property

    Private m_strCity As New VarChar(50)
    Public Property City() As String
        Get
            Return m_strCity.Value
        End Get
        Set(ByVal Value As String)
            m_strCity.Value = Value
        End Set
    End Property

    Private m_strState As New VarChar(2)
    Public Property State() As String
        Get
            Return m_strState.Value
        End Get
        Set(ByVal Value As String)
            m_strState.Value = Value
        End Set
    End Property

    Private m_strZip As New VarChar(16)
    Public Property Zip() As String
        Get
            Return m_strZip.Value
        End Get
        Set(ByVal Value As String)
            m_strZip.Value = Value
        End Set
    End Property

    Private m_strCONTACT As New VarChar(40)
    Public Property Contact() As String
        Get
            Return m_strCONTACT.Value
        End Get
        Set(ByVal Value As String)
            m_strCONTACT.Value = Value
        End Set
    End Property

    Private m_strPHONE As New VarChar(20)
    Public Property Phone() As String
        Get
            Return m_strPHONE.Value
        End Get
        Set(ByVal Value As String)
            m_strPHONE.Value = Value
        End Set
    End Property

    Private m_strACTIVE As New VarChar(1)
    Public Property Active() As Boolean
        Get
            If m_strACTIVE.Value.Equals("1") Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal Value As Boolean)
            If Value = True Then
                m_strACTIVE.Value = "1"
            Else
                m_strACTIVE.Value = "0"
            End If
        End Set
    End Property

    Private m_strEMAIL As New VarChar(60)
    Public Property EMail() As String
        Get
            Return m_strEMAIL.Value
        End Get
        Set(ByVal Value As String)
            m_strEMAIL.Value = Value
        End Set

    End Property

    Private m_strAddressID As Integer
    Public Property AddressId() As Integer
        Get
            Return m_strAddressID
        End Get
        Set(ByVal Value As Integer)
            m_strAddressID = Value
        End Set
    End Property

    Private m_strPassword As New VarChar(10)
    Public Property Password() As String
        Get
            Return m_strPassword.Value
        End Get
        Set(ByVal Value As String)
            m_strPassword.Value = Value
        End Set
    End Property

    Public Function GetIfUniqueLocId(ByVal p_strLocId As String) As Boolean
        '   This function will look in the UN_TRACKING database to see if its LocationID is unique system-wide
        '   If it is, it will populate its members with the corresponding data otherwise it will return false

        Dim bRetVal As Boolean

        Try

            m_strSelect = "SELECT * FROM " & TRCTblPath & "Location WHERE LocationID = '" & p_strLocId & "'"
            PopulateDataset2(m_oDataAdapter, m_oDataSet, m_strSelect)

            Dim iRowCount As Integer = m_oDataSet.Tables(0).Rows.Count

            If iRowCount <> 1 Then
                bRetVal = False
            Else
                ' Transfer data from DataSet object to this object
                Dim oDataRow As DataRow = m_oDataSet.Tables(0).Rows(0)

                CustomerID = oDataRow("CustomerID")
                LocationID = oDataRow("LocationID")
                Name = oDataRow("NAME")
                Address1 = oDataRow("Address1")
                Address2 = oDataRow("Address2")
                City = oDataRow("City")
                State = oDataRow("State")
                Zip = oDataRow("Zip")
                Contact = oDataRow("CONTACT")
                Phone = oDataRow("PHONE")
                m_strACTIVE.Value = oDataRow("ACTIVE")
                EMail = oDataRow("EMAIL")
                AddressId = oDataRow("AddressID")
                Password = oDataRow("Password")

                bRetVal = True

            End If


        Catch ex As Exception

            'Message modified by Michael Pastor
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Critical Error")
            '- MessageBox.Show(ex.Message, "DEBUG")
            bRetVal = False

        End Try

        CleanUp()
        m_bIsEmpty = Not bRetVal
        Return bRetVal

    End Function

    Private Sub CleanUp()

        Try

            m_oDataAdapter = Nothing
            m_oDataSet = Nothing
            m_strSelect = ""

        Catch ex As Exception

            Return

        End Try

    End Sub

    Sub New()
        m_bIsEmpty = True
    End Sub

End Class
