Imports TTSI.BARCODES
Imports System.Data.SqlClient

Public Class TrackingEvent
    ' This class maps to a single Row of the UN_TRACKING.dbo.EVENT table, FOR NOW

    ' Members used to query the Unison Database
    Private m_oDataAdapter As SqlDataAdapter
    Private m_oDataSet As DataSet
    Private m_strSelectStmt As String
    Private m_strInsertStmt As String
    Private ReadOnly Property InsertStatement() As String
        Get
            m_strInsertStmt = "INSERT INTO " & TRCTblPath & "EVENT(EventCode, ScanDate, OperatorID, PointID, TicketNum, TrackingNum, ThirdPartyBarcode, ContainerBarcode, DeliveryOption, DeliveryComments, ToCity, ParcelType, Weight, Pieces, Void, ToLocID, ToAddID, ToLocName, RefNum, FromAddID, FromCustID, FromCustName, FromLocID, FromLocName, HHid, BatchNum, SignaturePath) " & _
                "VALUES " & Me.SqlValueList

            Return m_strInsertStmt
        End Get
    End Property

    ' Members used for reporting the objects status
    Private m_bIsEmpty As Boolean
    Public ReadOnly Property IsEmpty() As Boolean
        Get
            Return m_bIsEmpty
        End Get
    End Property

    ' Members that map to each row
    Private m_vcEventCode As VarChar
    Public Property EventCode() As String
        Get
            Return m_vcEventCode.Value
        End Get
        Set(ByVal Value As String)
            m_vcEventCode.Value = Value
        End Set
    End Property

    Private m_dtScanDate As Date
    Public Property ScanDate() As Date
        Get
            Return m_dtScanDate
        End Get
        Set(ByVal Value As Date)
            m_dtScanDate = Value
        End Set
    End Property

    Private m_oOperatorId As TPCOperatorBC
    Public Property OperatorId() As TPCOperatorBC
        Get
            Return m_oOperatorId
        End Get
        Set(ByVal Value As TPCOperatorBC)
            m_oOperatorId = Value
        End Set
    End Property

    Private m_oPointId As TPCPointBC
    Public Property PointId() As TPCPointBC
        Get
            Return m_oPointId
        End Get
        Set(ByVal Value As TPCPointBC)
            m_oPointId = Value
        End Set
    End Property

    Private m_vcTicketNum As VarChar
    Public Property TicketNum() As String
        Get
            Return m_vcTicketNum.Value
        End Get
        Set(ByVal Value As String)
            m_vcTicketNum.Value = Value
        End Set
    End Property

    Private m_oTrackingNum As TPCBarcode
    Public Property TrackingNumber() As TPCBarcode
        Get
            Return m_oTrackingNum
        End Get
        Set(ByVal Value As TPCBarcode)
            m_oTrackingNum = Value
        End Set
    End Property

    Private m_oThirdPartyBarcode As Barcode
    Public Property ThirdPartyBarcode() As Barcode
        Get
            Return m_oThirdPartyBarcode
        End Get
        Set(ByVal Value As Barcode)
            m_oThirdPartyBarcode = Value
        End Set
    End Property

    Private m_oContainerBarcode As TPCBarcode
    Public Property ContainerBarcode() As TPCBarcode
        Get
            Return m_oContainerBarcode
        End Get
        Set(ByVal Value As TPCBarcode)
            m_oContainerBarcode = Value
        End Set
    End Property

    Private m_vcDeliveryOption As VarChar
    Public Property DeliveryOption() As String
        Get
            Return m_vcDeliveryOption.Value
        End Get
        Set(ByVal Value As String)
            m_vcDeliveryOption.Value = Value
        End Set
    End Property

    Private m_vcDeliveryComments As VarChar
    Public Property DeliveryComments() As String
        Get
            Return m_vcDeliveryComments.Value
        End Get
        Set(ByVal Value As String)
            m_vcDeliveryComments.Value = Value
        End Set
    End Property

    Private m_vcToCity As VarChar
    Public Property ToCity() As String
        Get
            Return m_vcToCity.Value
        End Get
        Set(ByVal Value As String)
            m_vcToCity.Value = Value
        End Set
    End Property

    Private m_vcParcelType As VarChar
    Public Property ParcelType() As String
        Get
            Return m_vcParcelType.Value
        End Get
        Set(ByVal Value As String)
            m_vcParcelType.Value = Value
        End Set
    End Property

    Private m_Weight As Decimal
    Public Property Weight() As Decimal
        Get
            Return m_Weight
        End Get
        Set(ByVal Value As Decimal)
            m_Weight = Value
        End Set
    End Property

    Private m_vcPieces As VarChar
    Public Property Pieces() As String
        Get
            Return m_vcPieces.Value
        End Get
        Set(ByVal Value As String)
            m_vcPieces.Value = Value
        End Set
    End Property

    Private m_bVoid As Boolean
    Public Property Void() As Boolean
        Get
            Return m_bVoid
        End Get
        Set(ByVal Value As Boolean)
            m_bVoid = Value
        End Set
    End Property

    Private m_vcToLocId As VarChar
    Public Property ToLocationId() As String
        Get
            Return m_vcToLocId.Value
        End Get
        Set(ByVal Value As String)
            m_vcToLocId.Value = Value
        End Set
    End Property

    Private m_iToAddId As Integer
    Public Property ToAddressId() As Integer
        Get
            Return m_iToAddId
        End Get
        Set(ByVal Value As Integer)
            m_iToAddId = Value
        End Set
    End Property

    Private m_vcToLocName As VarChar
    Public Property ToLocationName() As String
        Get
            Return m_vcToLocName.Value
        End Get
        Set(ByVal Value As String)
            m_vcToLocName.Value = Value
        End Set
    End Property

    Private m_vcRefNum As VarChar
    Public Property ReferenceNumber() As String
        Get
            Return m_vcRefNum.Value
        End Get
        Set(ByVal Value As String)
            m_vcRefNum.Value = Value
        End Set
    End Property

    Private m_iRowId As Integer ' Auto-Increment
    Public Property RowId() As Integer
        Get
            Return m_iRowId
        End Get
        Set(ByVal Value As Integer)
            m_iRowId = Value
        End Set
    End Property

    Private m_iFromAddId As Integer
    Public Property FromAddressId() As Integer
        Get
            Return m_iFromAddId
        End Get
        Set(ByVal Value As Integer)
            m_iFromAddId = Value
        End Set
    End Property

    Private m_vcFromCustId As VarChar
    Public Property FromCustomerId() As String
        Get
            Return m_vcFromCustId.Value
        End Get
        Set(ByVal Value As String)
            m_vcFromCustId.Value = Value
        End Set
    End Property

    Private m_vcFromCustName As VarChar
    Public Property FromCustomerName() As String
        Get
            Return m_vcFromCustName.Value
        End Get
        Set(ByVal Value As String)
            m_vcFromCustName.Value = Value
        End Set
    End Property

    Private m_vcFromLocId As VarChar
    Public Property FromLocationId() As String
        Get
            Return m_vcFromLocId.Value
        End Get
        Set(ByVal Value As String)
            m_vcFromLocId.Value = Value
        End Set
    End Property

    Private m_vcFromLocName As VarChar
    Public Property FromLocationName() As String
        Get
            Return m_vcFromLocName.Value
        End Get
        Set(ByVal Value As String)
            m_vcFromLocName.Value = Value
        End Set
    End Property

    Private m_vcHHid As VarChar
    Public Property ScannerId() As String
        Get
            Return m_vcHHid.Value
        End Get
        Set(ByVal Value As String)
            m_vcHHid.Value = Value
        End Set
    End Property

    Private m_vcBatchNum As VarChar
    Public Property BatchNumber() As String
        Get
            Return m_vcBatchNum.Value
        End Get
        Set(ByVal Value As String)
            m_vcBatchNum.Value = Value
        End Set
    End Property

    Private m_vcSignaturePath As VarChar
    Public Property SignatureFile() As String
        Get
            Return m_vcSignaturePath.Value
        End Get
        Set(ByVal Value As String)
            m_vcSignaturePath.Value = Value
        End Set
    End Property

    Private Sub CleanUp()

        Try

            m_oDataAdapter = Nothing
            m_oDataSet = Nothing
            m_strSelectStmt = ""
            m_strInsertStmt = ""

        Catch ex As Exception

            Return

        End Try

    End Sub

    Public ReadOnly Property SqlValueList() As String

        Get

            Dim strValueList As String

            strValueList = "("

            strValueList += "'" & Me.EventCode & "',"

            strValueList += "'" & Me.ScanDate & "',"

            If IsNothing(Me.OperatorId) Then
                strValueList += "'',"
            Else
                strValueList += "'" & Me.OperatorId.ToString & "',"
            End If

            If IsNothing(Me.PointId) Then
                strValueList += "'',"
            Else
                strValueList += "'" & Me.PointId.ToString & "',"
            End If

            strValueList += "'" & Me.TicketNum & "',"

            If IsNothing(Me.TrackingNumber) Then
                strValueList += "'',"
            Else
                strValueList += "'" & Me.TrackingNumber.ToString & "',"
            End If

            If IsNothing(Me.ThirdPartyBarcode) Then
                strValueList += "'',"
            Else
                strValueList += "'" & Me.ThirdPartyBarcode.ToString & "',"
            End If

            If IsNothing(Me.ContainerBarcode) Then
                strValueList += "'',"
            Else
                strValueList += "'" & Me.ContainerBarcode.ToString & "',"
            End If

            strValueList += "'" & Me.DeliveryOption & "',"

            strValueList += "'" & Me.DeliveryComments & "',"

            strValueList += "'" & Me.ToCity & "',"

            strValueList += "'" & Me.ParcelType & "',"

            If Me.Weight = -1 Then
                strValueList += "NULL,"
            Else
                strValueList += Me.Weight.ToString & ","
            End If

            strValueList += "'" & Me.Pieces & "',"

            If Me.Void = True Then
                strValueList += "'T',"
            Else
                strValueList += "'F',"
            End If

            strValueList += "'" & Me.ToLocationId & "',"

            If Me.ToAddressId = -1 Then
                strValueList += "NULL,"
            Else
                strValueList += Me.ToAddressId & ","
            End If

            strValueList += "'" & Me.ToLocationName & "',"

            strValueList += "'" & Me.ReferenceNumber & "',"

            If Me.FromAddressId = -1 Then
                strValueList += "NULL,"
            Else
                strValueList += Me.FromAddressId.ToString & ","
            End If

            strValueList += "'" & Me.FromCustomerId & "',"

            strValueList += "'" & Me.FromCustomerName & "',"

            strValueList += "'" & Me.FromLocationId & "',"

            strValueList += "'" & Me.FromLocationName & "',"

            strValueList += "'" & Me.ScannerId & "',"

            strValueList += "'" & Me.BatchNumber & "',"

            strValueList += "'" & Me.SignatureFile & "'"

            strValueList += ")"

            Return strValueList

        End Get

    End Property

    Public Function Insert() As Boolean

        Return ExecuteQuery(Me.InsertStatement)

    End Function

    Sub New()

        ' Initialize a new instance of this object
        m_vcEventCode = New VarChar(2)
        m_dtScanDate = Date.Now
        m_vcTicketNum = New VarChar(7)
        m_vcDeliveryOption = New VarChar(1)
        m_vcDeliveryComments = New VarChar(50)
        m_vcToCity = New VarChar(32)
        m_vcParcelType = New VarChar(20)
        m_Weight = -1 ' -1 Translates to NULL in the Database, which is not the same as 0
        m_vcPieces = New VarChar(10)
        m_vcToLocId = New VarChar(10)
        m_iToAddId = -1 ' -1 Translates to NULL in the Database, which is not the same as 0
        m_vcToLocName = New VarChar(70)
        m_vcRefNum = New VarChar(40)
        m_iFromAddId = -1 ' -1 Translates to NULL in the Database, which is not the same as 0
        m_vcFromCustId = New VarChar(10)
        m_vcFromCustName = New VarChar(70)
        m_vcFromLocId = New VarChar(10)
        m_vcFromLocName = New VarChar(70)
        m_vcHHid = New VarChar(4)
        m_vcBatchNum = New VarChar(40)
        m_vcSignaturePath = New VarChar(255)

        ' Initialize Status Members
        m_bIsEmpty = True

    End Sub


End Class
