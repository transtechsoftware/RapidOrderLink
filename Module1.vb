Imports System.Data
Imports System.Text
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports TTSI.BARCODES
Imports TTSI.UTILITES
Imports System.IO


Module Module1

    Public TRCDBName As String = "UN_TRACKING"
    Public TRCTblPath As String = TRCDBName & ".dbo."

    Public WEIGHTDBName As String = "UN_WEIGHT" '"RoutesModule"
    Public WEIGHTTblPath As String = WEIGHTDBName & ".dbo."

    Public AppDBName As String = "UNISON" ' "UNISON" '
    Public AppTblPath As String = AppDBName & ".dbo."

    'Public dtLastLocationFix As DateTime = Date.Now
    Public dtLastLocationFix As DateTime = "01/01/2014"

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    ' Connection String
    Private strConnection As String = "Server=12.158.238.68;Database=UN_ORDERS;Uid=RUSR1;Pwd=TOPPRI;"
    'Private strConnection As String = "Server=67.112.189.227;Database=UN_ORDERS;Uid=RUSR1;Pwd=TOPPRI;"
    Private sScanListConnection As String = "Server=12.158.238.68;Database=UN_TRACKING;Uid=sa;Pwd=4183771.4604$T;"
    'Private sScanListConnection As String = "Server=67.112.189.227;Database=UN_TRACKING;Uid=sa;Pwd=4183771.4604$T;"

    ' Connection Objects
    Private cnSqlQuery As SqlConnection ' For main loop
    Private cnSqlInsert As SqlConnection ' For calling spInsertOrder
    Private cnSqlUpdate As SqlConnection ' For calling spUpdateOrder
    Private cnSqlDelete As SqlConnection ' For calling spDeleteOrder
    Private cnSqlInsertPUM As SqlConnection ' For calling spInsertPUManifestRecord

    ' Command Objects
    Private cmd As New SqlCommand
    Private cmdInsert As SqlCommand
    Private cmdUpdate As SqlCommand
    Private cmdDelete As SqlCommand
    Private cmdInsertPUM As SqlCommand

    ' Parameter Object
    Private prmSqlParameter As SqlParameter

    ' Data Readers
    Private drReader As SqlDataReader

    ' File System Object
    Private sIngramMicroPath As String = "D:\FTPHOME\INGRAM-MICRO\INBOUND"
    Private sPartnerBooksPath As String = "D:\FTPHOME\PARTNER-BOOKS\INBOUND"
    Private sMedExpressPath As String = "D:\FTPHOME\MEDEXPRESS\INBOUND"
    Private sBluePackagePath As String = "D:\FTPHOME\BLUEPACKAGE\INBOUND"
    Private sMedExpressSdPath As String = "D:\FTPHOME\MEDEXPRESSSD\INBOUND"
    Private sAutoZonePath As String = "D:\FTPHOME\AUTOZONE\INBOUND"
    Private sDCSDeliveryPath As String = "D:\FTPHOME\DCS\INBOUND"
    Private sTrackingPath As String = "D:\FTPHOME\TRACKING"
    Private sPiaDeliveryPath As String = "D:\FTPHOME\PIA"
    Private sProCourierPath As String = "D:\FTPHOME\PROCOURIER"
    Private sDio2CsPath As String = "C:\FTPHOME\DIO2CS"

    Function AnythingChanged(ByVal p_iRapidOrderID As Integer) As Boolean

        Dim drUnison, drRapid As SqlDataReader
        Dim cnUnison As New SqlConnection(strConnection)
        Dim cnRapid As New SqlConnection(strConnection)
        Dim cmdUnison, cmdRapid As SqlCommand
        Dim strUnisonQuery As String = "SELECT TOP 1 OrderID, InvoiceID FROM UN_ORDERS.dbo.RapidOrderHistory roh WHERE OrderID = @a ORDER BY LastModifiedDate DESC"
        Dim strRapidQuery As String = "SELECT TOP 1 [ID] AS OrderID, InvoiceID FROM NETSERVER3.RTMGMT.dbo.Orders o WHERE [ID] = @a ORDER BY ModifiedDate DESC"

        strUnisonQuery = strUnisonQuery.Replace("@a", p_iRapidOrderID)
        strRapidQuery = strRapidQuery.Replace("@a", p_iRapidOrderID)

        Try
            ' Open the connection
            If Not cnUnison.State = ConnectionState.Open Then cnUnison.Open()
            If Not cnRapid.State = ConnectionState.Open Then cnRapid.Open()

            ' Populate the data readers
            cmdUnison = New SqlCommand(strUnisonQuery, cnUnison)
            cmdRapid = New SqlCommand(strRapidQuery, cnRapid)

            drUnison = cmdUnison.ExecuteReader
            drRapid = cmdRapid.ExecuteReader

            ' Compare each column in Unison to Correspoinding Rapid Column.  If any one is different, return true.  If all the same return false
            If drUnison.HasRows And drRapid.HasRows Then

                drUnison.Read()
                drRapid.Read()

                If Trim(CStr(drUnison("InvoiceID"))).ToUpper <> Trim(CStr(drRapid("InvoiceID"))).ToUpper Then Return True

            End If

            Return False

        Catch ex As Exception

            Dim s As String = ex.Message
            s = s

        Finally

            cnUnison.Close()
            cnRapid.Close()

        End Try

    End Function

    Sub ProcessOrders()

        ' Instantiate Connections
        cnSqlQuery = New SqlConnection(strConnection)
        cnSqlInsert = New SqlConnection(strConnection)
        cnSqlUpdate = New SqlConnection(strConnection)
        cnSqlDelete = New SqlConnection(strConnection)

        ' Instantiate Commands
        cmd = New SqlCommand("SELECT * FROM NETSERVER3.RTMGMT.dbo.Orders_Q WHERE Processed = 'N' ORDER BY RowId ASC", cnSqlQuery)
        'cmdInsert = New SqlCommand("spInsertOrder", cnSqlInsert)
        cmdInsert = New SqlCommand("spInsertOrderHistory", cnSqlInsert)
        'cmdUpdate = New SqlCommand("spUpdateOrder", cnSqlUpdate)
        cmdUpdate = New SqlCommand("spUpdateOrderHistory", cnSqlUpdate)
        'cmdDelete = New SqlCommand("spDeleteOrder", cnSqlDelete)
        cmdDelete = New SqlCommand("spDeleteOrderHistory", cnSqlDelete)

        ' Set Command Types
        cmdInsert.CommandType = CommandType.StoredProcedure
        cmdUpdate.CommandType = CommandType.StoredProcedure
        cmdDelete.CommandType = CommandType.StoredProcedure

        ' Flags
        Dim bUpdateVisual As Boolean = False

        Try
            'open the connection
            If Not cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Open()

            'populate the DataReader
            drReader = cmd.ExecuteReader

            'Provide Visual Feedback
            If drReader.HasRows Then
                Console.WriteLine("Changes to ORDERS detected.")
                Console.Write("Processing...")
                bUpdateVisual = True
            End If

            'loop through the results and do something
            Do While drReader.Read

                'Update Visual Feedback
                Console.Write(".")

                'Open Stored Procedure Connetions
                If Not cnSqlInsert.State = ConnectionState.Open Then cnSqlInsert.Open()
                If Not cnSqlUpdate.State = ConnectionState.Open Then cnSqlUpdate.Open()
                If Not cnSqlDelete.State = ConnectionState.Open Then cnSqlDelete.Open()

                'Console.WriteLine(drReader("OrderID") & " - " & drReader("RowID"))
                Dim strEvent As String = drReader("Event")
                Dim strRapidID As String = drReader("OrderID")
                Dim strRowID As String = drReader("RowID")
                Dim strCaller As String = drReader("Caller")
                Dim strWebOrderID As String = "" 'Add code to extract weborderid from "[WEBxyz] plus Any Text"

                Select Case strEvent
                    Case "I"
                        ' @RAPIDID
                        'prmSqlParameter = cmdInsert.Parameters.Add("@RAPIDID", strRapidID)
                        prmSqlParameter = cmdInsert.Parameters.AddWithValue("@RAPIDID", strRapidID)
                        prmSqlParameter.Direction = ParameterDirection.Input
                        prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' @Q_ROWID
                        'prmSqlParameter = cmdInsert.Parameters.Add("@Q_ROWID", strRowID)
                        prmSqlParameter = cmdInsert.Parameters.AddWithValue("@Q_ROWID", strRowID)
                        prmSqlParameter.Direction = ParameterDirection.Input
                        prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' @WEBORDERID
                        ''prmSqlParameter = cmdInsert.Parameters.Add("@WEBORDERID", Nothing)
                        ''prmSqlParameter.Direction = ParameterDirection.Input
                        ''prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' Execute Command
                        cmdInsert.ExecuteNonQuery()
                        cmdInsert.Parameters.Clear()
                    Case "U"
                        ' @Q_Only
                        If AnythingChanged(strRapidID) Then
                            prmSqlParameter = cmdUpdate.Parameters.AddWithValue("@Q_Only", "1")
                            prmSqlParameter.Direction = ParameterDirection.Input
                            prmSqlParameter.SqlDbType = SqlDbType.Int
                            ' @RAPIDID
                            prmSqlParameter = cmdUpdate.Parameters.AddWithValue("@RAPIDID", strRapidID)
                            prmSqlParameter.Direction = ParameterDirection.Input
                            prmSqlParameter.SqlDbType = SqlDbType.Int
                            ' @Q_ROWID
                            prmSqlParameter = cmdUpdate.Parameters.AddWithValue("@Q_ROWID", strRowID)
                            prmSqlParameter.Direction = ParameterDirection.Input
                            prmSqlParameter.SqlDbType = SqlDbType.Int
                        Else
                            prmSqlParameter = cmdUpdate.Parameters.AddWithValue("@Q_Only", "0")
                            prmSqlParameter.Direction = ParameterDirection.Input
                            prmSqlParameter.SqlDbType = SqlDbType.Int
                            ' @RAPIDID
                            prmSqlParameter = cmdUpdate.Parameters.AddWithValue("@RAPIDID", strRapidID)
                            prmSqlParameter.Direction = ParameterDirection.Input
                            prmSqlParameter.SqlDbType = SqlDbType.Int
                            ' @Q_ROWID
                            prmSqlParameter = cmdUpdate.Parameters.AddWithValue("@Q_ROWID", strRowID)
                            prmSqlParameter.Direction = ParameterDirection.Input
                            prmSqlParameter.SqlDbType = SqlDbType.Int
                        End If
                        ' Execute Command
                        cmdUpdate.ExecuteNonQuery()
                        cmdUpdate.Parameters.Clear()
                    Case "D"
                        ' @RAPIDID
                        ''prmSqlParameter = cmdDelete.Parameters.Add("@RAPIDID", strRapidID)
                        ''prmSqlParameter.Direction = ParameterDirection.Input
                        ''prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' @Q_ROWID
                        ''prmSqlParameter = cmdDelete.Parameters.Add("@Q_ROWID", strRowID)
                        ''prmSqlParameter.Direction = ParameterDirection.Input
                        ''prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' Execute Command
                        ''cmdDelete.ExecuteNonQuery()
                        ''cmdDelete.Parameters.Clear()
                    Case Else
                        Console.WriteLine("[Module1.225] Unknown Event Type - " & strEvent)
                End Select
            Loop

            If bUpdateVisual = True Then
                Console.WriteLine("Done!")
                Console.Write("Waiting for Changes to Rapid...")
            End If

        Catch ex As Exception
            'Error Handling Code Goes Here
            Console.WriteLine("[Module1.236]" & ex.Message)
        Finally
            'clean up code that need to run no matter what

            'close the data reader
            drReader.Close()

            'close the connections
            cnSqlQuery.Close()
            cnSqlInsert.Close()
            cnSqlUpdate.Close()
            cnSqlDelete.Close()

            If Not cnSqlQuery Is Nothing Then
                cnSqlQuery.Dispose()
            End If

            If Not cnSqlInsert Is Nothing Then
                cnSqlInsert.Dispose()
            End If

            If Not cnSqlUpdate Is Nothing Then
                cnSqlUpdate.Dispose()
            End If

            If Not cnSqlDelete Is Nothing Then
                cnSqlDelete.Dispose()
            End If

        End Try

    End Sub

    Sub ProcessOrderCOD()

        ' Instantiate Connections
        cnSqlQuery = New SqlConnection(strConnection)
        cnSqlInsert = New SqlConnection(strConnection)
        cnSqlUpdate = New SqlConnection(strConnection)
        cnSqlDelete = New SqlConnection(strConnection)

        ' Instantiate Commands
        cmd = New SqlCommand("SELECT * FROM NETSERVER3.RTMGMT.dbo.OrderCOD_Q WHERE Processed = 'N' ORDER BY RowId ASC", cnSqlQuery)
        cmdInsert = New SqlCommand("spInsertOrderCOD", cnSqlInsert)
        cmdUpdate = New SqlCommand("spUpdateOrderCOD", cnSqlUpdate)
        cmdDelete = New SqlCommand("spDeleteOrderCOD", cnSqlDelete)

        ' Set Command Types
        cmdInsert.CommandType = CommandType.StoredProcedure
        cmdUpdate.CommandType = CommandType.StoredProcedure
        cmdDelete.CommandType = CommandType.StoredProcedure

        ' Flags
        Dim bUpdateVisual As Boolean = False

        Try
            'open the connection
            If Not cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Open()

            'populate the DataReader
            drReader = cmd.ExecuteReader

            'Provide Visual Feedback
            If drReader.HasRows Then
                Console.WriteLine("Changes to ORDERCOD detected.")
                Console.Write("Processing...")
                bUpdateVisual = True
            End If

            'loop through the results and do something
            Do While drReader.Read

                'Update Visual Feedback
                Console.Write(".")

                'Open Stored Procedure Connetions
                If Not cnSqlInsert.State = ConnectionState.Open Then cnSqlInsert.Open()
                If Not cnSqlUpdate.State = ConnectionState.Open Then cnSqlUpdate.Open()
                If Not cnSqlDelete.State = ConnectionState.Open Then cnSqlDelete.Open()

                'Console.WriteLine(drReader("OrderID") & " - " & drReader("RowID"))
                Dim strEvent As String = drReader("Event")
                Dim strRapidID As String = drReader("OrderID")
                Dim strRowID As String = drReader("RowID")

                Select Case strEvent
                    Case "I"
                        ' @RAPIDID
                        prmSqlParameter = cmdInsert.Parameters.AddWithValue("@RAPIDID", strRapidID)
                        prmSqlParameter.Direction = ParameterDirection.Input
                        prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' @Q_ROWID
                        prmSqlParameter = cmdInsert.Parameters.AddWithValue("@Q_ROWID", strRowID)
                        prmSqlParameter.Direction = ParameterDirection.Input
                        prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' Execute Command
                        cmdInsert.ExecuteNonQuery()
                        cmdInsert.Parameters.Clear()
                    Case "U"
                        ' @RAPIDID
                        prmSqlParameter = cmdUpdate.Parameters.AddWithValue("@RAPIDID", strRapidID)
                        prmSqlParameter.Direction = ParameterDirection.Input
                        prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' @Q_ROWID
                        prmSqlParameter = cmdUpdate.Parameters.AddWithValue("@Q_ROWID", strRowID)
                        prmSqlParameter.Direction = ParameterDirection.Input
                        prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' Execute Command
                        cmdUpdate.ExecuteNonQuery()
                        cmdUpdate.Parameters.Clear()
                    Case "D"
                        ' @RAPIDID
                        prmSqlParameter = cmdDelete.Parameters.AddWithValue("@RAPIDID", strRapidID)
                        prmSqlParameter.Direction = ParameterDirection.Input
                        prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' @Q_ROWID
                        prmSqlParameter = cmdDelete.Parameters.AddWithValue("@Q_ROWID", strRowID)
                        prmSqlParameter.Direction = ParameterDirection.Input
                        prmSqlParameter.SqlDbType = SqlDbType.Int
                        ' Execute Command
                        cmdDelete.ExecuteNonQuery()
                        cmdDelete.Parameters.Clear()
                    Case Else
                        Console.WriteLine("[Module1.359] Unknown Event Type - " & strEvent)
                End Select
            Loop

            If bUpdateVisual = True Then
                Console.WriteLine("Done!")
                Console.Write("Waiting for Changes to Rapid...")
            End If

        Catch ex As Exception
            'Error Handling Code Goes Here
            Console.WriteLine("[Module1.370]" & ex.Message)
        Finally
            'clean up code that need to run no matter what

            'close the data reader
            drReader.Close()

            'close the connections
            cnSqlQuery.Close()
            cnSqlInsert.Close()
            cnSqlUpdate.Close()
            cnSqlDelete.Close()

            If Not cnSqlQuery Is Nothing Then
                cnSqlQuery.Dispose()
            End If

            If Not cnSqlInsert Is Nothing Then
                cnSqlInsert.Dispose()
            End If

            If Not cnSqlUpdate Is Nothing Then
                cnSqlUpdate.Dispose()
            End If

            If Not cnSqlDelete Is Nothing Then
                cnSqlDelete.Dispose()
            End If

        End Try

    End Sub

    Private Function MatchingLabelFound(ByVal p_sBarcode As String, ByVal p_dtScanDate As Date) As String

        Dim sReturn As String = Nothing
        Dim sb As StringBuilder

        Try

            ' Instantiate Connections
            cnSqlQuery = New SqlConnection(sScanListConnection)

            ' Instantiate Commands
            sb = New StringBuilder
            sb.Append("SELECT count(*) as MatchingRecords FROM ")
            sb.Append(TRCTblPath)
            sb.Append("Event WHERE TrackingNum = '")
            sb.Append(p_sBarcode)
            sb.Append("' AND ScanDate Between '")
            sb.Append(p_dtScanDate.AddDays(-21).ToShortDateString())
            sb.Append("' AND '")
            sb.Append(p_dtScanDate.AddDays(7).ToShortDateString())
            sb.Append("' AND (EventCode = 'L' or EventCode = 'OO')")

            cmd = New SqlCommand(sb.ToString(), cnSqlQuery)

            ' Set Command Types

            'open the connection
            If Not cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Open()

            'populate the DataReader
            drReader = cmd.ExecuteReader

            'Provide Visual Feedback
            If drReader.Read Then

                If CInt(drReader("MatchingRecords")) = 1 Then

                    cmd.Dispose()
                    drReader.Close()

                    sb.Length = 0
                    sb.Append("SELECT ScanDate as LabelDate FROM ")
                    sb.Append(TRCTblPath)
                    sb.Append("Event WHERE TrackingNum = '")
                    sb.Append(p_sBarcode)
                    sb.Append("' AND ScanDate Between '")
                    sb.Append(p_dtScanDate.AddDays(-21).ToShortDateString())
                    sb.Append("' AND '")
                    sb.Append(p_dtScanDate.AddDays(7).ToShortDateString())
                    sb.Append("' AND (EventCode = 'L' or EventCode = 'OO')")

                    cmd = New SqlCommand(sb.ToString(), cnSqlQuery)

                    'open the connection
                    If Not cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Open()

                    'populate the DataReader
                    drReader = cmd.ExecuteReader

                    'Provide Visual Feedback
                    If drReader.Read Then

                        sReturn = ManifestRowId(p_sBarcode, Convert.ToDateTime(drReader("LabelDate")))

                    End If

                End If

            End If



        Catch ex As Exception

            Console.WriteLine("[Module1.477]" & ex.Message)
            sReturn = Nothing

        Finally

            If Not drReader Is Nothing Then

                drReader.Close()

            End If

        End Try

        Return sReturn
    End Function

    'Private Sub ImportIngramMicroManifests()

    '    Dim bResult As Boolean = False
    '    Dim oConn As SqlConnection = Nothing
    '    Dim oCmd As SqlCommand = Nothing
    '    Dim oTran As SqlTransaction = Nothing
    '    Dim oPickupManifest As PickupManifest = Nothing

    '    Dim sFiles() As String = Nothing
    '    Dim sFolders() As String = Nothing
    '    Dim sValidFiles() As String = Nothing
    '    Dim sValidFilesPath() As String = Nothing

    '    Dim sFileNameParts() As String = Nothing
    '    Dim iParts As Integer = 0

    '    Dim sFileName As String = Nothing
    '    Dim iFileName As Integer = 0
    '    Dim iFirstThree As Integer = 0

    '    Dim sFileExtension As String = Nothing

    '    Dim sTmp() As String = Nothing
    '    Dim i As Int32 = 0
    '    Dim j As Int32 = 0

    '    Try

    '        'Get List of Files that Must be Processed
    '        sIngramMicroPath = sIngramMicroPath.ToUpper
    '        If Directory.Exists(sIngramMicroPath) Then
    '            sFiles = Directory.GetFiles(sIngramMicroPath)
    '            For i = 0 To sFiles.Length - 1

    '                sFiles(i) = sFiles(i).ToUpper

    '                sFileNameParts = sFiles(i).Split(".")
    '                iParts = sFileNameParts.Length

    '                sFolders = sFileNameParts(iParts - 2).Split("\")

    '                sFileName = sFolders(sFolders.Length - 1)
    '                sFileExtension = sFileNameParts(iParts - 1)

    '                iFileName = sFileName.Length
    '                If iFileName >= 3 Then iFirstThree = 3 Else iFirstThree = iFileName


    '                'If (sFileNameParts(iParts - 2).Substring(1, 3).CompareTo("AIM") = 0) And (sFileNameParts(iParts - 1).CompareTo("UPD") = 0) Then
    '                If (sFileName.Substring(0, iFirstThree).CompareTo("AIM") = 0) And (sFileExtension.CompareTo("UPD") = 0) Then

    '                    ReDim Preserve sValidFiles(j)
    '                    ReDim Preserve sValidFilesPath(j)

    '                    Dim iNameLength As Integer = sFiles(i).Length
    '                    sValidFiles(j) = sFiles(i).Substring(iNameLength - 23, 23)
    '                    sValidFilesPath(j) = sFiles(i).Substring(0, iNameLength - 23)
    '                    j += 1
    '                End If
    '            Next
    '        Else
    '            Console.WriteLine("Path does not exist for Ingram-Micro:" & sIngramMicroPath)
    '            Exit Sub
    '        End If


    '        If Not sValidFiles Is Nothing Then

    '            Console.WriteLine("Ingram-Micro Files Detected. Beginning Import.")

    '            ' We can use the same SqlConnection for the entire batch of files
    '            If oConn Is Nothing Then
    '                oConn = New SqlConnection(sScanListConnection)
    '                oConn.Open()
    '            End If

    '            ' Process Batch of Valid Files
    '            For i = 0 To sValidFiles.Length - 1

    '                ' Either the entire file is processed and renamed or the operation on that file fails
    '                oTran = oConn.BeginTransaction()
    '                oCmd = New SqlCommand("spInsertPuManifestRecordV0", oConn, oTran)
    '                oCmd.CommandType = CommandType.StoredProcedure

    '                ' Read in the next file
    '                oPickupManifest = New PickupManifest(sValidFilesPath(i) & sValidFiles(i))

    '                If Not oPickupManifest.Records Is Nothing Then

    '                    bResult = InsertManifestRecords(oPickupManifest.Records, oCmd, "V0")
    '                    If bResult Then
    '                        sTmp = sValidFiles(i).Split(".")
    '                        sTmp(1) = sTmp(1) & "_ARC"

    '                        Dim sOldFileName As String = sValidFilesPath(i) & sValidFiles(i)
    '                        Dim sNewFileName As String = sValidFilesPath(i) & sTmp(0) & "." & sTmp(1)

    '                        'File.Move(sValidFiles(i), sTmp(0) & "." & sTmp(1))
    '                        File.Move(sOldFileName, sNewFileName)
    '                        sTmp = Nothing
    '                        Console.WriteLine(sValidFiles(i) & " was successfully processed")
    '                    Else
    '                        Console.WriteLine(sValidFiles(i) & " was NOT successfully processed")
    '                    End If

    '                Else

    '                    ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                    Console.WriteLine("[Module1.601]" & oPickupManifest.ErrorMessage & ". " & sValidFiles(i) & "failed to process.")
    '                    bResult = False

    '                End If

    '                If bResult = True Then
    '                    oTran.Commit()
    '                Else
    '                    oTran.Rollback()
    '                End If

    '                oTran = Nothing
    '                oCmd = Nothing

    '            Next

    '        End If


    '    Catch ex As Exception

    '        Console.WriteLine(ex.Message, "ImportIngramMicroManifests() Status")
    '        If Not oTran Is Nothing Then oTran.Rollback()

    '    Finally

    '        oCmd = Nothing
    '        oTran = Nothing

    '        If Not oConn Is Nothing Then
    '            oConn.Close()
    '            oConn.Dispose()
    '        End If

    '    End Try


    'End Sub

    Private Sub ImportPartnerBooksManifests()

        Dim bResult As Boolean = False
        Dim oConn As SqlConnection = Nothing
        Dim oCmd As SqlCommand = Nothing
        Dim oTran As SqlTransaction = Nothing
        Dim oPickupManifest As PickupManifest = Nothing

        Dim sFiles() As String = Nothing
        Dim sFolders() As String = Nothing
        Dim sValidFiles() As String = Nothing
        Dim sValidFilesPath() As String = Nothing

        Dim sFileNameParts() As String = Nothing
        Dim iParts As Integer = 0

        Dim sFileName As String = Nothing
        Dim iFileName As Integer = 0
        Dim iFirstThree As Integer = 0

        Dim sFileExtension As String = Nothing

        Dim sTmp() As String = Nothing
        Dim i As Int32 = 0
        Dim j As Int32 = 0

        Try

            'Get List of Files that Must be Processed
            sPartnerBooksPath = sPartnerBooksPath.ToUpper
            If Directory.Exists(sPartnerBooksPath) Then
                sFiles = Directory.GetFiles(sPartnerBooksPath)
                For i = 0 To sFiles.Length - 1

                    sFiles(i) = sFiles(i).ToUpper

                    sFileNameParts = sFiles(i).Split(".")
                    iParts = sFileNameParts.Length

                    sFolders = sFileNameParts(iParts - 2).Split("\")

                    sFileName = sFolders(sFolders.Length - 1)
                    sFileExtension = sFileNameParts(iParts - 1)

                    iFileName = sFileName.Length

                    If IsNumeric(sFileName.Substring(1, 6)) And (iFileName = 7) And (sFileExtension.CompareTo("TXT") = 0) Then

                        ReDim Preserve sValidFiles(j)
                        ReDim Preserve sValidFilesPath(j)

                        Dim iNameLength As Integer = sFiles(i).Length
                        sValidFiles(j) = sFiles(i).Substring(iNameLength - 11, 11)
                        sValidFilesPath(j) = sFiles(i).Substring(0, iNameLength - 11)
                        j += 1

                    End If
                Next
            Else
                Console.WriteLine("Path does not exist for PARTNERS WEST DISTRIBUTING:" & sPartnerBooksPath)
                Exit Sub
            End If


            If Not sValidFiles Is Nothing Then

                Console.WriteLine("Partner Books Files Detected. Beginning Import.")

                ' We can use the same SqlConnection for the entire batch of files
                If oConn Is Nothing Then
                    oConn = New SqlConnection(sScanListConnection)
                    oConn.Open()
                End If

                ' Process Batch of Valid Files
                For i = 0 To sValidFiles.Length - 1

                    ' Either the entire file is processed and renamed or the operation on that file fails
                    oTran = oConn.BeginTransaction()
                    oCmd = New SqlCommand("spInsertPuManifestRecordV1", oConn, oTran)
                    oCmd.CommandType = CommandType.StoredProcedure

                    ' Read in the next file
                    oPickupManifest = New PickupManifestMapperV1(sValidFilesPath(i) & sValidFiles(i))

                    If Not oPickupManifest.Records Is Nothing Then

                        bResult = InsertManifestRecords(oPickupManifest.Records, oCmd, "V1")
                        If bResult Then
                            sTmp = sValidFiles(i).Split(".")
                            sTmp(1) = sTmp(1) & "_ARC"

                            Dim sOldFileName As String = sValidFilesPath(i) & sValidFiles(i)
                            Dim sNewFileName As String = sValidFilesPath(i) & sTmp(0) & "." & sTmp(1)

                            'File.Move(sValidFiles(i), sTmp(0) & "." & sTmp(1))
                            File.Move(sOldFileName, sNewFileName)
                            sTmp = Nothing
                            Console.WriteLine(sValidFiles(i) & " was successfully processed")
                        Else
                            Console.WriteLine(sValidFiles(i) & " was NOT successfully processed")
                        End If

                    Else

                        ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Console.WriteLine("[Module1.746]" & oPickupManifest.ErrorMessage & ". " & sValidFiles(i) & "failed to process.")
                        bResult = False

                    End If

                    If bResult = True Then
                        oTran.Commit()
                    Else
                        oTran.Rollback()
                    End If

                    oTran = Nothing
                    oCmd = Nothing

                Next

            End If


        Catch ex As Exception

            Console.WriteLine(ex.Message, "ImportPartnerBooksManifests() Status")
            If Not oTran Is Nothing Then oTran.Rollback()

        Finally

            oCmd = Nothing
            oTran = Nothing

            If Not oConn Is Nothing Then
                oConn.Close()
                oConn.Dispose()
            End If

        End Try


    End Sub

    Private Sub ImportMedExpressManifests()

        Dim bResult As Boolean = False
        Dim oConn As SqlConnection = Nothing
        Dim oCmd As SqlCommand = Nothing
        Dim oTran As SqlTransaction = Nothing
        Dim oPickupManifest As PickupManifest = Nothing

        Dim sFiles() As String = Nothing
        Dim sFolders() As String = Nothing
        Dim sValidFiles() As String = Nothing
        Dim sValidFilesPath() As String = Nothing

        Dim sFileNameParts() As String = Nothing
        Dim iParts As Integer = 0

        Dim sFileName As String = Nothing
        Dim iFileName As Integer = 0
        Dim iFirstThree As Integer = 0

        Dim sFileExtension As String = Nothing

        Dim sTmp() As String = Nothing
        Dim i As Int32 = 0
        Dim j As Int32 = 0

        Try

            'Get List of Files that Must be Processed
            sMedExpressPath = sMedExpressPath.ToUpper
            If Directory.Exists(sMedExpressPath) Then
                sFiles = Directory.GetFiles(sMedExpressPath)
                For i = 0 To sFiles.Length - 1

                    sFiles(i) = sFiles(i).ToUpper

                    sFileNameParts = sFiles(i).Split(".")
                    iParts = sFileNameParts.Length

                    sFolders = sFileNameParts(iParts - 2).Split("\")

                    sFileName = sFolders(sFolders.Length - 1)
                    sFileExtension = sFileNameParts(iParts - 1)

                    iFileName = sFileName.Length

                    If IsNumeric(sFileName.Substring(0, 8)) And (iFileName = 11) And (sFileExtension.CompareTo("TXT") = 0) Then

                        ReDim Preserve sValidFiles(j)
                        ReDim Preserve sValidFilesPath(j)

                        Dim iNameLength As Integer = sFiles(i).Length
                        sValidFiles(j) = sFiles(i).Substring(iNameLength - 15, 15)
                        sValidFilesPath(j) = sFiles(i).Substring(0, iNameLength - 15)
                        j += 1

                    End If
                Next
            Else
                Console.WriteLine("Path does not exist for Med Express Manifests:" & sMedExpressPath)
                Exit Sub
            End If


            If Not sValidFiles Is Nothing Then

                Console.WriteLine("Med Express Files Detected. Beginning Import.")

                ' We can use the same SqlConnection for the entire batch of files
                If oConn Is Nothing Then
                    oConn = New SqlConnection(sScanListConnection)
                    oConn.Open()
                End If

                ' Process Batch of Valid Files
                For i = 0 To sValidFiles.Length - 1

                    ' Either the entire file is processed and renamed or the operation on that file fails
                    oTran = oConn.BeginTransaction()
                    oCmd = New SqlCommand("spInsertPuManifestRecordV2", oConn, oTran)
                    oCmd.CommandType = CommandType.StoredProcedure

                    ' Read in the next file
                    oPickupManifest = New PickupManifestMapper(sValidFilesPath(i) & sValidFiles(i), "V2")

                    If Not oPickupManifest.Records Is Nothing Then

                        bResult = InsertManifestRecords(oPickupManifest.Records, oCmd, oPickupManifest.FileVersion)
                        If bResult Then
                            sTmp = sValidFiles(i).Split(".")
                            sTmp(1) = sTmp(1) & "_ARC"

                            Dim sOldFileName As String = sValidFilesPath(i) & sValidFiles(i)
                            Dim sNewFileName As String = sValidFilesPath(i) & sTmp(0) & "." & sTmp(1)

                            'File.Move(sValidFiles(i), sTmp(0) & "." & sTmp(1))
                            File.Move(sOldFileName, sNewFileName)
                            sTmp = Nothing
                            Console.WriteLine(sValidFiles(i) & " was successfully processed")
                        Else
                            Console.WriteLine(sValidFiles(i) & " was NOT successfully processed")
                        End If

                    Else

                        ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Console.WriteLine("[Module1.891]" & oPickupManifest.ErrorMessage & ". " & sValidFiles(i) & "failed to process.")
                        bResult = False

                    End If

                    If bResult = True Then
                        oTran.Commit()
                    Else
                        oTran.Rollback()
                    End If

                    oTran = Nothing
                    oCmd = Nothing

                Next

            End If


        Catch ex As Exception

            Console.WriteLine(ex.Message, "ImportMedExpressManifests() Status")
            If Not oTran Is Nothing Then oTran.Rollback()

        Finally

            oCmd = Nothing
            oTran = Nothing

            If Not oConn Is Nothing Then
                oConn.Close()
                oConn.Dispose()
            End If

        End Try


    End Sub

    Private Sub ImportMedExpressSdManifests()

        Dim bResult As Boolean = False
        Dim oConn As SqlConnection = Nothing
        Dim oCmd As SqlCommand = Nothing
        Dim oTran As SqlTransaction = Nothing
        Dim oPickupManifest As PickupManifest = Nothing

        Dim sFiles() As String = Nothing
        Dim sFolders() As String = Nothing
        Dim sValidFiles() As String = Nothing
        Dim sValidFilesPath() As String = Nothing

        Dim sFileNameParts() As String = Nothing
        Dim iParts As Integer = 0

        Dim sFileName As String = Nothing
        Dim iFileName As Integer = 0
        Dim iFirstThree As Integer = 0

        Dim sFileExtension As String = Nothing

        Dim sTmp() As String = Nothing
        Dim i As Int32 = 0
        Dim j As Int32 = 0

        Try

            'Get List of Files that Must be Processed
            sMedExpressSdPath = sMedExpressSdPath.ToUpper
            If Directory.Exists(sMedExpressSdPath) Then
                sFiles = Directory.GetFiles(sMedExpressSdPath)
                For i = 0 To sFiles.Length - 1

                    sFiles(i) = sFiles(i).ToUpper

                    sFileNameParts = sFiles(i).Split(".")
                    iParts = sFileNameParts.Length

                    sFolders = sFileNameParts(iParts - 2).Split("\")

                    sFileName = sFolders(sFolders.Length - 1)
                    sFileExtension = sFileNameParts(iParts - 1)

                    iFileName = sFileName.Length

                    If IsNumeric(sFileName.Substring(0, 8)) And (iFileName = 11) And (sFileExtension.CompareTo("TXT") = 0) Then

                        ReDim Preserve sValidFiles(j)
                        ReDim Preserve sValidFilesPath(j)

                        Dim iNameLength As Integer = sFiles(i).Length
                        sValidFiles(j) = sFiles(i).Substring(iNameLength - 15, 15)
                        sValidFilesPath(j) = sFiles(i).Substring(0, iNameLength - 15)
                        j += 1

                    End If
                Next
            Else
                Console.WriteLine("Path does not exist for Med Express SameDay Manifests:" & sMedExpressSdPath)
                Exit Sub
            End If


            If Not sValidFiles Is Nothing Then

                Console.WriteLine("Med Express SameDay Files Detected. Beginning Import.")

                ' We can use the same SqlConnection for the entire batch of files
                If oConn Is Nothing Then
                    oConn = New SqlConnection(sScanListConnection)
                    oConn.Open()
                End If

                ' Process Batch of Valid Files
                For i = 0 To sValidFiles.Length - 1

                    ' Either the entire file is processed and renamed or the operation on that file fails
                    oTran = oConn.BeginTransaction()
                    oCmd = New SqlCommand("spInsertPuManifestRecordV2", oConn, oTran)
                    oCmd.CommandType = CommandType.StoredProcedure

                    ' Read in the next file
                    oPickupManifest = New PickupManifestMapper(sValidFilesPath(i) & sValidFiles(i), "V2s")

                    If Not oPickupManifest.Records Is Nothing Then

                        bResult = InsertManifestRecords(oPickupManifest.Records, oCmd, oPickupManifest.FileVersion)
                        If bResult Then
                            sTmp = sValidFiles(i).Split(".")
                            sTmp(1) = sTmp(1) & "_ARC"

                            Dim sOldFileName As String = sValidFilesPath(i) & sValidFiles(i)
                            Dim sNewFileName As String = sValidFilesPath(i) & sTmp(0) & "." & sTmp(1)

                            'File.Move(sValidFiles(i), sTmp(0) & "." & sTmp(1))
                            File.Move(sOldFileName, sNewFileName)
                            sTmp = Nothing
                            Console.WriteLine(sValidFiles(i) & " was successfully processed")
                        Else
                            Console.WriteLine(sValidFiles(i) & " was NOT successfully processed")
                        End If

                    Else

                        ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Console.WriteLine("[Module1.1036]" & oPickupManifest.ErrorMessage & ". " & sValidFiles(i) & "failed to process.")
                        bResult = False

                    End If

                    If bResult = True Then
                        oTran.Commit()
                    Else
                        oTran.Rollback()
                    End If

                    oTran = Nothing
                    oCmd = Nothing

                Next

            End If


        Catch ex As Exception

            Console.WriteLine(ex.Message, "ImportMedExpressManifests() Status")
            If Not oTran Is Nothing Then oTran.Rollback()

        Finally

            oCmd = Nothing
            oTran = Nothing

            If Not oConn Is Nothing Then
                oConn.Close()
                oConn.Dispose()
            End If

        End Try


    End Sub

    'Private Sub ImportDSCDeliveryManifests()

    '    Dim bResult As Boolean = False
    '    Dim oConn As SqlConnection = Nothing
    '    Dim oCmd As SqlCommand = Nothing
    '    Dim oTran As SqlTransaction = Nothing
    '    Dim oPickupManifest As PickupManifest = Nothing

    '    Dim sFiles() As String = Nothing
    '    Dim sFolders() As String = Nothing
    '    Dim sValidFiles() As String = Nothing
    '    Dim sValidFilesPath() As String = Nothing

    '    Dim sFileNameParts() As String = Nothing
    '    Dim iParts As Integer = 0

    '    Dim sFileName As String = Nothing
    '    Dim iFileName As Integer = 0
    '    Dim iFirstThree As Integer = 0

    '    Dim sFileExtension As String = Nothing

    '    Dim sTmp() As String = Nothing
    '    Dim i As Int32 = 0
    '    Dim j As Int32 = 0

    '    Try

    '        'Get List of Files that Must be Processed
    '        sDCSDeliveryPath = sDCSDeliveryPath.ToUpper
    '        If Directory.Exists(sDCSDeliveryPath) Then
    '            sFiles = Directory.GetFiles(sDCSDeliveryPath)
    '            For i = 0 To sFiles.Length - 1

    '                sFiles(i) = sFiles(i).ToUpper

    '                sFileNameParts = sFiles(i).Split(".")
    '                iParts = sFileNameParts.Length

    '                sFolders = sFileNameParts(iParts - 2).Split("\")

    '                sFileName = sFolders(sFolders.Length - 1)
    '                sFileExtension = sFileNameParts(iParts - 1)

    '                iFileName = sFileName.Length

    '                If IsNumeric(sFileName.Substring(0, 8)) And (iFileName = 11) And (sFileExtension.CompareTo("TXT") = 0) Then

    '                    ReDim Preserve sValidFiles(j)
    '                    ReDim Preserve sValidFilesPath(j)

    '                    Dim iNameLength As Integer = sFiles(i).Length
    '                    sValidFiles(j) = sFiles(i).Substring(iNameLength - 15, 15)
    '                    sValidFilesPath(j) = sFiles(i).Substring(0, iNameLength - 15)
    '                    j += 1

    '                End If
    '            Next
    '        Else
    '            Console.WriteLine("Path does not exist for DSC Delivery Manifests:" & sDCSDeliveryPath)
    '            Exit Sub
    '        End If


    '        If Not sValidFiles Is Nothing Then

    '            Console.WriteLine("DSC Delivery Files Detected. Beginning Import.")

    '            ' We can use the same SqlConnection for the entire batch of files
    '            If oConn Is Nothing Then
    '                oConn = New SqlConnection(sScanListConnection)
    '                oConn.Open()
    '            End If

    '            ' Process Batch of Valid Files
    '            For i = 0 To sValidFiles.Length - 1

    '                ' Either the entire file is processed and renamed or the operation on that file fails
    '                oTran = oConn.BeginTransaction()
    '                oCmd = New SqlCommand("spInsertPuManifestRecordV2", oConn, oTran)
    '                oCmd.CommandType = CommandType.StoredProcedure

    '                ' Read in the next file
    '                oPickupManifest = New PickupManifestMapper(sValidFilesPath(i) & sValidFiles(i), "V5")

    '                If Not oPickupManifest.Records Is Nothing Then

    '                    bResult = InsertManifestRecords(oPickupManifest.Records, oCmd, oPickupManifest.FileVersion)
    '                    If bResult Then
    '                        sTmp = sValidFiles(i).Split(".")
    '                        sTmp(1) = sTmp(1) & "_ARC"

    '                        Dim sOldFileName As String = sValidFilesPath(i) & sValidFiles(i)
    '                        Dim sNewFileName As String = sValidFilesPath(i) & sTmp(0) & "." & sTmp(1)

    '                        'File.Move(sValidFiles(i), sTmp(0) & "." & sTmp(1))
    '                        File.Move(sOldFileName, sNewFileName)
    '                        sTmp = Nothing
    '                        Console.WriteLine(sValidFiles(i) & " was successfully processed")
    '                    Else
    '                        Console.WriteLine(sValidFiles(i) & " was NOT successfully processed")
    '                    End If

    '                Else

    '                    ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                    Console.WriteLine("[Module1.1181]" & oPickupManifest.ErrorMessage & ". " & sValidFiles(i) & "failed to process.")
    '                    bResult = False

    '                End If

    '                If bResult = True Then
    '                    oTran.Commit()
    '                Else
    '                    oTran.Rollback()
    '                End If

    '                oTran = Nothing
    '                oCmd = Nothing

    '            Next

    '        End If


    '    Catch ex As Exception

    '        Console.WriteLine(ex.Message, "ImportDSCDeliveryManifests() Status")
    '        If Not oTran Is Nothing Then oTran.Rollback()

    '    Finally

    '        oCmd = Nothing
    '        oTran = Nothing

    '        If Not oConn Is Nothing Then
    '            oConn.Close()
    '            oConn.Dispose()
    '        End If

    '    End Try


    'End Sub

    'Private Sub ImportPiaDeliveryManifests()

    '    Dim bResult As Boolean = False
    '    Dim oConn As SqlConnection = Nothing
    '    Dim oCmd As SqlCommand = Nothing
    '    Dim oTran As SqlTransaction = Nothing
    '    Dim oPickupManifest As PickupManifest = Nothing

    '    Dim sFiles() As String = Nothing
    '    Dim sFolders() As String = Nothing
    '    Dim sValidFiles() As String = Nothing
    '    Dim sValidFilesPath() As String = Nothing

    '    Dim sFileNameParts() As String = Nothing
    '    Dim iParts As Integer = 0

    '    Dim sFileName As String = Nothing
    '    Dim iFileName As Integer = 0
    '    Dim iFirstThree As Integer = 0

    '    Dim sFileExtension As String = Nothing

    '    Dim sTmp() As String = Nothing
    '    Dim i As Int32 = 0
    '    Dim j As Int32 = 0

    '    Try

    '        'Get List of Files that Must be Processed
    '        sPiaDeliveryPath = sPiaDeliveryPath.ToUpper
    '        If Directory.Exists(sPiaDeliveryPath) Then
    '            sFiles = Directory.GetFiles(sPiaDeliveryPath)
    '            For i = 0 To sFiles.Length - 1

    '                sFiles(i) = sFiles(i).ToUpper

    '                sFileNameParts = sFiles(i).Split(".")
    '                iParts = sFileNameParts.Length

    '                sFolders = sFileNameParts(iParts - 2).Split("\")

    '                sFileName = sFolders(sFolders.Length - 1)
    '                sFileExtension = sFileNameParts(iParts - 1)

    '                iFileName = sFileName.Length

    '                'If IsNumeric(sFileName.Substring(0, 8)) And (iFileName = 11) And (sFileExtension.CompareTo("TXT") = 0) Then
    '                If (sFileExtension.CompareTo("TXT_HW") = 0) Then

    '                    ReDim Preserve sValidFiles(j)
    '                    ReDim Preserve sValidFilesPath(j)

    '                    Dim iNameLength As Integer = sFiles(i).Length
    '                    sValidFiles(j) = sFiles(i).Substring(iNameLength - 15, 15)
    '                    sValidFilesPath(j) = sFiles(i).Substring(0, iNameLength - 15)
    '                    j += 1

    '                End If
    '            Next
    '        Else
    '            Console.WriteLine("Path does not exist for PIA Delivery Manifests:" & sPiaDeliveryPath)
    '            Exit Sub
    '        End If


    '        If Not sValidFiles Is Nothing Then

    '            Console.WriteLine("PIA Delivery Files Detected. Beginning Import.")

    '            ' We can use the same SqlConnection for the entire batch of files
    '            If oConn Is Nothing Then
    '                oConn = New SqlConnection(sScanListConnection)
    '                oConn.Open()
    '            End If

    '            ' Process Batch of Valid Files
    '            For i = 0 To sValidFiles.Length - 1

    '                ' Either the entire file is processed and renamed or the operation on that file fails
    '                oTran = oConn.BeginTransaction()
    '                'oCmd = New SqlCommand("spInsertPuManifestRecordV2", oConn, oTran)
    '                'oCmd = New SqlCommand("spInsertPuManifestRecordV4", oConn, oTran)
    '                oCmd = New SqlCommand("spInsertPuManifestRecordV3", oConn, oTran)
    '                oCmd.CommandType = CommandType.StoredProcedure

    '                ' Read in the next file
    '                oPickupManifest = New PickupManifestMapper(sValidFilesPath(i) & sValidFiles(i), "V6")

    '                If Not oPickupManifest.Records Is Nothing Then

    '                    bResult = InsertManifestRecords(oPickupManifest.Records, oCmd, oPickupManifest.FileVersion)
    '                    If bResult Then
    '                        sTmp = sValidFiles(i).Split(".")
    '                        sTmp(1) = sTmp(1) & "_ARC"

    '                        Dim sOldFileName As String = sValidFilesPath(i) & sValidFiles(i)
    '                        Dim sNewFileName As String = sValidFilesPath(i) & sTmp(0) & "." & sTmp(1)

    '                        'File.Move(sValidFiles(i), sTmp(0) & "." & sTmp(1))
    '                        File.Move(sOldFileName, sNewFileName)
    '                        sTmp = Nothing
    '                        Console.WriteLine(sValidFiles(i) & " was successfully processed")
    '                    Else
    '                        Console.WriteLine(sValidFiles(i) & " was NOT successfully processed")
    '                    End If

    '                Else

    '                    ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                    Console.WriteLine("[Module1.1181]" & oPickupManifest.ErrorMessage & ". " & sValidFiles(i) & "failed to process.")
    '                    bResult = False

    '                End If

    '                If bResult = True Then
    '                    oTran.Commit()
    '                    'oTran.Rollback() 'DEBUG
    '                Else
    '                    oTran.Rollback()
    '                End If

    '                oTran = Nothing
    '                oCmd = Nothing

    '            Next

    '        End If


    '    Catch ex As Exception

    '        Console.WriteLine(ex.Message, "ImportPiaDeliveryManifests() Status")
    '        If Not oTran Is Nothing Then oTran.Rollback()

    '    Finally

    '        oCmd = Nothing
    '        oTran = Nothing

    '        If Not oConn Is Nothing Then
    '            oConn.Close()
    '            oConn.Dispose()
    '        End If

    '    End Try


    'End Sub

    'Private Sub ImportProCourierManifests()

    '    Dim bResult As Boolean = False
    '    Dim oConn As SqlConnection = Nothing
    '    Dim oCmd As SqlCommand = Nothing
    '    Dim oTran As SqlTransaction = Nothing
    '    Dim oPickupManifest As PickupManifest = Nothing

    '    Dim sFiles() As String = Nothing
    '    Dim sFolders() As String = Nothing
    '    Dim sValidFiles() As String = Nothing
    '    Dim sValidFilesPath() As String = Nothing

    '    Dim sFileNameParts() As String = Nothing
    '    Dim iParts As Integer = 0

    '    Dim sFileName As String = Nothing
    '    Dim iFileName As Integer = 0
    '    Dim iFirstThree As Integer = 0

    '    Dim sFileExtension As String = Nothing

    '    Dim sTmp() As String = Nothing
    '    Dim i As Int32 = 0
    '    Dim j As Int32 = 0

    '    Try

    '        'Get List of Files that Must be Processed
    '        sProCourierPath = sProCourierPath.ToUpper
    '        If Directory.Exists(sProCourierPath) Then
    '            sFiles = Directory.GetFiles(sProCourierPath)
    '            For i = 0 To sFiles.Length - 1

    '                sFiles(i) = sFiles(i).ToUpper

    '                sFileNameParts = sFiles(i).Split(".")
    '                iParts = sFileNameParts.Length

    '                sFolders = sFileNameParts(iParts - 2).Split("\")

    '                sFileName = sFolders(sFolders.Length - 1)
    '                sFileExtension = sFileNameParts(iParts - 1)

    '                iFileName = sFileName.Length

    '                If IsNumeric(sFileName.Substring(0, 8)) And (iFileName = 11) And (sFileExtension.CompareTo("TXT") = 0) Then

    '                    ReDim Preserve sValidFiles(j)
    '                    ReDim Preserve sValidFilesPath(j)

    '                    Dim iNameLength As Integer = sFiles(i).Length
    '                    sValidFiles(j) = sFiles(i).Substring(iNameLength - 15, 15)
    '                    sValidFilesPath(j) = sFiles(i).Substring(0, iNameLength - 15)
    '                    j += 1

    '                End If
    '            Next
    '        Else
    '            Console.WriteLine("Path does not exist for ProCourier Manifests:" & sProCourierPath)
    '            Exit Sub
    '        End If


    '        If Not sValidFiles Is Nothing Then

    '            Console.WriteLine("ProCourier Delivery Files Detected. Beginning Import.")

    '            ' We can use the same SqlConnection for the entire batch of files
    '            If oConn Is Nothing Then
    '                oConn = New SqlConnection(sScanListConnection)
    '                oConn.Open()
    '            End If

    '            ' Process Batch of Valid Files
    '            For i = 0 To sValidFiles.Length - 1

    '                ' Either the entire file is processed and renamed or the operation on that file fails
    '                oTran = oConn.BeginTransaction()
    '                oCmd = New SqlCommand("spInsertPuManifestRecordV5", oConn, oTran)
    '                oCmd.CommandType = CommandType.StoredProcedure

    '                ' Read in the next file
    '                oPickupManifest = New PickupManifestMapper(sValidFilesPath(i) & sValidFiles(i), "V7")

    '                If Not oPickupManifest.Records Is Nothing Then

    '                    bResult = InsertManifestRecords(oPickupManifest.Records, oCmd, oPickupManifest.FileVersion)
    '                    If bResult Then
    '                        sTmp = sValidFiles(i).Split(".")
    '                        sTmp(1) = sTmp(1) & "_ARC"

    '                        Dim sOldFileName As String = sValidFilesPath(i) & sValidFiles(i)
    '                        Dim sNewFileName As String = sValidFilesPath(i) & sTmp(0) & "." & sTmp(1)

    '                        'File.Move(sValidFiles(i), sTmp(0) & "." & sTmp(1))
    '                        File.Move(sOldFileName, sNewFileName)
    '                        sTmp = Nothing
    '                        Console.WriteLine(sValidFiles(i) & " was successfully processed")
    '                    Else
    '                        Console.WriteLine(sValidFiles(i) & " was NOT successfully processed")
    '                    End If

    '                Else

    '                    ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                    Console.WriteLine("[Module1.891]" & oPickupManifest.ErrorMessage & ". " & sValidFiles(i) & "failed to process.")
    '                    bResult = False

    '                End If

    '                If bResult = True Then
    '                    oTran.Commit()
    '                Else
    '                    oTran.Rollback()
    '                End If

    '                oTran = Nothing
    '                oCmd = Nothing

    '            Next

    '        End If


    '    Catch ex As Exception

    '        Console.WriteLine(ex.Message, "ImportProCourierDeliveryManifests() Status")
    '        If Not oTran Is Nothing Then oTran.Rollback()

    '    Finally

    '        oCmd = Nothing
    '        oTran = Nothing

    '        If Not oConn Is Nothing Then
    '            oConn.Close()
    '            oConn.Dispose()
    '        End If

    '    End Try


    'End Sub

    Private Sub ImportDIO2CSManifests()

        Dim bResult As Boolean = False
        Dim oConn As SqlConnection = Nothing
        Dim oCmd As SqlCommand = Nothing
        Dim oTran As SqlTransaction = Nothing
        Dim oPickupManifest As PickupManifest = Nothing

        Dim sFiles() As String = Nothing
        Dim sFolders() As String = Nothing
        Dim sValidFiles() As String = Nothing
        Dim sValidFilesPath() As String = Nothing

        Dim sFileNameParts() As String = Nothing
        Dim iParts As Integer = 0

        Dim sFileName As String = Nothing
        Dim iFileName As Integer = 0
        Dim iFirstThree As Integer = 0

        Dim sFileExtension As String = Nothing

        Dim sTmp() As String = Nothing
        Dim i As Int32 = 0
        Dim j As Int32 = 0

        Try

            'Get List of Files that Must be Processed
            sDio2CsPath = sDio2CsPath.ToUpper

            If Directory.Exists(sDio2CsPath) Then

                sFiles = Directory.GetFiles(sDio2CsPath)
                For i = 0 To sFiles.Length - 1

                    sFiles(i) = sFiles(i).ToUpper

                    sFileNameParts = sFiles(i).Split(".")
                    iParts = sFileNameParts.Length

                    sFolders = sFileNameParts(iParts - 2).Split("\")

                    sFileName = sFolders(sFolders.Length - 1)
                    sFileExtension = sFileNameParts(iParts - 1)

                    iFileName = sFileName.Length

                    If IsNumeric(sFileName.Substring(0, 8)) And (iFileName = 11) And (sFileExtension.CompareTo("TXT") = 0) Then

                        ReDim Preserve sValidFiles(j)
                        ReDim Preserve sValidFilesPath(j)

                        Dim iNameLength As Integer = sFiles(i).Length
                        sValidFiles(j) = sFiles(i).Substring(iNameLength - 15, 15)
                        sValidFilesPath(j) = sFiles(i).Substring(0, iNameLength - 15)
                        j += 1

                    End If
                Next

            Else

                Console.WriteLine("Path does not exist for DIO to CS Manifests:" & sDio2CsPath)
                Exit Sub

            End If


            If Not sValidFiles Is Nothing Then

                Console.WriteLine("DIO to CS Delivery Files Detected. Beginning Import.")

                ' We can use the same SqlConnection for the entire batch of files
                If oConn Is Nothing Then
                    oConn = New SqlConnection(sScanListConnection)
                    oConn.Open()
                End If

                ' Process Batch of Valid Files
                For i = 0 To sValidFiles.Length - 1

                    ' Either the entire file is processed and renamed or the operation on that file fails
                    oTran = oConn.BeginTransaction()
                    oCmd = New SqlCommand("spInsertPuManifestRecordV8", oConn, oTran)
                    oCmd.CommandType = CommandType.StoredProcedure

                    ' Read in the next file
                    oPickupManifest = New PickupManifestMapper(sValidFilesPath(i) & sValidFiles(i), "V8")

                    If Not oPickupManifest.Records Is Nothing Then

                        bResult = InsertManifestRecords(oPickupManifest.Records, oCmd, oPickupManifest.FileVersion)
                        If bResult Then
                            sTmp = sValidFiles(i).Split(".")
                            sTmp(1) = sTmp(1) & "_ARC"

                            Dim sOldFileName As String = sValidFilesPath(i) & sValidFiles(i)
                            Dim sNewFileName As String = sValidFilesPath(i) & sTmp(0) & "." & sTmp(1)

                            'File.Move(sValidFiles(i), sTmp(0) & "." & sTmp(1))
                            File.Move(sOldFileName, sNewFileName)
                            sTmp = Nothing
                            Console.WriteLine(sValidFiles(i) & " was successfully processed")
                        Else
                            Console.WriteLine(sValidFiles(i) & " was NOT successfully processed")
                        End If

                    Else

                        ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Console.WriteLine("[Module1.891]" & oPickupManifest.ErrorMessage & ". " & sValidFiles(i) & "failed to process.")
                        bResult = False

                    End If

                    If bResult = True Then
                        oTran.Commit()
                    Else
                        oTran.Rollback()
                    End If

                    oTran = Nothing
                    oCmd = Nothing

                Next

            End If


        Catch ex As Exception

            Console.WriteLine(ex.Message, "ImportProCourierDeliveryManifests() Status")
            If Not oTran Is Nothing Then oTran.Rollback()

        Finally

            oCmd = Nothing
            oTran = Nothing

            If Not oConn Is Nothing Then
                oConn.Close()
                oConn.Dispose()
            End If

        End Try


    End Sub

    Private Function InsertManifestRecords(ByRef p_oPickupManifest As TTSI.BARCODES.PickupManifestRecords, ByRef p_oCmd As SqlCommand, ByRef p_sFileVersion As String) As Boolean

        Dim bReturn As Boolean = True

        Dim oParam As SqlParameter = Nothing

        Dim sTrackingNum As String
        Dim sOrderId As String
        Dim sFromCustId As String
        Dim sFromCustName As String
        Dim sFromLocId As String
        Dim sFromLocName As String
        Dim sFromLocStreet As String
        Dim sFromLocAddress2 As String
        Dim sFromLocCity As String
        Dim sFromLocState As String
        Dim sFromLocZip As String
        Dim sFromLocPhone As String
        Dim sFromLocContact As String
        Dim sFromLocEmail As String
        Dim sToCustId As String
        Dim sToCustName As String
        Dim sToLocId As String
        Dim sToLocName As String
        Dim sToLocStreet As String
        Dim sToLocAddress2 As String
        Dim sToLocCity As String
        Dim sToLocState As String
        Dim sToLocZip As String
        Dim sToLocContact As String
        Dim sToLocPhone As String
        Dim sToLocEmail As String
        Dim fWeight As Decimal
        Dim sPieces As String
        Dim sSentByName As String
        Dim sCartonCode As String
        Dim sDimensions As String
        Dim sServiceLevel As String
        Dim sBillType As String
        Dim sBillNum As String
        Dim dtTranDate As Date
        Dim sVoid As String
        Dim sReferenceNumber As String
        Dim sPONumber As String
        Dim sThirdPartyBillNum As String
        Dim sModifiers As String
        Dim fDeclaredValue As Decimal
        Dim sUniqueRowId As String

        Dim iResult As Integer = 0

        Dim iTrackingNum As Integer = 25
        Dim iOrderId As Integer = 40
        Dim iFromCustId As Integer = 10
        Dim iFromCustName As Integer = 70
        Dim iFromLocId As Integer = 10
        Dim iFromLocName As Integer = 70
        Dim iFromLocStreet As Integer = 40
        Dim iFromLocAddress2 As Integer = 30
        Dim iFromLocCity As Integer = 32
        Dim iFromLocState As Integer = 2
        Dim iFromLocZip As Integer = 10
        Dim iFromLocPhone As Integer = 20
        Dim iFromLocContact As Integer = 40
        Dim iFromLocEmail As Integer = 30
        Dim iToCustId As Integer = 10
        Dim iToCustName As Integer = 70
        Dim iToLocId As Integer = 10
        Dim iToLocName As Integer = 70
        Dim iToLocStreet As Integer = 40
        Dim iToLocAddress2 As Integer = 30
        Dim iToLocCity As Integer = 32
        Dim iToLocState As Integer = 2
        Dim iToLocZip As Integer = 10
        Dim iToLocContact As Integer = 40
        Dim iToLocPhone As Integer = 20
        Dim iToLocEmail As Integer = 30
        'Dim iWeight As Integer = xx
        Dim iPieces As Integer = 10
        Dim iSentByName As Integer = 30
        Dim iCartonCode As Integer = 20
        Dim iDimensions As Integer = 20
        Dim iServiceLevel As Integer = 20
        Dim iBillType As Integer = 20
        Dim iBillNum As Integer = 50
        'Dim iTranDate As Integer = xx
        Dim iVoid As Integer = 1
        Dim iReferenceNumber As Integer = 40
        Dim iPONumber As Integer = 40
        Dim iThirdPartyBillNum As Integer = 40
        Dim iModifiers As Integer = 122
        'Dim iDeclaredValue As Integer = xx
        Dim iUniqueRowId As Integer = 29

        Dim iLineNumber As Integer = 0


        For Each oRec As PickupManifestRecord In p_oPickupManifest

            iLineNumber += 1

            With oRec

                ' LineNumber
                .LineNumber = iLineNumber

                ' TrackingNum
                sTrackingNum = .TrackingNumber
                AddInputParam_VarChar(p_oCmd, "@TrackingNum", iTrackingNum, sTrackingNum)
                ' OrderId
                sOrderId = .OrderId
                AddInputParam_VarChar(p_oCmd, "@OrderId", iOrderId, sOrderId)
                ' FromCustId
                sFromCustId = .FromCustId
                AddInputParam_VarChar(p_oCmd, "@FromCustId", iFromCustId, sFromCustId)
                ' FromCustName
                sFromCustName = .FromCustName
                AddInputParam_VarChar(p_oCmd, "@FromCustName", iFromCustName, sFromCustName)
                ' FromLocId
                sFromLocId = .FromLocId
                AddInputParam_VarChar(p_oCmd, "@FromLocId", iFromLocId, sFromLocId)
                ' FromLocName
                sFromLocName = .FromLocName
                AddInputParam_VarChar(p_oCmd, "@FromLocName", iFromLocName, sFromLocName)
                ' FromLocStreet
                sFromLocStreet = .FromLocStreet
                AddInputParam_VarChar(p_oCmd, "@FromLocStreet", iFromLocStreet, sFromLocStreet)
                ' FromLocAddress2
                sFromLocAddress2 = .FromLocAddress2
                AddInputParam_VarChar(p_oCmd, "@FromLocAddress2", iFromLocAddress2, sFromLocAddress2)
                ' FromLocCity
                sFromLocCity = .FromLocCity
                AddInputParam_VarChar(p_oCmd, "@FromLocCity", iFromLocCity, sFromLocCity)
                ' FromLocState
                sFromLocState = .FromLocState
                AddInputParam_VarChar(p_oCmd, "@FromLocState", iFromLocState, sFromLocState)
                ' FromLocZip
                sFromLocZip = .FromLocZip
                AddInputParam_VarChar(p_oCmd, "@FromLocZip", iFromLocZip, sFromLocZip)
                ' FromLocPhone
                sFromLocPhone = .FromLocPhone
                AddInputParam_VarChar(p_oCmd, "@FromLocPhone", iFromLocPhone, sFromLocPhone)
                ' FromLocContact
                sFromLocContact = .FromLocContact
                AddInputParam_VarChar(p_oCmd, "@FromLocContact", iFromLocContact, sFromLocContact)
                ' FromLocEmail
                sFromLocEmail = .FromLocEmail
                AddInputParam_VarChar(p_oCmd, "@FromLocEmail", iFromLocEmail, sFromLocEmail)
                ' ToCustId
                sToCustId = .ToCustId
                AddInputParam_VarChar(p_oCmd, "@ToCustId", iToCustId, sToCustId)
                ' ToCustName
                sToCustName = .ToCustName
                AddInputParam_VarChar(p_oCmd, "@ToCustName", iToCustName, sToCustName)
                ' ToLocId
                sToLocId = .ToLocId
                AddInputParam_VarChar(p_oCmd, "@ToLocId", iToLocId, sToLocId)
                ' ToLocName
                sToLocName = .ToLocName
                AddInputParam_VarChar(p_oCmd, "@ToLocName", iToLocName, sToLocName)
                ' ToLocStreet
                sToLocStreet = .ToLocStreet
                AddInputParam_VarChar(p_oCmd, "@ToLocStreet", iToLocStreet, sToLocStreet)
                ' ToLocAddress2
                sToLocAddress2 = .ToLocAddress2
                AddInputParam_VarChar(p_oCmd, "@ToLocAddress2", iToLocAddress2, sToLocAddress2)
                ' ToLocCity
                sToLocCity = .ToLocCity
                AddInputParam_VarChar(p_oCmd, "@ToLocCity", iToLocCity, sToLocCity)
                ' ToLocState
                sToLocState = .ToLocState
                AddInputParam_VarChar(p_oCmd, "@ToLocState", iToLocState, sToLocState)
                ' ToLocZip
                sToLocZip = .ToLocZip
                AddInputParam_VarChar(p_oCmd, "@ToLocZip", iToLocZip, sToLocZip)
                ' ToLocContact
                sToLocContact = .ToLocContact
                AddInputParam_VarChar(p_oCmd, "@ToLocContact", iToLocContact, sToLocContact)
                ' ToLocPhone
                sToLocPhone = .ToLocPhone
                AddInputParam_VarChar(p_oCmd, "@ToLocPhone", iToLocPhone, sToLocPhone)
                ' ToLocEmail
                sToLocEmail = .ToLocEmail
                AddInputParam_VarChar(p_oCmd, "@ToLocEmail", iToLocEmail, sToLocEmail)
                ' Weight
                fWeight = .Weight
                AddInputParam_Decimal(p_oCmd, "@Weight", fWeight)
                ' Pieces
                sPieces = .Pieces
                AddInputParam_VarChar(p_oCmd, "@Pieces", iPieces, sPieces)
                ' SentByName
                sSentByName = .SentByName
                AddInputParam_VarChar(p_oCmd, "@SentByName", iSentByName, sSentByName)
                ' CartonCode
                sCartonCode = .CartonCode
                AddInputParam_VarChar(p_oCmd, "@CartonCode", iCartonCode, sCartonCode)
                ' Dimensions
                sDimensions = .Dimensions
                AddInputParam_VarChar(p_oCmd, "@Dimensions", iDimensions, sDimensions)
                ' ServiceLevel
                sServiceLevel = .ServiceLevel
                AddInputParam_VarChar(p_oCmd, "@ServiceLevel", iServiceLevel, sServiceLevel)
                ' BillType
                sBillType = .BillType
                AddInputParam_VarChar(p_oCmd, "@BillType", iBillType, sBillType)
                ' BillNum
                sBillNum = .BillNum
                AddInputParam_VarChar(p_oCmd, "@BillNum", iBillNum, sBillNum)
                ' TranDate
                dtTranDate = .TranDate
                AddInputParam_DateTime(p_oCmd, "@TranDate", dtTranDate)
                ' Void
                sVoid = .Void
                AddInputParam_VarChar(p_oCmd, "@Void", iVoid, sVoid)
                ' ReferenceNumber
                sReferenceNumber = .ReferenceNumber
                AddInputParam_VarChar(p_oCmd, "@ReferenceNumber", iReferenceNumber, sReferenceNumber)
                ' PONumber
                sPONumber = .PONumber
                AddInputParam_VarChar(p_oCmd, "@PONumber", iPONumber, sPONumber)
                ' ThirdPartyBillNum
                sThirdPartyBillNum = .ThirdPartyBillNum
                AddInputParam_VarChar(p_oCmd, "@ThirdPartyBillNum", iThirdPartyBillNum, sThirdPartyBillNum)
                ' Modifiers
                sModifiers = .Modifiers
                AddInputParam_VarChar(p_oCmd, "@Modifiers", iModifiers, sModifiers)
                ' DeclaredValue
                fDeclaredValue = .DeclaredValue
                AddInputParam_Decimal(p_oCmd, "@DeclaredValue", fDeclaredValue)
                ' UniqueRowId (Specialized Case:  Only necessary when TrackingNumber is not unique i.e. MedEx)
                If (String.Compare(p_sFileVersion, "V2") = 0) _
                    Or (String.Compare(p_sFileVersion, "V3") = 0) _
                    Or (String.Compare(p_sFileVersion, "V4") = 0) _
                    Or (String.Compare(p_sFileVersion, "V5") = 0) _
                    Or (String.Compare(p_sFileVersion, "V6") = 0) _
                    Or (String.Compare(p_sFileVersion, "V7") = 0) _
                    Or (String.Compare(p_sFileVersion, "V8") = 0) _
                    Then
                    sUniqueRowId = .UniqueRecordId
                    AddInputParam_VarChar(p_oCmd, "@UniqueRowId", iUniqueRowId, sUniqueRowId)
                End If

                'sFromLocName = .FromLocName

            End With

            iResult = p_oCmd.ExecuteNonQuery()

            ' Analyze Stored Procedure Result for errors
            If (String.Compare(p_sFileVersion, "V0") = 0) Then
                'v1.0 of spInsertPuManifestRecordV0 updates as follows...
                '   8 rows if FromLoc does not exist and ToLoc does not exist
                '   7 rows if FromLoc does exist but ToLoc does not exist
                '   5 rows if FromLoc exists and ToLoc exists
                If (iResult = 5 Or iResult = 7 Or iResult = 8) Then
                    p_oCmd.Parameters.Clear()
                Else
                    bReturn = False
                    Exit For
                End If
            ElseIf (String.Compare(p_sFileVersion, "V1") = 0) _
                    Or (String.Compare(p_sFileVersion, "V6") = 0) Then
                'v1.0 of spInsertPuManifestRecordV1 updates as follows...
                '   7 rows if FromLoc does not exist and ToLoc does not exist
                '   6 rows if FromLoc does exist but ToLoc does not exist
                '   5 rows if FromLoc exists and ToLoc exists
                If (iResult = 4 Or iResult = 5 Or iResult = 6) Then
                    p_oCmd.Parameters.Clear()
                Else
                    bReturn = False
                    Exit For
                End If
            ElseIf (String.Compare(p_sFileVersion, "V2") = 0) _
                    Or (String.Compare(p_sFileVersion, "V3") = 0) _
                    Or (String.Compare(p_sFileVersion, "V4") = 0) _
                    Or (String.Compare(p_sFileVersion, "V5") = 0) _
                    Then
                'v1.0 of spInsertPuManifestRecordV2 updates as follows...
                '   8 rows if FromLoc does not exist and ToLoc does not exist
                '   7 rows if FromLoc does exist but ToLoc does not exist
                '   5 rows if FromLoc exists and ToLoc exists
                If (iResult = 5 Or iResult = 7 Or iResult = 8) Then
                    p_oCmd.Parameters.Clear()
                Else
                    bReturn = False
                    Exit For
                End If
            ElseIf (String.Compare(p_sFileVersion, "V7") = 0) _
                    Then
                'v1.0 of spInsertPuManifestRecordV2 updates as follows...
                '   10 rows if FromLoc does not exist and ToLoc does not exist
                '    9 rows if FromLoc does not exist and ToLoc does not exist and ToLoc has no default route specified for that zip
                '    8 rows if FromLoc does exist but ToLoc does not exist
                '    7 rows if FromLoc does exist but ToLoc does not exist and ToLoc has no default route specified for that zip
                '    7 rows if FromLoc does not exist but ToLoc does exist
                '    5 rows if FromLoc does exist and ToLoc does exists
                If (iResult = 5 Or iResult = 7 Or iResult = 8 Or iResult = 9 Or iResult = 10) Then
                    p_oCmd.Parameters.Clear()
                Else
                    bReturn = False
                    Exit For
                End If
            ElseIf (String.Compare(p_sFileVersion, "V8") = 0) Then
                'v1.0 of spInsertPuManifestRecordV2 updates as follows...
                '   10 rows if FromLoc does not exist and ToLoc does not exist
                '    9 rows if FromLoc does not exist and ToLoc does not exist and ToLoc has no default route specified for that zip
                '    8 rows if FromLoc does exist but ToLoc does not exist
                '    7 rows if FromLoc does exist but ToLoc does not exist and ToLoc has no default route specified for that zip
                '    7 rows if FromLoc does not exist but ToLoc does exist
                '    5 rows if FromLoc does exist and ToLoc does exists
                If (iResult = 5 Or iResult = 7 Or iResult = 8 Or iResult = 9 Or iResult = 10) Then
                    p_oCmd.Parameters.Clear()
                Else
                    bReturn = False
                    Exit For
                End If
            Else
                bReturn = False
                Exit For
            End If

            ' DEBUG CODE
            Dim sb As New StringBuilder
            sb.Append("Tracking Number = ")
            sb.Append(oRec.TrackingNumber)
            sb.Append(", Void = ")
            sb.Append(oRec.Void)
            sb.Append(", UniqueRowId = ")
            sb.Append(oRec.UniqueRecordId)
            sb.Append(", FromState = ")
            sb.Append(oRec.FromLocState)
            sb.Append(", FromCity = ")
            sb.Append(oRec.FromLocCity)
            sb.Append(", Dimensions = ")
            sb.Append(oRec.Dimensions)
            sb.Append(", ToLocId = ")
            sb.Append(oRec.ToLocId)

            Console.WriteLine(sb.ToString())

        Next

        Return bReturn
        'Return False

    End Function

    Private Sub AddInputParam_VarChar(ByRef p_oCmd As SqlCommand, ByVal p_sColName As String, ByVal p_iColLength As Integer, ByVal p_sValue As String)

        Dim oParam As New SqlParameter

        oParam = p_oCmd.Parameters.Add(p_sColName, SqlDbType.VarChar, p_iColLength)
        oParam.Direction = ParameterDirection.Input
        oParam.SqlDbType = SqlDbType.VarChar
        oParam.Value = p_sValue

    End Sub

    Private Sub AddInputParam_DateTime(ByRef p_oCmd As SqlCommand, ByVal p_sColName As String, ByVal p_dtValue As DateTime)

        Dim oParam As New SqlParameter

        oParam = p_oCmd.Parameters.Add(p_sColName, SqlDbType.DateTime)
        oParam.Direction = ParameterDirection.Input
        oParam.SqlDbType = SqlDbType.DateTime
        oParam.Value = p_dtValue

    End Sub

    Private Sub AddInputParam_Decimal(ByRef p_oCmd As SqlCommand, ByVal p_sColName As String, ByVal p_fValue As Decimal)

        Dim oParam As New SqlParameter

        oParam = p_oCmd.Parameters.Add(p_sColName, SqlDbType.Decimal)
        oParam.Direction = ParameterDirection.Input
        oParam.SqlDbType = SqlDbType.Decimal
        oParam.Value = p_fValue

    End Sub

    Private Function ImportSmsData() As Boolean

        ' local variables
        Dim bReturnValue As Boolean = True
        Dim bUpdateVisual As Boolean = False
        Dim sb As New StringBuilder
        Dim sCmd As String
        Dim oDataAdapter As SqlDataAdapter = Nothing
        Dim oDataSet As DataSet = Nothing

        Try
            ' Instantiate Connections
            cnSqlQuery = New SqlConnection(sScanListConnection)

            ' Instantiate Commands
            'cmd = New SqlCommand("SELECT count(*) as PendingRecords FROM UN_TRACKING.dbo.fSMS_ReadyToImport()", cnSqlQuery)
            cmd = New SqlCommand("SELECT count(*) as PendingRecords FROM UN_TRACKING.dbo.SMS_Enhanced", cnSqlQuery)

            ' Set Command Types

            'open the connection
            If Not cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Open()
            'cmd.CommandTimeout = 120

            'populate the DataReader
            drReader = cmd.ExecuteReader

            'Provide Visual Feedback
            If drReader.Read Then

                If CInt(drReader("PendingRecords")) > 0 Then

                    drReader.Close()

                    Console.WriteLine(" Pending SMSList records detected.")
                    Console.Write("Processing ")
                    bUpdateVisual = True

                    ''TO DO:  Execute queries to compensate for unresolved bugs until those bugs are fixed.
                    'Dim bSuccess As Boolean = False

                    'sb.Length = 0
                    'sb.Append("update un_tracking.dbo.scanlist set barcode = upper(barcode)")
                    'bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    'sb.Length = 0
                    'sb.Append("update un_tracking.dbo.scanlist set OperatorID = 'E0000001' where OperatorID = '0'")
                    'bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    'sb.Length = 0
                    'sb.Append("update un_tracking.dbo.scanlist set PointID = 'P0000497' where PointID = '0'")
                    'bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    'sb.Length = 0
                    'sb.Append("update un_tracking.dbo.scanlist set BatchId = BatchId - 10000 where processed = 0 and batchid >= 10000")
                    'bSuccess = ExecuteQuery(sb.ToString, cmd)


                    ' Get Data from ScanList table
                    sb.Length = 0
                    sb.Append("select RowId, 'TR' + '|' ")
                    sb.Append(" + ISNULL(OperatorID,'E0000000') + '|' ")
                    sb.Append(" + ISNULL(PointId,'P0000000') + '|' ")
                    sb.Append(" + ISNULL(rtrim(TRNUM),'') + '|' ")
                    sb.Append(" + '0' + '|' ")
                    sb.Append(" + '1'  + '|' ")
                    sb.Append(" + '0' + '|' ")
                    sb.Append(" + '0' + '|' ")
                    sb.Append(" + case when datepart(month,sent_dt) < 10 then '0' ")
                    sb.Append(" + cast(datepart(month,sent_dt) as varchar) else cast(datepart(month,sent_dt) as varchar) end ")
                    sb.Append(" + case when datepart(day,sent_dt) < 10 then '0' ")
                    sb.Append(" + cast(datepart(day,sent_dt) as varchar) else cast(datepart(day,sent_dt) as varchar) end ")
                    sb.Append(" + cast(datepart(year,sent_dt) as varchar) + case when datepart(hour,sent_dt) < 10 then '0' ")
                    sb.Append(" + cast(datepart(hour,sent_dt) as varchar) else cast(datepart(hour,sent_dt) as varchar) end ")
                    sb.Append(" + case when datepart(minute,sent_dt) < 10 then '0' ")
                    sb.Append(" + cast(datepart(minute,sent_dt) as varchar) else cast(datepart(minute,sent_dt) as varchar) end ")
                    sb.Append(" + case when datepart(second,sent_dt) < 10 then '0' ")
                    sb.Append(" + cast(datepart(second,sent_dt) as varchar) else cast(datepart(second,sent_dt) as varchar) end + '|' ")
                    sb.Append(" + case when datepart(month,sent_dt) < 10 then '0' ")
                    sb.Append(" + cast(datepart(month,sent_dt) as varchar) else cast(datepart(month,sent_dt) as varchar) end ")
                    sb.Append(" + case when datepart(day,sent_dt) < 10 then '0' ")
                    sb.Append(" + cast(datepart(day,sent_dt) as varchar) else cast(datepart(day,sent_dt) as varchar) end ")
                    sb.Append(" + cast(datepart(year,sent_dt) as varchar) + case when datepart(hour,sent_dt) < 10 then '0' ")
                    sb.Append(" + cast(datepart(hour,sent_dt) as varchar) else cast(datepart(hour,sent_dt) as varchar) end ")
                    sb.Append(" + case when datepart(minute,sent_dt) < 10 then '0' ")
                    sb.Append(" + cast(datepart(minute,sent_dt) as varchar) else cast(datepart(minute,sent_dt) as varchar) end ")
                    sb.Append(" + case when datepart(second,sent_dt) < 10 then '0' ")
                    sb.Append(" + cast(datepart(second,sent_dt) as varchar) else cast(datepart(second,sent_dt) as varchar) end ")
                    sb.Append(" + '|' ")
                    sb.Append(" + '0' + '|' ")
                    sb.Append(" + 'CELL' + '|' ")
                    sb.Append(" + RTRIM(ToAddId) ")
                    sb.Append(" as RecordString from un_tracking.dbo.fSMS_ReadyToImport() ")

                    sCmd = sb.ToString()

                    PopulateDataset2(oDataAdapter, oDataSet, sCmd)

                    If Not oDataSet Is Nothing Then

                        If oDataSet.Tables.Count = 1 Then

                            If oDataSet.Tables(0).Rows.Count > 0 Then

                                Dim iRowId As Integer = oDataSet.Tables(0).Rows(0).Item("RowId")
                                Dim oSmsList As New SmsList(oDataSet)
                                Dim i As Integer = -1

                                If Not oSmsList.Records Is Nothing Then

                                    For Each oRec As ScanRecord In oSmsList.Records

                                        Console.Write(".")

                                        i = i + 1
                                        iRowId = oDataSet.Tables(0).Rows(i).Item("RowId")

                                        If Not ImportScanListRecord(oRec) Then

                                            Console.WriteLine("[Module.1686]" & oRec.ErrorMessage)
                                            'If MessageBox.Show(oRec.ErrorMessage + ". Do you want to continue?", "Error Importing Record at RowID " & iRowId.ToString(), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                                            '    bReturnValue = False
                                            '    Exit For
                                            'End If
                                        Else

                                            sb.Length = 0
                                            sb.Append("UPDATE un_tracking.dbo.sms_in SET Processed = 1 WHERE RowId = " & iRowId)
                                            sCmd = sb.ToString()
                                            bReturnValue = ExecuteQuery(sCmd)

                                            If bReturnValue = False Then
                                                ''MessageBox.Show("ScanList Record at RowId " & iRowId & " was processed properly, but its flag was not updated", "Problem Importing ScanList Record", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                Console.WriteLine("SmsList Record at RowId " & iRowId & " was processed properly, but its flag was not updated", "Problem Importing ScanList Record")
                                            End If

                                        End If

                                    Next

                                Else

                                    ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Console.WriteLine(oSmsList.ErrorMessage, "Import SmsList Status")
                                    bReturnValue = False

                                End If

                            Else

                                ''MessageBox.Show("No Records Founds to Import", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Console.WriteLine("No Records Founds to Import", "Import SmsList Status")
                                bReturnValue = True

                            End If

                        Else

                            ''MessageBox.Show("There were no records to import", "Import ScanList Status", MessageBoxButtons.OK)
                            Console.WriteLine("There were no records to import", "Import SmsList Status")
                            bReturnValue = True

                        End If

                    Else

                        ''MessageBox.Show("Database Error", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Console.WriteLine("Database Error", "Import SmsList Status")
                        bReturnValue = False

                    End If

                End If

            End If

            If bUpdateVisual = True Then

                Console.WriteLine("Done!")

            End If

        Catch ex As Exception

            'Error Handling Code Goes Here
            bReturnValue = False
            Console.WriteLine("[Module1.1753]" & ex.Message)

        Finally

            'clean up code that need to run no matter what

            'close the data reader
            drReader.Close()

            If Not cnSqlQuery Is Nothing Then
                If cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Close()
                cnSqlQuery.Dispose()
            End If

            If Not cnSqlInsert Is Nothing Then
                If cnSqlInsert.State = ConnectionState.Open Then cnSqlInsert.Close()
                cnSqlInsert.Dispose()
            End If

            If Not cnSqlUpdate Is Nothing Then
                If cnSqlUpdate.State = ConnectionState.Open Then cnSqlUpdate.Close()
                cnSqlUpdate.Dispose()
            End If

            If Not cnSqlDelete Is Nothing Then
                If cnSqlDelete.State = ConnectionState.Open Then cnSqlDelete.Close()
                cnSqlDelete.Dispose()
            End If

        End Try

        Return bReturnValue

    End Function

    Private Function ImportScanList() As Boolean

        ' local variables
        Dim bReturnValue As Boolean = True
        Dim bUpdateVisual As Boolean = False
        Dim sb As New StringBuilder
        Dim sCmd As String
        Dim oDataAdapter As SqlDataAdapter = Nothing
        Dim oDataSet As DataSet = Nothing

        Try
            ' Instantiate Connections
            cnSqlQuery = New SqlConnection(sScanListConnection)

            ' Instantiate Commands
            'cmd = New SqlCommand("SELECT count(*) as PendingRecords FROM UN_TRACKING.dbo.ScanList WHERE Processed = 0", cnSqlQuery)
            cmd = New SqlCommand("SELECT count(*) as PendingRecords FROM UN_TRACKING.dbo.ScanList WHERE Processed = 0 and EventCode <> 'EZ'", cnSqlQuery)

            ' Set Command Types

            'open the connection
            If Not cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Open()

            'populate the DataReader
            drReader = cmd.ExecuteReader

            'Provide Visual Feedback
            If drReader.Read Then

                If CInt(drReader("PendingRecords")) > 0 Then

                    drReader.Close()

                    Console.WriteLine(" Pending ScanList records detected.")
                    Console.Write("Processing ")
                    bUpdateVisual = True

                    'TO DO:  Execute queries to compensate for unresolved bugs until those bugs are fixed.
                    Dim bSuccess As Boolean = False

                    sb.Length = 0
                    sb.Append("update un_tracking.dbo.scanlist set barcode = upper(barcode)")
                    bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    sb.Length = 0
                    sb.Append("update un_tracking.dbo.scanlist set OperatorID = 'E0000001' where OperatorID = '0'")
                    bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    sb.Length = 0
                    sb.Append("update un_tracking.dbo.scanlist set PointID = 'P0000497' where PointID = '0'")
                    bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    ''WHY???
                    'sb.Length = 0
                    'sb.Append("update un_tracking.dbo.scanlist set BatchId = BatchId - 10000 where processed = 0 and batchid >= 10000")
                    'bSuccess = ExecuteQuery(sb.ToString, cmd)


                    ' Get Data from ScanList table
                    sb.Length = 0
                    'sb.Append("select top 1")
                    sb.Append("select ")
                    sb.Append(" RowId,")
                    'EventCode
                    sb.Append(" EventCode + '|' +")
                    'OperatorID
                    sb.Append(" OperatorID + '|' +")
                    'PointId
                    sb.Append(" PointId + '|' +")
                    'Barcode
                    sb.Append(" rtrim(Barcode) + '|' +")
                    'Weight
                    sb.Append(" cast(Weight as varchar) + '|' +")
                    'X
                    sb.Append(" case when charindex('of',x) = 0 then x else rtrim(substring(x,1,len(x) - (charindex('of',x) + 1))) end  + '|' +")
                    'ScanError
                    sb.Append(" ScanError + '|' +")
                    'BatchID
                    sb.Append(" cast(BatchId as varchar) + '|' +")
                    'ScanDate
                    sb.Append(" case when datepart(month,ScanDate) < 10 then '0' + cast(datepart(month,ScanDate) as varchar) else cast(datepart(month,ScanDate) as varchar) end +")
                    sb.Append(" case when datepart(day,ScanDate) < 10 then '0' + cast(datepart(day,ScanDate) as varchar) else cast(datepart(day,ScanDate) as varchar) end +	cast(datepart(year,ScanDate) as varchar) +")
                    sb.Append(" case when datepart(hour,ScanDate) < 10 then '0' + cast(datepart(hour,ScanDate) as varchar) else cast(datepart(hour,ScanDate) as varchar) end +")
                    sb.Append(" case when datepart(minute,ScanDate) < 10 then '0' + cast(datepart(minute,ScanDate) as varchar) else cast(datepart(minute,ScanDate) as varchar) end +")
                    sb.Append(" case when datepart(second,ScanDate) < 10 then '0' + cast(datepart(second,ScanDate) as varchar) else cast(datepart(second,ScanDate) as varchar) end + '|' +")
                    'BatchDate
                    sb.Append(" case when datepart(month,BatchDate) < 10 then '0' + cast(datepart(month,BatchDate) as varchar) else cast(datepart(month,BatchDate) as varchar) end +")
                    sb.Append(" case when datepart(day,BatchDate) < 10 then '0' + cast(datepart(day,BatchDate) as varchar) else cast(datepart(day,BatchDate) as varchar) end +	cast(datepart(year,BatchDate) as varchar) +")
                    sb.Append(" case when datepart(hour,BatchDate) < 10 then '0' + cast(datepart(hour,BatchDate) as varchar) else cast(datepart(hour,BatchDate) as varchar) end +")
                    sb.Append(" case when datepart(minute,BatchDate) < 10 then '0' + cast(datepart(minute,BatchDate) as varchar) else cast(datepart(minute,BatchDate) as varchar) end +")
                    sb.Append(" case when datepart(second,BatchDate) < 10 then '0' + cast(datepart(second,BatchDate) as varchar) else cast(datepart(second,BatchDate) as varchar) end + '|' +")
                    sb.Append(" cast(ErrorLog as varchar) + '|' +")
                    sb.Append(" HHid as RecordString")
                    sb.Append(" from un_tracking.dbo.scanlist")
                    sb.Append(" where ScanError % 10 = 0 and Processed = 0 and EventCode <> 'EZ'")

                    sCmd = sb.ToString()

                    PopulateDataset2(oDataAdapter, oDataSet, sCmd)

                    If Not oDataSet Is Nothing Then

                        If oDataSet.Tables.Count = 1 Then

                            If oDataSet.Tables(0).Rows.Count > 0 Then

                                Dim iRowId As Integer = oDataSet.Tables(0).Rows(0).Item("RowId")
                                Dim oScanList As New ScanList(oDataSet)
                                Dim i As Integer = -1

                                If Not oScanList.Records Is Nothing Then

                                    For Each oRec As ScanRecord In oScanList.Records

                                        Console.Write(".")

                                        i = i + 1
                                        iRowId = oDataSet.Tables(0).Rows(i).Item("RowId")

                                        If Not ImportScanListRecord(oRec) Then

                                            Console.WriteLine("[Module1.1906]" & oRec.ErrorMessage & ": RowId = " & iRowId.ToString())
                                            'If MessageBox.Show(oRec.ErrorMessage + ". Do you want to continue?", "Error Importing Record at RowID " & iRowId.ToString(), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                                            '    bReturnValue = False
                                            '    Exit For
                                            'End If
                                        Else

                                            sb.Length = 0
                                            sb.Append("UPDATE un_tracking.dbo.scanlist SET Processed = 1, ProcessDate = '" & Date.Now().ToShortDateString & " " & Date.Now().ToShortTimeString & "' WHERE RowId = " & iRowId)
                                            sCmd = sb.ToString()
                                            bReturnValue = ExecuteQuery(sCmd)

                                            If bReturnValue = False Then
                                                ''MessageBox.Show("ScanList Record at RowId " & iRowId & " was processed properly, but its flag was not updated", "Problem Importing ScanList Record", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                Console.WriteLine("ScanList Record at RowId " & iRowId & " was processed properly, but its flag was not updated", "Problem Importing ScanList Record")
                                            End If

                                        End If

                                    Next

                                Else

                                    ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Console.WriteLine("[Module1.1930]" & oScanList.ErrorMessage, "Import ScanList Status")
                                    bReturnValue = False

                                End If

                            Else

                                ''MessageBox.Show("No Records Founds to Import", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Console.WriteLine("[Module1.1938]" & "No Records Founds to Import", "Import ScanList Status")
                                bReturnValue = True

                            End If

                        Else

                            ''MessageBox.Show("There were no records to import", "Import ScanList Status", MessageBoxButtons.OK)
                            Console.WriteLine("[Module1.1946]" & "There were no records to import", "Import ScanList Status")
                            bReturnValue = True

                        End If

                    Else

                        ''MessageBox.Show("Database Error", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Console.WriteLine("[Module1.1954]" & "Database Error", "Import ScanList Status")
                        bReturnValue = False

                    End If

                End If

            End If

            If bUpdateVisual = True Then

                Console.WriteLine("Done!")

            End If

        Catch ex As Exception

            'Error Handling Code Goes Here
            bReturnValue = False
            Console.WriteLine("[Module1.1973]" & ex.Message)

        Finally

            'clean up code that need to run no matter what

            'close the data reader
            drReader.Close()

            If Not cnSqlQuery Is Nothing Then
                If cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Close()
                cnSqlQuery.Dispose()
            End If

            If Not cnSqlInsert Is Nothing Then
                If cnSqlInsert.State = ConnectionState.Open Then cnSqlInsert.Close()
                cnSqlInsert.Dispose()
            End If

            If Not cnSqlUpdate Is Nothing Then
                If cnSqlUpdate.State = ConnectionState.Open Then cnSqlUpdate.Close()
                cnSqlUpdate.Dispose()
            End If

            If Not cnSqlDelete Is Nothing Then
                If cnSqlDelete.State = ConnectionState.Open Then cnSqlDelete.Close()
                cnSqlDelete.Dispose()
            End If

        End Try

        Return bReturnValue

    End Function

    Private Sub PurgeScanList()

        ' local variables
        Dim oDataAdapter As SqlDataAdapter = Nothing
        Dim oDataSet As DataSet = Nothing

        Try
            ' Instantiate Connections
            cnSqlQuery = New SqlConnection(sScanListConnection)

            ' Instantiate Commands
            'cmd = New SqlCommand("SELECT count(*) as PendingRecords FROM UN_TRACKING.dbo.ScanList WHERE Processed = '0'", cnSqlQuery)
            cmd = New SqlCommand("SELECT count(*) as PendingRecords FROM UN_TRACKING.dbo.ScanList WHERE Processed = '0' and EventCode <> 'EZ'", cnSqlQuery)

            'open the connection
            If Not cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Open()

            'populate the DataReader
            drReader = cmd.ExecuteReader

            'Provide Visual Feedback
            If drReader.Read Then

                If CInt(drReader("PendingRecords")) > 0 Then

                    Dim sb As New StringBuilder
                    Dim sCmd As String

                    ' Get Data from ScanList table
                    sb.Append("select RowId from un_tracking.dbo.scanlist where ScanError % 10 <> 0 and Processed = 0 and EventCode <> 'EZ'")
                    sCmd = sb.ToString()

                    PopulateDataset2(oDataAdapter, oDataSet, sCmd)

                    If Not oDataSet Is Nothing Then

                        If oDataSet.Tables.Count = 1 Then

                            If oDataSet.Tables(0).Rows.Count > 0 Then

                                Dim iRowId As Integer

                                For Each dr As DataRow In oDataSet.Tables(0).Rows
                                    iRowId = dr.Item("RowId")
                                    sb.Length = 0
                                    sb.Append("UPDATE un_tracking.dbo.scanlist SET Processed = 1, ProcessDate = '" & Date.Now().ToShortDateString & " " & Date.Now().ToShortTimeString & "' WHERE RowId = " & iRowId)
                                    sCmd = sb.ToString()
                                    ExecuteQuery(sCmd)
                                Next

                            End If

                        End If

                    Else

                        ''MessageBox.Show("Database Error", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Console.WriteLine("[Module1.2064] Database Error while attempting to purge the scanlist.", "Import ScanList Status")

                    End If

                End If

            End If


        Catch ex As Exception

            'Error Handling Code Goes Here
            Console.WriteLine("[Module1.2076]" & ex.Message)

        Finally

            'clean up code that need to run no matter what

            'close the data reader
            drReader.Close()

            If Not cnSqlQuery Is Nothing Then
                cnSqlQuery.Close()
                cnSqlQuery.Dispose()
            End If

            If Not cnSqlInsert Is Nothing Then
                cnSqlInsert.Close()
                cnSqlInsert.Dispose()
            End If

            If Not cnSqlUpdate Is Nothing Then
                cnSqlUpdate.Close()
                cnSqlUpdate.Dispose()
            End If

            If Not cnSqlDelete Is Nothing Then
                cnSqlDelete.Close()
                cnSqlDelete.Dispose()
            End If

        End Try

    End Sub

    Private Function ImportEZScanListDD() As Boolean ' This will import the Delivery Events from EZ-Scan.  This function will put an EZ into ScanList EventCode which is translated to DD in Event Table

        ' local variables
        Dim bReturnValue As Boolean = True
        Dim bUpdateVisual As Boolean = False
        Dim sb As New StringBuilder
        Dim sCmd As String
        Dim oDataAdapter As SqlDataAdapter = Nothing
        Dim oDataSet As DataSet = Nothing

        Try
            ' Instantiate Connections
            cnSqlQuery = New SqlConnection(sScanListConnection)
            cnSqlInsert = New SqlConnection(sScanListConnection)

            ' Transfer New EZ-Scan Uploads to ScanList Table from TopCourier database
            cmdInsert = New SqlCommand("insert into UN_TRACKING.dbo.SCANLIST (StationRowID,EventCode,BatchId,Barcode,ScanDate,ScanError,HHid,OperatorID,PointID,[weight],X,BatchDate) select c.BarID, 'EZ' as EventCode, c.BarDocID, substring(c.Barcode,1,25) as Barcode, convert(datetime,SWITCHOFFSET(CONVERT(datetimeoffset,d.CreationDate),DATENAME(TzOffset,SYSDATETIMEOFFSET()))) as ScanDate, 0 as ScanError,'00EZ' as HHID, '0' as OperatorId, '0' as PointId, 0 as [Weight], '1' as X, convert(datetime,SWITCHOFFSET(CONVERT(datetimeoffset,d.CreationDate),DATENAME(TzOffset,SYSDATETIMEOFFSET()))) as BatchDate from TopCourier.dbo.tbl_BarcodeDetail c join TopCourier.dbo.tbl_BarcodeDoc d on c.BarDocID = d.BarDocID and d.Status = 'Delivery' and c.BarID > (select MAX(StationRowid) from UN_TRACKING.dbo.SCANLIST where EventCode = 'EZ')")
            If Not cnSqlInsert.State = ConnectionState.Open Then cnSqlInsert.Open()
            cmdInsert.Connection = cnSqlInsert
            cmdInsert.ExecuteNonQuery()

            ' Instantiate Commands
            'cmd = New SqlCommand("SELECT count(*) as PendingRecords FROM UN_TRACKING.dbo.ScanList WHERE Processed = 0", cnSqlQuery)
            cmd = New SqlCommand("SELECT count(*) as PendingRecords FROM UN_TRACKING.dbo.ScanList WHERE Processed = 0 and EventCode = 'EZ'", cnSqlQuery)

            ' Set Command Types

            'open the connection
            If Not cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Open()

            'populate the DataReader
            drReader = cmd.ExecuteReader

            'Provide Visual Feedback
            If drReader.Read Then

                If CInt(drReader("PendingRecords")) > 0 Then

                    drReader.Close()

                    Console.WriteLine(" Pending EZ (Delivery) ScanList records detected.")
                    Console.Write("Processing ")
                    bUpdateVisual = True

                    ''TO DO:  Execute queries to compensate for unresolved bugs until those bugs are fixed.
                    Dim bSuccess As Boolean = False

                    sb.Length = 0
                    sb.Append("update un_tracking.dbo.scanlist set barcode = upper(barcode)")
                    bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    sb.Length = 0
                    sb.Append("update un_tracking.dbo.scanlist set OperatorID = 'E0000001' where OperatorID = '0'")
                    bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    sb.Length = 0
                    sb.Append("update un_tracking.dbo.scanlist set PointID = 'P0000497' where PointID = '0'")
                    bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    ''WHY???
                    ''sb.Length = 0
                    ''sb.Append("update un_tracking.dbo.scanlist set BatchId = BatchId - 10000 where processed = 0 and batchid >= 10000")
                    ''bSuccess = ExecuteQuery(sb.ToString, cmd)


                    ' Get Data from ScanList table
                    ' NOTE: In Phase 1 of EZ-Scan, Only Delivery scans are supported thus Event Code is always 'DD'.
                    '       This will change in subsequent phases
                    sb.Length = 0
                    'sb.Append("select top 1")
                    sb.Append("select ")
                    sb.Append(" RowId,")
                    'EventCode
                    sb.Append(" 'DD' + '|' +")
                    'OperatorID
                    sb.Append(" OperatorID + '|' +")
                    'PointId
                    sb.Append(" PointId + '|' +")
                    'Barcode
                    sb.Append(" rtrim(Barcode) + '|' +")
                    'Weight
                    sb.Append(" cast(Weight as varchar) + '|' +")
                    'X
                    sb.Append(" case when charindex('of',x) = 0 then x else rtrim(substring(x,1,len(x) - (charindex('of',x) + 1))) end  + '|' +")
                    'ScanError
                    sb.Append(" ScanError + '|' +")
                    'BatchID
                    sb.Append(" cast(BatchId as varchar) + '|' +")
                    'ScanDate
                    sb.Append(" case when datepart(month,ScanDate) < 10 then '0' + cast(datepart(month,ScanDate) as varchar) else cast(datepart(month,ScanDate) as varchar) end +")
                    sb.Append(" case when datepart(day,ScanDate) < 10 then '0' + cast(datepart(day,ScanDate) as varchar) else cast(datepart(day,ScanDate) as varchar) end +	cast(datepart(year,ScanDate) as varchar) +")
                    sb.Append(" case when datepart(hour,ScanDate) < 10 then '0' + cast(datepart(hour,ScanDate) as varchar) else cast(datepart(hour,ScanDate) as varchar) end +")
                    sb.Append(" case when datepart(minute,ScanDate) < 10 then '0' + cast(datepart(minute,ScanDate) as varchar) else cast(datepart(minute,ScanDate) as varchar) end +")
                    sb.Append(" case when datepart(second,ScanDate) < 10 then '0' + cast(datepart(second,ScanDate) as varchar) else cast(datepart(second,ScanDate) as varchar) end + '|' +")
                    'BatchDate
                    sb.Append(" case when datepart(month,BatchDate) < 10 then '0' + cast(datepart(month,BatchDate) as varchar) else cast(datepart(month,BatchDate) as varchar) end +")
                    sb.Append(" case when datepart(day,BatchDate) < 10 then '0' + cast(datepart(day,BatchDate) as varchar) else cast(datepart(day,BatchDate) as varchar) end +	cast(datepart(year,BatchDate) as varchar) +")
                    sb.Append(" case when datepart(hour,BatchDate) < 10 then '0' + cast(datepart(hour,BatchDate) as varchar) else cast(datepart(hour,BatchDate) as varchar) end +")
                    sb.Append(" case when datepart(minute,BatchDate) < 10 then '0' + cast(datepart(minute,BatchDate) as varchar) else cast(datepart(minute,BatchDate) as varchar) end +")
                    sb.Append(" case when datepart(second,BatchDate) < 10 then '0' + cast(datepart(second,BatchDate) as varchar) else cast(datepart(second,BatchDate) as varchar) end + '|' +")
                    sb.Append(" cast(ErrorLog as varchar) + '|' +")
                    sb.Append(" HHid as RecordString")
                    sb.Append(" from un_tracking.dbo.scanlist")
                    sb.Append(" where ScanError % 10 = 0 and Processed = 0 and EventCode = 'EZ'")

                    sCmd = sb.ToString()

                    PopulateDataset2(oDataAdapter, oDataSet, sCmd)

                    If Not oDataSet Is Nothing Then

                        If oDataSet.Tables.Count = 1 Then

                            If oDataSet.Tables(0).Rows.Count > 0 Then

                                Dim iRowId As Integer = oDataSet.Tables(0).Rows(0).Item("RowId")
                                Dim oScanList As New ScanList(oDataSet)
                                Dim i As Integer = -1

                                If Not oScanList.Records Is Nothing Then

                                    For Each oRec As ScanRecord In oScanList.Records

                                        Console.Write(".")

                                        i = i + 1
                                        iRowId = oDataSet.Tables(0).Rows(i).Item("RowId")

                                        If Not ImportScanListRecord(oRec) Then

                                            Console.WriteLine("[Module1.1906]" & oRec.ErrorMessage & ": RowId = " & iRowId.ToString())
                                            'If MessageBox.Show(oRec.ErrorMessage + ". Do you want to continue?", "Error Importing Record at RowID " & iRowId.ToString(), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                                            '    bReturnValue = False
                                            '    Exit For
                                            'End If
                                        Else

                                            sb.Length = 0
                                            sb.Append("UPDATE un_tracking.dbo.scanlist SET Processed = 1, ProcessDate = '" & Date.Now().ToShortDateString & " " & Date.Now().ToShortTimeString & "' WHERE RowId = " & iRowId)
                                            sCmd = sb.ToString()
                                            bReturnValue = ExecuteQuery(sCmd)

                                            If bReturnValue = False Then
                                                ''MessageBox.Show("ScanList Record at RowId " & iRowId & " was processed properly, but its flag was not updated", "Problem Importing ScanList Record", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                Console.WriteLine("ScanList Record at RowId " & iRowId & " was processed properly, but its flag was not updated", "Problem Importing ScanList Record")
                                            End If

                                        End If

                                    Next

                                Else

                                    ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Console.WriteLine("[Module1.1930]" & oScanList.ErrorMessage, "Import ScanList Status")
                                    bReturnValue = False

                                End If

                            Else

                                ''MessageBox.Show("No Records Founds to Import", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Console.WriteLine("[Module1.1938]" & "No Records Founds to Import", "Import ScanList Status")
                                bReturnValue = True

                            End If

                        Else

                            ''MessageBox.Show("There were no records to import", "Import ScanList Status", MessageBoxButtons.OK)
                            Console.WriteLine("[Module1.1946]" & "There were no records to import", "Import ScanList Status")
                            bReturnValue = True

                        End If

                    Else

                        ''MessageBox.Show("Database Error", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Console.WriteLine("[Module1.1954]" & "Database Error", "Import ScanList Status")
                        bReturnValue = False

                    End If

                End If

            End If

            If bUpdateVisual = True Then

                Console.WriteLine("Done!")

            End If

        Catch ex As Exception

            'Error Handling Code Goes Here
            bReturnValue = False
            Console.WriteLine("[Module1.1973]" & ex.Message)

        Finally

            'clean up code that need to run no matter what

            'close the data reader
            drReader.Close()

            If Not cnSqlQuery Is Nothing Then
                If cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Close()
                cnSqlQuery.Dispose()
            End If

            If Not cnSqlInsert Is Nothing Then
                If cnSqlInsert.State = ConnectionState.Open Then cnSqlInsert.Close()
                cnSqlInsert.Dispose()
            End If

            If Not cnSqlUpdate Is Nothing Then
                If cnSqlUpdate.State = ConnectionState.Open Then cnSqlUpdate.Close()
                cnSqlUpdate.Dispose()
            End If

            If Not cnSqlDelete Is Nothing Then
                If cnSqlDelete.State = ConnectionState.Open Then cnSqlDelete.Close()
                cnSqlDelete.Dispose()
            End If

        End Try

        Return bReturnValue

    End Function

    Private Function ImportEZScanListAF() As Boolean ' This will import the Pickup Events from EZScan

        ' local variables
        Dim bReturnValue As Boolean = True
        Dim bUpdateVisual As Boolean = False
        Dim sb As New StringBuilder
        Dim sCmd As String
        Dim oDataAdapter As SqlDataAdapter = Nothing
        Dim oDataSet As DataSet = Nothing

        Try
            ' Instantiate Connections
            cnSqlQuery = New SqlConnection(sScanListConnection)
            cnSqlInsert = New SqlConnection(sScanListConnection)

            ' Transfer New EZ-Scan Uploads to ScanList Table from TopCourier database
            cmdInsert = New SqlCommand("insert into UN_TRACKING.dbo.SCANLIST (StationRowID,EventCode,BatchId,Barcode,ScanDate,ScanError,HHid,OperatorID,PointID,[weight],X,BatchDate) select c.BarID, 'AF' as EventCode, c.BarDocID, substring(c.Barcode,1,25) as Barcode, convert(datetime,SWITCHOFFSET(CONVERT(datetimeoffset,d.CreationDate),DATENAME(TzOffset,SYSDATETIMEOFFSET()))) as ScanDate, 0 as ScanError,'00EZ' as HHID, '0' as OperatorId, '0' as PointId, 0 as [Weight], '1' as X, convert(datetime,SWITCHOFFSET(CONVERT(datetimeoffset,d.CreationDate),DATENAME(TzOffset,SYSDATETIMEOFFSET()))) as BatchDate from TopCourier.dbo.tbl_BarcodeDetail c join TopCourier.dbo.tbl_BarcodeDoc d on c.BarDocID = d.BarDocID and d.Status = 'Pickup' and c.BarID > (select coalesce(MAX(StationRowid),2500) from UN_TRACKING.dbo.SCANLIST where EventCode = 'AF')")
            If Not cnSqlInsert.State = ConnectionState.Open Then cnSqlInsert.Open()
            cmdInsert.Connection = cnSqlInsert
            cmdInsert.ExecuteNonQuery()

            ' Instantiate Commands
            'cmd = New SqlCommand("SELECT count(*) as PendingRecords FROM UN_TRACKING.dbo.ScanList WHERE Processed = 0", cnSqlQuery)
            cmd = New SqlCommand("SELECT count(*) as PendingRecords FROM UN_TRACKING.dbo.ScanList WHERE Processed = 0 and EventCode = 'AF'", cnSqlQuery)

            ' Set Command Types

            'open the connection
            If Not cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Open()

            'populate the DataReader
            drReader = cmd.ExecuteReader

            'Provide Visual Feedback
            If drReader.Read Then

                If CInt(drReader("PendingRecords")) > 0 Then

                    drReader.Close()

                    Console.WriteLine(" Pending EZ (Pickup) ScanList records detected.")
                    Console.Write("Processing ")
                    bUpdateVisual = True

                    ''TO DO:  Execute queries to compensate for unresolved bugs until those bugs are fixed.
                    Dim bSuccess As Boolean = False

                    sb.Length = 0
                    sb.Append("update un_tracking.dbo.scanlist set barcode = upper(barcode)")
                    bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    sb.Length = 0
                    sb.Append("update un_tracking.dbo.scanlist set OperatorID = 'E0000001' where OperatorID = '0'")
                    bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    sb.Length = 0
                    sb.Append("update un_tracking.dbo.scanlist set PointID = 'P0000497' where PointID = '0'")
                    bSuccess = ExecuteQuery(sb.ToString(), cmd, False)

                    ''WHY???
                    ''sb.Length = 0
                    ''sb.Append("update un_tracking.dbo.scanlist set BatchId = BatchId - 10000 where processed = 0 and batchid >= 10000")
                    ''bSuccess = ExecuteQuery(sb.ToString, cmd)


                    ' Get Data from ScanList table
                    ' NOTE: In Phase 1 of EZ-Scan, Only Delivery scans are supported thus Event Code is always 'DD'.
                    '       This will change in subsequent phases
                    sb.Length = 0
                    'sb.Append("select top 1")
                    sb.Append("select ")
                    sb.Append(" RowId,")
                    'EventCode
                    sb.Append(" 'AF' + '|' +")
                    'OperatorID
                    sb.Append(" OperatorID + '|' +")
                    'PointId
                    sb.Append(" PointId + '|' +")
                    'Barcode
                    sb.Append(" rtrim(Barcode) + '|' +")
                    'Weight
                    sb.Append(" cast(Weight as varchar) + '|' +")
                    'X
                    sb.Append(" case when charindex('of',x) = 0 then x else rtrim(substring(x,1,len(x) - (charindex('of',x) + 1))) end  + '|' +")
                    'ScanError
                    sb.Append(" ScanError + '|' +")
                    'BatchID
                    sb.Append(" cast(BatchId as varchar) + '|' +")
                    'ScanDate
                    sb.Append(" case when datepart(month,ScanDate) < 10 then '0' + cast(datepart(month,ScanDate) as varchar) else cast(datepart(month,ScanDate) as varchar) end +")
                    sb.Append(" case when datepart(day,ScanDate) < 10 then '0' + cast(datepart(day,ScanDate) as varchar) else cast(datepart(day,ScanDate) as varchar) end +	cast(datepart(year,ScanDate) as varchar) +")
                    sb.Append(" case when datepart(hour,ScanDate) < 10 then '0' + cast(datepart(hour,ScanDate) as varchar) else cast(datepart(hour,ScanDate) as varchar) end +")
                    sb.Append(" case when datepart(minute,ScanDate) < 10 then '0' + cast(datepart(minute,ScanDate) as varchar) else cast(datepart(minute,ScanDate) as varchar) end +")
                    sb.Append(" case when datepart(second,ScanDate) < 10 then '0' + cast(datepart(second,ScanDate) as varchar) else cast(datepart(second,ScanDate) as varchar) end + '|' +")
                    'BatchDate
                    sb.Append(" case when datepart(month,BatchDate) < 10 then '0' + cast(datepart(month,BatchDate) as varchar) else cast(datepart(month,BatchDate) as varchar) end +")
                    sb.Append(" case when datepart(day,BatchDate) < 10 then '0' + cast(datepart(day,BatchDate) as varchar) else cast(datepart(day,BatchDate) as varchar) end +	cast(datepart(year,BatchDate) as varchar) +")
                    sb.Append(" case when datepart(hour,BatchDate) < 10 then '0' + cast(datepart(hour,BatchDate) as varchar) else cast(datepart(hour,BatchDate) as varchar) end +")
                    sb.Append(" case when datepart(minute,BatchDate) < 10 then '0' + cast(datepart(minute,BatchDate) as varchar) else cast(datepart(minute,BatchDate) as varchar) end +")
                    sb.Append(" case when datepart(second,BatchDate) < 10 then '0' + cast(datepart(second,BatchDate) as varchar) else cast(datepart(second,BatchDate) as varchar) end + '|' +")
                    sb.Append(" cast(ErrorLog as varchar) + '|' +")
                    sb.Append(" HHid as RecordString")
                    sb.Append(" from un_tracking.dbo.scanlist")
                    sb.Append(" where ScanError % 10 = 0 and Processed = 0 and EventCode = 'AF'")

                    sCmd = sb.ToString()

                    PopulateDataset2(oDataAdapter, oDataSet, sCmd)

                    If Not oDataSet Is Nothing Then

                        If oDataSet.Tables.Count = 1 Then

                            If oDataSet.Tables(0).Rows.Count > 0 Then

                                Dim iRowId As Integer = oDataSet.Tables(0).Rows(0).Item("RowId")
                                Dim oScanList As New ScanList(oDataSet)
                                Dim i As Integer = -1

                                If Not oScanList.Records Is Nothing Then

                                    For Each oRec As ScanRecord In oScanList.Records

                                        Console.Write(".")

                                        i = i + 1
                                        iRowId = oDataSet.Tables(0).Rows(i).Item("RowId")

                                        If Not ImportScanListRecord(oRec) Then

                                            Console.WriteLine("[Module1.1906]" & oRec.ErrorMessage & ": RowId = " & iRowId.ToString())
                                            'If MessageBox.Show(oRec.ErrorMessage + ". Do you want to continue?", "Error Importing Record at RowID " & iRowId.ToString(), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                                            '    bReturnValue = False
                                            '    Exit For
                                            'End If
                                        Else

                                            sb.Length = 0
                                            sb.Append("UPDATE un_tracking.dbo.scanlist SET Processed = 1, ProcessDate = '" & Date.Now().ToShortDateString & " " & Date.Now().ToShortTimeString & "' WHERE RowId = " & iRowId)
                                            sCmd = sb.ToString()
                                            bReturnValue = ExecuteQuery(sCmd)

                                            If bReturnValue = False Then
                                                ''MessageBox.Show("ScanList Record at RowId " & iRowId & " was processed properly, but its flag was not updated", "Problem Importing ScanList Record", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                Console.WriteLine("ScanList Record at RowId " & iRowId & " was processed properly, but its flag was not updated", "Problem Importing ScanList Record")
                                            End If

                                        End If

                                    Next

                                Else

                                    ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    Console.WriteLine("[Module1.1930]" & oScanList.ErrorMessage, "Import ScanList Status")
                                    bReturnValue = False

                                End If

                            Else

                                ''MessageBox.Show("No Records Founds to Import", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Console.WriteLine("[Module1.1938]" & "No Records Founds to Import", "Import ScanList Status")
                                bReturnValue = True

                            End If

                        Else

                            ''MessageBox.Show("There were no records to import", "Import ScanList Status", MessageBoxButtons.OK)
                            Console.WriteLine("[Module1.1946]" & "There were no records to import", "Import ScanList Status")
                            bReturnValue = True

                        End If

                    Else

                        ''MessageBox.Show("Database Error", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Console.WriteLine("[Module1.1954]" & "Database Error", "Import ScanList Status")
                        bReturnValue = False

                    End If

                End If

            End If

            If bUpdateVisual = True Then

                Console.WriteLine("Done!")

            End If

        Catch ex As Exception

            'Error Handling Code Goes Here
            bReturnValue = False
            Console.WriteLine("[Module1.1973]" & ex.Message)

        Finally

            'clean up code that need to run no matter what

            'close the data reader
            drReader.Close()

            If Not cnSqlQuery Is Nothing Then
                If cnSqlQuery.State = ConnectionState.Open Then cnSqlQuery.Close()
                cnSqlQuery.Dispose()
            End If

            If Not cnSqlInsert Is Nothing Then
                If cnSqlInsert.State = ConnectionState.Open Then cnSqlInsert.Close()
                cnSqlInsert.Dispose()
            End If

            If Not cnSqlUpdate Is Nothing Then
                If cnSqlUpdate.State = ConnectionState.Open Then cnSqlUpdate.Close()
                cnSqlUpdate.Dispose()
            End If

            If Not cnSqlDelete Is Nothing Then
                If cnSqlDelete.State = ConnectionState.Open Then cnSqlDelete.Close()
                cnSqlDelete.Dispose()
            End If

        End Try

        Return bReturnValue

    End Function

    'Friend Function ImportScanList() As Boolean
    '    Dim bReturnValue As Boolean = True

    '    ' Data Access Variables
    '    Dim oDataAdapter As SqlDataAdapter
    '    Dim oDataSet As DataSet
    '    Dim oDailyEntry As DataRow
    '    Dim iRowCount As Integer

    '    ' Utility Variables
    '    Dim sb As New StringBuilder
    '    Dim sCmd As String

    '    Try
    '        ' Get Data from ScanList table
    '        sb.Append("select")
    '        sb.Append(" RowId,")
    '        'EventCode
    '        sb.Append(" EventCode + '|' +")
    '        'OperatorID
    '        sb.Append(" OperatorID + '|' +")
    '        'PointId
    '        sb.Append(" PointId + '|' +")
    '        'Barcode
    '        sb.Append(" rtrim(Barcode) + '|' +")
    '        'Weight
    '        sb.Append(" cast(Weight as varchar) + '|' +")
    '        'X
    '        sb.Append(" case when charindex('of',x) = 0 then x else rtrim(substring(x,1,len(x) - (charindex('of',x) + 1))) end  + '|' +")
    '        'ScanError
    '        sb.Append(" ScanError + '|' +")
    '        'BatchID
    '        sb.Append(" cast(BatchId as varchar) + '|' +")
    '        'ScanDate
    '        sb.Append(" case when datepart(month,ScanDate) < 10 then '0' + cast(datepart(month,ScanDate) as varchar) else cast(datepart(month,ScanDate) as varchar) end +")
    '        sb.Append(" case when datepart(day,ScanDate) < 10 then '0' + cast(datepart(day,ScanDate) as varchar) else cast(datepart(day,ScanDate) as varchar) end +	cast(datepart(year,ScanDate) as varchar) +")
    '        sb.Append(" case when datepart(hour,ScanDate) < 10 then '0' + cast(datepart(hour,ScanDate) as varchar) else cast(datepart(hour,ScanDate) as varchar) end +")
    '        sb.Append(" case when datepart(minute,ScanDate) < 10 then '0' + cast(datepart(minute,ScanDate) as varchar) else cast(datepart(minute,ScanDate) as varchar) end +")
    '        sb.Append(" case when datepart(second,ScanDate) < 10 then '0' + cast(datepart(second,ScanDate) as varchar) else cast(datepart(second,ScanDate) as varchar) end + '|' +")
    '        'BatchDate
    '        sb.Append(" case when datepart(month,BatchDate) < 10 then '0' + cast(datepart(month,BatchDate) as varchar) else cast(datepart(month,BatchDate) as varchar) end +")
    '        sb.Append(" case when datepart(day,BatchDate) < 10 then '0' + cast(datepart(day,BatchDate) as varchar) else cast(datepart(day,BatchDate) as varchar) end +	cast(datepart(year,BatchDate) as varchar) +")
    '        sb.Append(" case when datepart(hour,BatchDate) < 10 then '0' + cast(datepart(hour,BatchDate) as varchar) else cast(datepart(hour,BatchDate) as varchar) end +")
    '        sb.Append(" case when datepart(minute,BatchDate) < 10 then '0' + cast(datepart(minute,BatchDate) as varchar) else cast(datepart(minute,BatchDate) as varchar) end +")
    '        sb.Append(" case when datepart(second,BatchDate) < 10 then '0' + cast(datepart(second,BatchDate) as varchar) else cast(datepart(second,BatchDate) as varchar) end + '|' +")
    '        sb.Append(" cast(ErrorLog as varchar) as RecordString")
    '        sb.Append(" from un_tracking.dbo.scanlist")
    '        sb.Append(" where ScanError % 10 = 0 and Processed = 0")

    '        sCmd = sb.ToString()

    '        PopulateDataset2(oDataAdapter, oDataSet, sCmd)

    '        If Not oDataSet Is Nothing Then
    '            If oDataSet.Tables.Count = 1 Then
    '                If oDataSet.Tables(0).Rows.Count > 0 Then
    '                    Dim iRowId As Integer = oDataSet.Tables(0).Rows(0).Item("RowId")
    '                    Dim oScanList As New ScanList(oDataSet)
    '                    Dim i As Integer = -1

    '                    If Not oScanList.Records Is Nothing Then
    '                        For Each oRec As ScanRecord In oScanList.Records
    '                            i = i + 1
    '                            iRowId = oDataSet.Tables(0).Rows(i).Item("RowId")

    '                            If Not ImportScanListRecord(oRec) Then
    '                                If MessageBox.Show(oRec.ErrorMessage + ". Do you want to continue?", "Error Importing Record at RowID " & iRowId.ToString(), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
    '                                    bReturnValue = False
    '                                    Exit For
    '                                End If
    '                            Else
    '                                sb.Length = 0
    '                                sb.Append("UPDATE un_tracking.dbo.scanlist SET Processed = 1, ProcessDate = '" & Date.Now().ToShortDateString & " " & Date.Now().ToShortTimeString & "' WHERE RowId = " & iRowId)
    '                                sCmd = sb.ToString()
    '                                bReturnValue = ExecuteQuery(sCmd)

    '                                If bReturnValue = False Then
    '                                    MessageBox.Show("ScanList Record at RowId " & iRowId & " was processed properly, but its flag was not updated", "Problem Importing ScanList Record", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    '                                End If
    '                            End If
    '                        Next
    '                    Else
    '                        MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                        bReturnValue = False
    '                    End If
    '                Else
    '                    MessageBox.Show("No Records Founds to Import", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
    '                    bReturnValue = True
    '                End If
    '            Else
    '                MessageBox.Show("There were no records to import", "Import ScanList Status", MessageBoxButtons.OK)
    '                bReturnValue = True
    '            End If
    '        Else
    '            MessageBox.Show("Database Error", "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '            bReturnValue = False
    '        End If
    '    Catch ex As Exception
    '        bReturnValue = False
    '    End Try
    '    Return bReturnValue
    'End Function

    Private Function ImportScanListRecord(ByVal p_oRec As ScanRecord) As Boolean

        Dim bReturnValue As Boolean = False

        Try

            Select Case p_oRec.BarcodeFormat
                Case BarcodeFactory.BarcodeFormat.TPC_Tracking
                    bReturnValue = ImportTPCTrkRec(p_oRec)
                Case BarcodeFactory.BarcodeFormat.Unity_HOSS
                    bReturnValue = ImportUnityHossRec(p_oRec)
                Case BarcodeFactory.BarcodeFormat.Unknown
                    bReturnValue = ImportThirdPartyRec(p_oRec)
                Case BarcodeFactory.BarcodeFormat.TPC_Operator, BarcodeFactory.BarcodeFormat.TPC_Point
                    ' These formats are not imported
                    bReturnValue = True
                Case Else
                    '  Create an exception log to show which records were not imported
                    'MessageBox.Show(oScanList.ErrorMessage, "Unreadable Format")
            End Select

        Catch ex As Exception

            bReturnValue = False

        End Try

        Return bReturnValue

    End Function

    Private Sub AppendEzScanData(ByRef p_oRec As ScanRecord, ByRef p_oEvent As TrackingEvent)

        Dim oDataAdapter As SqlDataAdapter = Nothing
        Dim oDataSet As DataSet = Nothing
        Dim sb As New StringBuilder

        sb.Append("Select * from TopCourier.dbo.tbl_BarcodeDoc where BarDocID = ")
        sb.Append(p_oRec.BatchID)

        PopulateDataset2(oDataAdapter, oDataSet, sb.ToString())

        If Not oDataSet Is Nothing Then

            If oDataSet.Tables(0).Rows.Count = 1 Then
                Dim oBarcodeDocRow As DataRow = oDataSet.Tables(0).Rows(0)
                Dim sUserId As String = String.Empty
                If TypeOf (oBarcodeDocRow.Item("UserID")) Is System.DBNull Then
                    sUserId = String.Empty
                Else
                    sUserId = oBarcodeDocRow.Item("UserId")
                End If
                If String.Compare(sUserId, "141") = 0 Then
                    p_oEvent.SignatureFile = "National Consolidated Couriers/" & oBarcodeDocRow.Item("Signature")
                Else
                    p_oEvent.SignatureFile = "Top Priority/" & oBarcodeDocRow.Item("Signature")
                End If
                p_oEvent.DeliveryComments = oBarcodeDocRow.Item("Name")
            End If

            End If

    End Sub

    Private Sub AppendEzScanData(ByRef p_oRec As ScanRecord, ByRef p_sDeliveryComments As String, ByRef p_sSignatureFile As String)

        Dim oDataAdapter As SqlDataAdapter = Nothing
        Dim oDataSet As DataSet = Nothing
        Dim sb As New StringBuilder

        sb.Append("Select * from TopCourier.dbo.tbl_BarcodeDoc where BarDocID = ")
        sb.Append(p_oRec.BatchID)

        PopulateDataset2(oDataAdapter, oDataSet, sb.ToString())

        If Not oDataSet Is Nothing Then

            If oDataSet.Tables(0).Rows.Count = 1 Then
                Dim oBarcodeDocRow As DataRow = oDataSet.Tables(0).Rows(0)
                Dim sUserId As String
                If TypeOf (oBarcodeDocRow.Item("UserId")) Is System.DBNull Then
                    sUserId = String.Empty
                Else
                    sUserId = oBarcodeDocRow.Item("UserId")
                End If
                If String.Compare(sUserId, "141") = 0 Then
                    p_sSignatureFile = "National Consolidated Couriers/" & oBarcodeDocRow.Item("Signature")
                Else
                    p_sSignatureFile = "Top Priority/" & oBarcodeDocRow.Item("Signature")
                End If
                p_sDeliveryComments = oBarcodeDocRow.Item("Name")
            End If

        End If

    End Sub

    Private Function ImportThirdPartyRec(ByVal p_oRec As ScanRecord) As Boolean

        'Dim strRec As String

        'strRec = p_oRec.EventCode & " + " & _
        'p_oRec.OperatorId & " + " & _
        'p_oRec.PointId & " + " & _
        'p_oRec.Barcode & " + " & _
        'p_oRec.Weight & " + " & _
        'p_oRec.TimeStamp

        Dim oBarcode As New Barcode(p_oRec.Barcode)

        ''Dim strCaption As String
        ''strCaption = "ImportThirdPartyRec(" & p_oRec.BarcodeName & ")"
        ''MessageBox.Show(strRec, strCaption)

        ' Create an Event with all know data
        Dim oEvent As New TrackingEvent

        ' Assign Values to oEvent and Persist
        oEvent.EventCode = p_oRec.EventCode
        oEvent.ScanDate = p_oRec.TimeStamp
        oEvent.OperatorId = New TPCOperatorBC(p_oRec.OperatorId)
        oEvent.PointId = New TPCPointBC(p_oRec.PointId)
        oEvent.ThirdPartyBarcode = oBarcode
        oEvent.Weight = p_oRec.Weight
        oEvent.Pieces = "1/1"
        oEvent.Void = False
        oEvent.BatchNumber = p_oRec.BatchID
        oEvent.ScannerId = p_oRec.HHid

        ' Add EZ-Scan Data if Applicable
        If String.Compare(p_oRec.EventCode, "DD") = 0 Then
            AppendEzScanData(p_oRec, oEvent)
        End If

        ' Add Support for Branch Delivery which will contain the route name in Delivery Comments field
        If String.Compare(p_oRec.EventCode, "BD") = 0 Then
            oEvent.DeliveryComments = p_oRec.DeliveryComments
        End If

        Return oEvent.Insert()

    End Function

    Private Function ImportUnityHossRec(ByVal p_oRec As ScanRecord) As Boolean

        '  There will be 3 categories of HOSS barcodes to process
        '   1) We have both the Destination & Source Location in our Database
        '       a) this will result in the most complete Event record entry since all data will be at our fingertips
        '   2) We have the Destination but not the Source in our Database
        '       a) this will result in complete info for the Source fields, but empty data for the source records
        '   3) We have neither the Source nor the Destination in our Database
        '       a) this will result in as much data as we can muster from the barcode itself

        '   Besides the destination and source location info, we still need to determine who is the paying customer for the movement.
        '   In order to do that we would need to import C info from the HOSS manifests and then map that C number to our Unison account 
        '   numbers.  Even then, we will only have access to the full C information for customers who concer our branches.
        '
        '   Given that we need so much info, the first implementation of this function will only record the data that we know to be
        '   100% accurate.

        Dim bRetVal As Boolean

        Try

            ' Extract the various parts of the Barcode
            Dim oBarcode As New HossBarcode(p_oRec.Barcode)

            ' Determine if Destination Location is Unique throughout System
            Dim oToLoc As New Location
            oToLoc.GetIfUniqueLocId(oBarcode.DestinationCode)

            ' Determine if Source Location is Unique throughout System
            Dim oFromLoc As New Location
            oFromLoc.GetIfUniqueLocId(oBarcode.SourceCode)

            ' Create an Event with all know data
            Dim oEvent As New TrackingEvent

            ' Assign Values to oEvent and Persist
            oEvent.EventCode = p_oRec.EventCode
            oEvent.ScanDate = p_oRec.TimeStamp
            oEvent.OperatorId = New TPCOperatorBC(p_oRec.OperatorId)
            oEvent.PointId = New TPCPointBC(p_oRec.PointId)
            oEvent.ThirdPartyBarcode = oBarcode
            oEvent.ParcelType = oBarcode.ServiceType
            oEvent.Weight = p_oRec.Weight
            oEvent.Pieces = "1/1"
            oEvent.Void = False
            If Not oToLoc.IsEmpty Then
                oEvent.ToLocationId = oToLoc.LocationID
                oEvent.ToAddressId = oToLoc.AddressId
                oEvent.ToLocationName = oToLoc.Name
            End If
            If Not oFromLoc.IsEmpty Then
                oEvent.FromLocationId = oFromLoc.LocationID
                oEvent.FromAddressId = oFromLoc.AddressId
                oEvent.FromLocationName = oFromLoc.Name
            End If

            bRetVal = oEvent.Insert()
            'bRetVal = True

        Catch ex As Exception

            ''MessageBox.Show(ex.Message)
            Console.WriteLine("[Module1.2337]" & ex.Message)
            bRetVal = False

        End Try

        Return bRetVal

    End Function

    Private Function ImportTPCTrkRec(ByVal p_oRec As ScanRecord) As Boolean

        Dim bReturnValue As Boolean = True

        Try

            ' Convert Barcode Strings into Specific Barcode Objects
            Dim oBC As New TPCBarcode(p_oRec.Barcode)

            ' Declare Variable for Datatbase access
            Dim oDataAdapter As SqlDataAdapter = Nothing
            Dim oDataSet As DataSet = Nothing
            Dim iCourierLabelID As Integer

            ' Declare Utility Variables
            Dim sb As New StringBuilder

            ' Determine if this Barcode is recorded in the CourierLabel Table and Act Accordingly
            ' It does not matter if it is voided or not; we want to record as much detail as possible for the scan.
            sb.Append("SELECT * FROM ")
            sb.Append(TRCTblPath)
            sb.Append("CourierLabels WHERE TrackingNum = '")
            sb.Append(oBC.Barcode)
            sb.Append("'")

            PopulateDataset2(oDataAdapter, oDataSet, sb.ToString())

            If Not oDataSet Is Nothing Then

                Dim iCourierLabelRows As Integer = oDataSet.Tables(0).Rows.Count
                Dim oCourierLabelRow As DataRow = Nothing

                Select Case iCourierLabelRows

                    Case 0 ' Info From ScanRecord Only; no entry in Event table possible.

                        'If S type, look for corresponding 'L' entry in Event Table.  
                        'If found, update Manifest & ManifestInvoice Table with actual weight
                        If String.Compare(oBC.PackageType, "S") = 0 Then

                            Dim sRowID As String = MatchingLabelFound(oBC.Barcode, p_oRec.TimeStamp)
                            If Not sRowID Is Nothing Then

                                bReturnValue = InsertEventAndUpdateManifest(p_oRec, sRowID)

                            Else

                                bReturnValue = False 'Data may not have been uploaded by the customer yet.  This ScanList record will be attempted again next time.

                            End If

                        Else

                            bReturnValue = InsertEventRecord(p_oRec)

                        End If

                    Case 1

                        ' Merge CourierLabel & ScanRecord Information then make entry into Event Table
                        oCourierLabelRow = oDataSet.Tables(0).Rows(0)
                        iCourierLabelID = oCourierLabelRow.Item("RowID")

                        ''BEGIN MODIFICATION for v2.39
                        'bReturnValue = InsertEventRecord(p_oRec, oCourierLabelRow)
                        If String.Compare(oBC.PackageType, "S") = 0 Then
                            Dim sRowID As String = MatchingLabelFound(oBC.Barcode, p_oRec.TimeStamp)
                            If Not sRowID Is Nothing Then
                                bReturnValue = InsertEventAndUpdateManifest(p_oRec, sRowID)
                            Else
                                bReturnValue = InsertEventRecord(p_oRec, oCourierLabelRow)
                            End If
                        Else
                            bReturnValue = InsertEventRecord(p_oRec, oCourierLabelRow)
                        End If
                        ''END MODIFICATION for v2.39

                        ' If Event Update was successful, apply weight charges to DailyEntry if applicable
                        If bReturnValue = True Then

                            'Add weight to all ACTIVE associations
                            sb.Remove(0, sb.Length)
                            sb.Append("SELECT * FROM ")
                            sb.Append(WEIGHTTblPath)
                            sb.Append("TrackingLink WHERE CourierLabelID = ")
                            sb.Append(iCourierLabelID)
                            sb.Append(" and Active = 1 Order by RowId desc")

                            oDataSet.Dispose()

                            PopulateDataset2(oDataAdapter, oDataSet, sb.ToString())

                            If Not oDataSet Is Nothing Then

                                'IMPORTANT NOTE:    This code is flawed.
                                '04/07/11           The intent of the following code was to allow a single barcode to be assigned
                                '                   to multiple weight plans, so instead of pasing the most current row, all rows 
                                '                   were passed.  This introduced a bug since inactive associations were also passed
                                '                   in.  The passing of inactive associations is not in and of itself a bug, since we want 
                                '                   to capture weight of old flip cards that may still be in circulation.  The solution to
                                '                   this problem is to treat active and inactive associations separately.


                                If oDataSet.Tables(0).Rows.Count > 0 Then 'Only most current assignment (past or present) is taken into account

                                    For Each dr As DataRow In oDataSet.Tables(0).Rows

                                        ''bReturnValue = InputDailyEntryRecord(p_oRec, oCourierLabelRow, oDataSet.Tables(0).Rows(0))
                                        bReturnValue = InputDailyEntryRecord(p_oRec, oCourierLabelRow, dr)

                                        If bReturnValue = False Then
                                            'TO-DO
                                            'Entry into Event table should be rolled back.  If a tracking number has a link to a 
                                            'weight plan, but the charge cannot be inserted, then the entire record should be rejected.
                                        End If

                                    Next

                                Else

                                    'Add weight to all IN-ACTIVE associations.  Since this is an ELSE, we can assume there are no ACTIVE 
                                    'associations for this barcode.  This is meant to capture weight for old flip cards that are still
                                    'being handled despite and service stop, service change, or change of address.
                                    sb.Length = 0

                                    sb.Remove(0, sb.Length)
                                    sb.Append("SELECT * FROM ")
                                    sb.Append(WEIGHTTblPath)
                                    sb.Append("TrackingLink WHERE CourierLabelID = ")
                                    sb.Append(iCourierLabelID)
                                    sb.Append(" and Active = 0 Order by RowId desc")

                                    oDataSet.Dispose()

                                    PopulateDataset2(oDataAdapter, oDataSet, sb.ToString())

                                    If Not oDataSet Is Nothing Then

                                        Dim iCount As Integer = oDataSet.Tables(0).Rows.Count

                                        If iCount = 1 Then

                                            bReturnValue = InputDailyEntryRecord(p_oRec, oCourierLabelRow, oDataSet.Tables(0).Rows(0))

                                            If bReturnValue = False Then
                                                'TO-DO
                                                'Entry into Event table should be rolled back.  If a tracking number has a link to a 
                                                'weight plan, but the charge cannot be inserted, then the entire record should be rejected.
                                            End If

                                        ElseIf iCount > 1 Then

                                            'TO-DO
                                            'Logic must be developed for this condition.  A tracking number may have been
                                            'assiged to more than 1 weight plan during its "active life" and thus would have multiple inactive
                                            'entries in the tracking link table which need to be * charged.  The problem is that a tracking number
                                            'that has only been assigned to one weight plan during its "active life" may also have multiple
                                            'inactive entries in the tracking link table if (for example) it has undergone multiple 
                                            'change-of-address) procedures.  The logic needs to determine whether the single charge in questions
                                            'needs to apply to 1 "inactive" plan or multiple "inactive" plans.  Suggestion tact:  Add a column to
                                            'the TrackingLink table that specifies the exact reason the link was deactivated.  This field could
                                            'be modeled after the HistoryLog field in Scanlist. -Sammy Nava

                                            'For Each dr As DataRow In oDataSet.Tables(0).Rows

                                            '    bReturnValue = InputDailyEntryRecord(p_oRec, oCourierLabelRow, dr)

                                            '    If bReturnValue = False Then
                                            '        'TO-DO
                                            '        'Entry into Event table should be rolled back.  If a tracking number has a link to a 
                                            '        'weight plan, but the charge cannot be inserted, then the entire record should be rejected.
                                            '    End If

                                            'Next

                                        End If

                                    End If

                                End If
                                'CX-00932-2337640-1485
                                bReturnValue = True 'No Link to Weight Plan.  Not a fatal error.

                            End If

                        Else

                            bReturnValue = False 'Error Inserting into Event Table

                        End If

                End Select

            Else

                bReturnValue = False 'Problem Quering Database

            End If

        Catch ex As Exception

            ''MessageBox.Show(ex.Message)
            Console.WriteLine("[Module1.2546]" & ex.Message)
            Return False

        End Try

        Return bReturnValue

    End Function

    Private Function InputDailyEntryRecord(ByVal p_oRec As ScanRecord, ByVal p_oCLRow As DataRow, ByVal p_oTLRow As DataRow) As Boolean

        'Make sure we only import WC rows
        If String.Compare(p_oRec.EventCode, "WC") <> 0 Then
            Return True 'Record was reviewed and no DailyEntry required
        End If

        ' Utility variables
        Dim bRetVal1 As Boolean = True
        Dim bRetVal2 As Boolean = True

        ' Convert Barcode Strings into Specific Barcode Objects for Easier Manipulation
        Dim oBC As New TPCBarcode(p_oRec.Barcode)

        ' Extract the necessary information from the CourierLabels Record
        Dim iWeight As Decimal = p_oRec.Weight
        Dim iWeightPlanID As Integer = p_oTLRow.Item("WeightPlanID")
        Dim sTranDate As String '= p_oRec.TimeStamp.ToShortDateString
        'If p_oRec.TimeStamp.Hour < 7 Then
        '    sTranDate = p_oRec.TimeStamp.AddDays(-1).ToShortDateString
        'Else
        '    sTranDate = p_oRec.TimeStamp.ToShortDateString
        'End If
        sTranDate = p_oRec.BatchDate

        ' Round the weight up or down for billing purposes.  Rule: If decimal portion of weight is > 0.25, round up, otherwide round down

        'Temporarily account for negative weight

        'Dim iPortion As Integer = IntegerPortion(iWeight)
        'Dim fPortion As Decimal = DecimalPortion(iWeight)

        'If fPortion > 0.25 Then
        '    iWeight = iPortion + 1 'Round UP
        'Else
        '    iWeight = iPortion ' Round Down
        'End If

        iWeight = Math.Abs(p_oRec.Weight)
        Dim iPortion As Integer = IntegerPortion(iWeight)
        Dim fPortion As Decimal = DecimalPortion(iWeight)

        If fPortion > 0.25 Then
            iWeight = iPortion + 1 'Round UP
        Else
            iWeight = iPortion ' Round Down
        End If
        If p_oRec.Weight < 0 Then iWeight = iWeight * -1

        ' Utility Variables
        Dim sCmd As String

        If DailyEntryExists(sTranDate, iWeightPlanID) Then
            'Prepare an Update Statement
            sCmd = DailyEntryUpdateString(sTranDate, iWeightPlanID, iWeight)
        Else
            'Prepare an Insert Statement
            sCmd = DailyEntryInsertString(sTranDate, iWeightPlanID, iWeight)
        End If

        ' Execute the Insert/Update
        bRetVal1 = ExecuteQuery(sCmd)

        ' Determine if the weight plan has a parent weight plan.  If it does, create a Daily entry for the parent as well.
        Dim iParentWeightPlan As Integer = GetParentWeightPlan(iWeightPlanID)

        'If iParentWeightPlan <> 0 Then
        '    If DailyEntryExists(sTranDate, iParentWeightPlan) Then
        '        'Prepare an Update Statement
        '        sCmd = DailyEntryUpdateString(sTranDate, iParentWeightPlan, iWeight)
        '    Else
        '        'Prepare an Insert Statement
        '        sCmd = DailyEntryInsertString(sTranDate, iParentWeightPlan, iWeight)
        '    End If
        '    'Execute the Insert/Update
        '    bRetVal2 = ExecuteQuery(sCmd)
        'End If

        ' Loop Backwards to Apply Charges to all Ancestors (if any)
        Do While iParentWeightPlan <> 0

            If DailyEntryExists(sTranDate, iParentWeightPlan) Then
                sCmd = DailyEntryUpdateString(sTranDate, iParentWeightPlan, iWeight)
            Else
                sCmd = DailyEntryInsertString(sTranDate, iParentWeightPlan, iWeight)
            End If

            bRetVal2 = ExecuteQuery(sCmd)

            iParentWeightPlan = GetParentWeightPlan(iParentWeightPlan)

        Loop

        'Return true if either statement succeeded, False if they both failed.
        If (bRetVal1 = False) And (bRetVal2 = False) Then
            Return False
        Else
            Return True
        End If

    End Function

    Private Function EventTableInsertStatement(ByVal p_oRec As ScanRecord, Optional ByVal p_oRow As DataRow = Nothing) As String

        ' Convert Barcode Strings into Specific Barcode Objects for Easier Manipulation
        Dim oBC As New TPCBarcode(p_oRec.Barcode)
        Dim oOP As New TPCOperatorBC(p_oRec.OperatorId)
        Dim oPT As New TPCPointBC(p_oRec.PointId)

        ' Declare Variables That Will Differ Based on Value of p_oRow
        Dim sToCity, sParcelType, sToLocID, sToAddID, sToLocName, sFromAddID, sFromCustID, sFromCustName, sFromLocID, sFromLocName As String
        ' Declare Variables That Will be based on MobileApp data if 'DD' which is assumed to come from EZScan
        Dim sDeliveryComments As String = Nothing
        Dim sSignature As String = Nothing
        ' Declare Variables That Will Be Relevant if 
        Dim sEmpty As String = String.Empty

        ' Declare Utility Variables
        Dim sb As New StringBuilder

        ' Initialzie Variables based on Value of p_oRow
        If Not IsNothing(p_oRow) Then

            sToCity = p_oRow("ToCity")
            sParcelType = p_oRow("ParcelType")
            sToLocID = p_oRow("ToLocID")
            sToAddID = p_oRow("ToAddID")
            sToLocName = p_oRow("ToLocName")
            sFromAddID = p_oRow("FromAddID")
            sFromCustID = p_oRow("FromCustID")
            sFromCustName = p_oRow("FromCustName")
            sFromLocID = p_oRow("FromLocID")
            sFromLocName = p_oRow("FromLocName")

            '' Add EZ-Scan Data if Applicable
            'If String.Compare(p_oRec.EventCode, "DD") = 0 Then
            '    AppendEzScanData(p_oRec, sDeliveryComments, sSignature)
            'End If

            '' Add Support for Branch Delivery which will contain the route name in Delivery Comments field
            'If String.Compare(p_oRec.EventCode, "BD") = 0 Then
            '    sDeliveryComments = p_oRec.DeliveryComments
            'End If

        Else

            sToCity = sEmpty
            sParcelType = sEmpty
            sToLocID = sEmpty
            sToAddID = "NULL"
            sToLocName = sEmpty
            sFromAddID = "NULL"
            sFromCustID = sEmpty
            sFromCustName = sEmpty
            sFromLocID = sEmpty
            sFromLocName = sEmpty
            sDeliveryComments = sEmpty
            sSignature = sEmpty

        End If

        ' Add EZ-Scan Data if Applicable
        If String.Compare(p_oRec.EventCode, "DD") = 0 Then
            AppendEzScanData(p_oRec, sDeliveryComments, sSignature)
        End If

        ' Add Support for Branch Delivery which will contain the route name in Delivery Comments field
        If String.Compare(p_oRec.EventCode, "BD") = 0 Then
            sDeliveryComments = p_oRec.DeliveryComments
        End If

        ' Construct the Insert Statement
        sb.Append("Insert into ")
        sb.Append(TRCTblPath)
        sb.Append("Event ")
        sb.Append("(EventCode, ScanDate, OperatorID, PointID, TicketNum, TrackingNum, ThirdPartyBarcode, ")
        sb.Append("ContainerBarcode, DeliveryOption, DeliveryComments, ToCity, ParcelType, Weight, Pieces, Void, ToLocID, ToAddID, ")
        sb.Append("ToLocName, RefNum, FromAddID, FromCustID, FromCustName, FromLocID, FromLocName, HHid, BatchNum, SignaturePath) ")
        sb.Append("VALUES ('")
        sb.Append(p_oRec.EventCode)
        sb.Append("', '")
        sb.Append(p_oRec.TimeStamp)
        sb.Append("', '")
        sb.Append(p_oRec.OperatorId)
        sb.Append("', '")
        sb.Append(p_oRec.PointId)
        sb.Append("', '', '") 'TicketNum is Empty
        sb.Append(oBC.Barcode)
        sb.Append("', '', '', '', '") 'ThirdPartyBarcode, ContainerBarcode, DeliveryOption
        sb.Append(sDeliveryComments)
        sb.Append("', '")
        sb.Append(sToCity)
        sb.Append("', '")
        sb.Append(sParcelType)
        sb.Append("', ")
        sb.Append(p_oRec.Weight)
        sb.Append(", 1, 'F', '") ' Pieces & Void are hard-coded to default values
        sb.Append(sToLocID)
        sb.Append("', ")
        sb.Append(sToAddID)
        sb.Append(", '")
        sb.Append(sToLocName)
        sb.Append("',NULL, ") ' RefNum is set to NULL
        sb.Append(sFromAddID)
        sb.Append(", '")
        sb.Append(sFromCustID)
        sb.Append("', '")
        sb.Append(sFromCustName)
        sb.Append("', '")
        sb.Append(sFromLocID)
        sb.Append("', '")
        sb.Append(sFromLocName)
        sb.Append("', '")
        sb.Append(p_oRec.HHid)
        sb.Append("', ")
        sb.Append(p_oRec.BatchID)
        sb.Append(", '")
        sb.Append(sSignature)
        sb.Append("')")

        Return sb.ToString()

    End Function

    Private Function InsertEventRecord(ByVal p_oRec As ScanRecord, Optional ByVal p_oRow As DataRow = Nothing) As Boolean

        ' Execute the Insert
        Dim sInsert As String = EventTableInsertStatement(p_oRec, p_oRow)
        Return ExecuteQuery(sInsert)

    End Function

    Private Function ManifestRowId(ByVal p_sBarcode As String, ByVal p_dTimeStamp As Date) As String

        Dim sReturn As String = Nothing
        Dim sb As StringBuilder

        Try

            sb = New StringBuilder

            If p_dTimeStamp.Month < 10 Then sb.Append("0")
            sb.Append(p_dTimeStamp.Month.ToString())

            If p_dTimeStamp.Day < 10 Then sb.Append("0")
            sb.Append(p_dTimeStamp.Day.ToString())

            sb.Append(p_dTimeStamp.Year.ToString())

            If p_dTimeStamp.Hour < 10 Then sb.Append("0")
            sb.Append(p_dTimeStamp.Hour.ToString())

            If p_dTimeStamp.Minute < 10 Then sb.Append("0")
            sb.Append(p_dTimeStamp.Minute.ToString())

            If p_dTimeStamp.Second < 10 Then sb.Append("0")
            sb.Append(p_dTimeStamp.Second.ToString())

            sb.Append(p_sBarcode)

            sReturn = sb.ToString()

        Catch ex As Exception

            Console.WriteLine("[Module1.2789]" & ex.Message)
            sReturn = String.Empty

        End Try

        Return sReturn

    End Function

    Private Function ContractWeight(ByVal p_dWeight As Double, Optional ByVal p_dFractionThereof As Double = 0.2D) As Double

        Dim dReturn As Double

        ' Round the weight up or down for billing purposes.  Rule: If decimal portion of weight is > 0.25, round up, otherwide round down
        Dim iPortion As Integer = IntegerPortion(p_dWeight)
        Dim fPortion As Decimal = DecimalPortion(p_dWeight)

        If fPortion > p_dFractionThereof Then
            dReturn = iPortion + 1 'Round UP
        Else
            dReturn = iPortion ' Round Down
        End If

        Return dReturn

    End Function

    Private Function ManifestTablesUpdateStatement(ByVal p_sRowID As String, ByVal p_dWeight As Double, ByVal p_sTableName As String) As String

        Dim sReturn As String = String.Empty
        Dim sb As StringBuilder

        Try

            sb = New StringBuilder

            sb.Append("UPDATE ")
            sb.Append(TRCTblPath)
            sb.Append(p_sTableName)
            sb.Append(" SET Weight = ")
            sb.Append(ContractWeight(p_dWeight).ToString())
            sb.Append(" WHERE RowId = '")
            sb.Append(p_sRowID)
            sb.Append("'")

            sReturn = sb.ToString()

        Catch ex As Exception

            Console.WriteLine("[Module1.2838]" & ex.Message)
            sReturn = String.Empty

        End Try

        Return sReturn

    End Function

    Private Function InsertEventAndUpdateManifest(ByVal p_oRec As ScanRecord, ByVal p_sRowID As String) As Boolean

        Dim bReturn As Boolean = False
        Dim sInsertEventTable As String = String.Empty
        Dim sUpdateManifestTable As String = String.Empty
        Dim sUpdateManifestInvoiceTable As String = String.Empty
        Dim sUpdateSettlementInvoiceTable As String = String.Empty
        Dim oConn As SqlConnection = Nothing
        Dim sTran As SqlTransaction = Nothing
        Dim oCmd As SqlCommand = Nothing

        Try



            'Prepare SQL for Event table Insert
            sInsertEventTable = EventTableInsertStatement(p_oRec)

            'Prepare SQL for Manifest table Update
            sUpdateManifestTable = ManifestTablesUpdateStatement(p_sRowID, p_oRec.Weight, "MANIFEST")

            'Prepare SQL for ManifestInvoice table Update
            sUpdateManifestInvoiceTable = ManifestTablesUpdateStatement(p_sRowID, p_oRec.Weight, "MANIFESTINVOICE")

            'Prepare SQL for SettlementInvoice table Update
            sUpdateSettlementInvoiceTable = ManifestTablesUpdateStatement(p_sRowID, p_oRec.Weight, "SETTLEMENTINVOICE")

            'Execute the transaction
            oConn = New SqlConnection(sScanListConnection)
            oConn.Open()

            sTran = oConn.BeginTransaction()

            If String.Compare(p_oRec.EventCode, "WC") <> 0 Then

                If New SqlCommand(sInsertEventTable, oConn, sTran).ExecuteNonQuery() = 1 Then
                    sTran.Commit()
                    bReturn = True
                Else
                    sTran.Rollback()
                    bReturn = False
                End If

            Else

                If New SqlCommand(sInsertEventTable, oConn, sTran).ExecuteNonQuery() = 1 Then

                    If New SqlCommand(sUpdateManifestTable, oConn, sTran).ExecuteNonQuery() = 1 Then

                        If New SqlCommand(sUpdateManifestInvoiceTable, oConn, sTran).ExecuteNonQuery() = 1 Then

                            If New SqlCommand(sUpdateSettlementInvoiceTable, oConn, sTran).ExecuteNonQuery() = 1 Then

                                sTran.Commit()
                                bReturn = True

                            Else

                                'If String.Compare(p_oRec.EventCode, "WC") = 0 Then
                                sTran.Rollback()
                                bReturn = False
                                'Else
                                '    sTran.Commit()
                                '    bReturn = True
                                'End If

                            End If

                        Else

                            'If String.Compare(p_oRec.EventCode, "WC") = 0 Then
                            sTran.Rollback()
                            bReturn = False
                            'Else
                            '    sTran.Commit()
                            '    bReturn = True
                            'End If

                        End If

                    Else

                        'If String.Compare(p_oRec.EventCode, "WC") = 0 Then
                        sTran.Rollback()
                        bReturn = False
                        'Else
                        '    sTran.Commit()
                        '    bReturn = True
                        'End If

                    End If

                Else

                    sTran.Rollback()
                    bReturn = False

                End If

            End If

        Catch ex As Exception

            Console.WriteLine("[Module1.2919]" & ex.Message)
            If Not sTran Is Nothing Then sTran.Rollback()
            If Not oConn Is Nothing Then oConn.Close()
            bReturn = False

        Finally

            If Not oConn Is Nothing Then
                oConn.Close()
                oConn.Dispose()
            End If

        End Try

        Return bReturn

    End Function

    '=====================================================================================================================
    '=====Function from GlobalVars.vb of RoutesModule2 Project============================================================
    '=====================================================================================================================
    Public KeyWords() As String = {" WITH ", " Where ", " Order ", " Group "}

    Public Function PopulateDataset2(ByRef xDataAdapter As SqlDataAdapter, ByRef dsData As DataSet, ByVal strSQL As String, Optional ByVal PreserveTbl As Boolean = False) As DataSet
        'Dim TblIndex As Integer
        Dim TblString As String
        Dim EndTblIndex As Integer = 0
        'Dim i As Integer
        Dim TblArray As Object()()
        Dim DataAdapter As SqlDataAdapter = Nothing

        PopulateDataset2 = Nothing
        If strSQL.Trim = "" Then Exit Function

        Dim localConn As New SqlConnection(sScanListConnection)

        PopulateDataset2 = Nothing
        If DataAdapter Is Nothing Then
            DataAdapter = New SqlDataAdapter
        End If
        If dsData Is Nothing Then
            dsData = New DataSet
        End If

        TblArray = TablesList(strSQL)
        If TblArray Is Nothing Then
            'Message modified by Michael Pastor
            ''MsgBox("PopulateDataset2: Cannot separate Table Names!", MsgBoxStyle.Exclamation, "Data Unavailable")
            Console.WriteLine("PopulateDataset2: Cannot separate Table Names!")
            '- MsgBox("PopulateDataset2: Cannot separate Table Names!")
            Exit Function
        End If

        If TblArray(0).Length > 1 Then
            TblString = TblArray(0)(1)
        Else
            TblString = TblArray(0)(0) 'Temp coding!!
        End If

        Try
            'Dim sqdtAdapter As New SqlDataAdapter(strSQL, localConn)
            DataAdapter.SelectCommand = New SqlCommand
            With DataAdapter.SelectCommand
                .Connection = localConn
                .CommandTimeout = 120
                .CommandText = strSQL
                .CommandType = CommandType.Text
            End With
            With DataAdapter
                .AcceptChangesDuringFill = True
                .MissingSchemaAction = MissingSchemaAction.AddWithKey
                If .TableMappings.Count <= 0 Then
                    .TableMappings.Add("Table", TblString)
                End If
                localConn.Open()
                If dsData.Tables.Count > 0 And PreserveTbl = False Then
                    'dsData.Tables(TblString).Clear()
                    If Not dsData.Relations Is Nothing Then
                        dsData.Relations.Clear()
                    End If
                    Dim tmpTable As DataTable
                    For Each tmpTable In dsData.Tables
                        tmpTable.Constraints.Clear()
                    Next
                    dsData.Tables.Clear()
                End If
                .Fill(dsData, TblString)
            End With
        Catch ex As System.Data.SqlClient.SqlException
            'Message NOT modified by Michael Pastor, due to format being identical to modified version.
            ''MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Company Profile")
            Console.WriteLine("[Module1.3113] Error: " & ex.Message, "CompanyProfile")
            PopulateDataset2 = Nothing
            Exit Function
        Finally
            localConn.Close()
            localConn = Nothing
            PopulateDataset2 = dsData

        End Try
    End Function

    Public Function TablesList(ByVal SqlString As String) As Object()()
        On Error GoTo ErrTrap

        Dim TblIndex As Integer
        Dim TblString, TblList() As String
        Dim EndTblIndex As Integer = 0
        Dim i As Integer
        'Dim TableAliasList() As Object


        TblIndex = InStr(SqlString, " From ", CompareMethod.Text) + Len(" From ") - 1
        For i = 0 To KeyWords.GetUpperBound(0)
            EndTblIndex = InStr(SqlString, KeyWords(i), CompareMethod.Text)
            If EndTblIndex > 0 Then Exit For
        Next i
        If EndTblIndex <= 0 Then
            TblString = SqlString.Substring(TblIndex)
        Else
            TblString = SqlString.Substring(TblIndex, EndTblIndex - TblIndex)
        End If
        TblString = TblString.Trim().ToUpper

        If TblString = "" Then
            'Message modified by Michael Pastor
            ''MsgBox("Table cannot be found.", MsgBoxStyle.Information, "Data Unavailable")
            Console.WriteLine("Module1.3149 Table cannot be found.", "Data Unavailable")
            '- MsgBox("Error: Table not found")
            Exit Function
        End If

        TblIndex = 0 : EndTblIndex = 0

        If TblString.IndexOf(" JOIN ") <= 0 Then
            TblList = TblString.Split(",")
        Else
            Dim TempTblStr As String
            Dim TempTblArr() As String
            Dim TempONArr() As String
            'Dim TempTblStr, TempTblArr(), TempONArr() As String
            Dim JoinDelims() As String = {"JOIN", "ON"}

            TempTblStr = TblString.Replace("(", "")
            TempTblStr = TempTblStr.Replace(")", "")
            TempTblStr = TempTblStr.ToUpper
            TempTblStr = TempTblStr.Replace(" JOIN ", " | ")
            TempTblArr = TempTblStr.Split("|")
            For i = 0 To TempTblArr.Length - 1
                TempTblArr(i) = TempTblArr(i).Replace(" LEFT ", " ")
                TempTblArr(i) = TempTblArr(i).Replace(" OUTER ", " ")
                TempTblArr(i) = TempTblArr(i).Replace(" INNER ", " ")
                TempTblArr(i) = TempTblArr(i).Replace(" CROSS ", " ")
                TempTblArr(i) = TempTblArr(i).Replace(" ON ", " | ")
                TempONArr = TempTblArr(i).Split("|")
                TempTblArr(i) = TempONArr(0)
                TempTblArr(i) = TempTblArr(i).Trim

                'TempONArr.Clear(TempONArr, 0, TempONArr.Length)
                Array.Clear(TempONArr, 0, TempONArr.Length)
                TempONArr = Nothing
            Next
            TblList = TempTblArr
        End If
        'Dim TableAliasList()() = {New String(TblList.GetUpperBound(0)) {}}
        Dim TableAliasList(TblList.GetUpperBound(0))() As String
        Dim Sales()() As Double = {New Double(11) {}}
        'Dim TableAliasList()() As Array
        'TableAliasList = New String(TblList.GetUpperBound(0)) {}

        For i = 0 To TblList.GetUpperBound(0)
            'TableAliasList(i) = New Array(1)
            TableAliasList(i) = TblList(i).Trim.Split(" ")
            'EndTblIndex = TblList(i).IndexOf(" ", TblIndex) 'Check for Alias
            'If EndTblIndex > 0 Then
            '    TblString = TblString.Substring(0, EndTblIndex).Trim
            'End If
        Next i
        TablesList = TableAliasList
        Exit Function
ErrTrap:
        'Message modified by Michael Pastor
        ''MsgBox("TablesList Error: " & Err.Description, MsgBoxStyle.Critical, "Critical Error")
        Console.WriteLine("Module1.3201 TablesList Error: " & Err.Description, "Critical Error")
        '- MsgBox("TablesList Error: " & Err.Description)
    End Function

    '    Public Function ExecuteQuery(ByVal Query As String, Optional ByVal cmdSQLTrans As SqlCommand = Nothing, Optional ByVal CloseConn As Boolean = True) As Boolean

    '        Dim Conn As SqlConnection '(sScanListConnection)
    '        Dim cmd As SqlCommand '= New SqlCommand(Query, Conn)
    '        Dim HadErr As Boolean = False

    '        If Query = "" Then

    '            ExecuteQuery = True
    '            GoTo Release

    '        Else

    '            ExecuteQuery = False

    '        End If

    '        ExecuteQuery = False

    '        FixSingleQuote(Query, True)

    '        If cmdSQLTrans Is Nothing Then

    '            Conn = New SqlConnection(sScanListConnection)
    '            cmd = New SqlCommand(Query, Conn)

    '        Else

    '            Conn = cmdSQLTrans.Connection
    '            cmd = cmdSQLTrans
    '            cmd.CommandText = Query

    '        End If

    '        Try

    '            If cmd.Connection.State <> ConnectionState.Open Then
    '                Conn.Open()
    '            End If
    '            cmd.ExecuteNonQuery()

    '        Catch ex As System.Data.SqlClient.SqlException
    '            'Message modified by Michael Pastor
    '            ''MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Critical Error")
    '            Console.WriteLine("[Module1.3249] Error: " & ex.Message, "Critical Error")
    '            '- MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
    '            HadErr = True
    '            GoTo Release

    '        Finally

    '            If HadErr = False Then
    '                ExecuteQuery = True
    '            End If
    '            'close the connection
    '        End Try

    'Release:
    '        If Not Conn Is Nothing Then
    '            If CloseConn = True And Conn.State = ConnectionState.Open Then
    '                If Not cmd.Transaction Is Nothing Then
    '                    If HadErr Then
    '                        cmd.Transaction.Rollback()
    '                    Else
    '                        cmd.Transaction.Commit()
    '                    End If
    '                End If
    '                Conn.Close()
    '                cmd = Nothing
    '            End If
    '        End If
    '        Conn = Nothing
    '    End Function

    Public Function ExecuteQuery(ByVal Query As String, Optional ByVal cmdSQLTrans As SqlCommand = Nothing, Optional ByVal CloseConn As Boolean = True) As Boolean

        Dim Conn As SqlConnection = Nothing  '(sScanListConnection)
        Dim cmd As SqlCommand = Nothing '= New SqlCommand(Query, Conn)
        Dim bHadError As Boolean = False
        Dim bReturn As Boolean = True

        Try

            If Query <> "" Then

                FixSingleQuote(Query, True)

                If cmdSQLTrans Is Nothing Then

                    Conn = New SqlConnection(sScanListConnection)
                    cmd = New SqlCommand(Query, Conn)

                Else

                    Conn = cmdSQLTrans.Connection
                    cmd = cmdSQLTrans
                    cmd.CommandText = Query

                End If

                If cmd.Connection.State <> ConnectionState.Open Then
                    Conn.Open()
                End If

                'Console.WriteLine("<DEBUG>")
                'Console.WriteLine(Query)
                cmd.ExecuteNonQuery()
                'Console.WriteLine("</DEBUG>")

            End If

        Catch ex As System.Data.SqlClient.SqlException

            Console.WriteLine("[Module1.3315] Error: " & ex.Message, "Critical Error")
            bHadError = True
            bReturn = False

        Catch ex As Exception

            Console.WriteLine("[Module1.3321] Error: " & ex.Message, "Critical Error")
            bHadError = True
            bReturn = False

        Finally

            'close the connection
            If Not Conn Is Nothing Then

                If CloseConn = True And Conn.State = ConnectionState.Open Then

                    If Not cmd.Transaction Is Nothing Then

                        If bHadError Then

                            cmd.Transaction.Rollback()

                        Else

                            cmd.Transaction.Commit()

                        End If

                    End If

                    Conn.Close()

                End If

                cmd = Nothing

            End If

        End Try

        Return bReturn

    End Function

    Public Sub FixSingleQuote(ByRef Query As String, Optional ByVal IsSQLQry As Boolean = False)
        Dim StrArr() As Char
        Dim i As Int32

        Exit Sub

        i = 0
        If IsSQLQry Then

            Dim j, Found, TotalFnd, Pos(Query.Length - 1), PosCorr, WHEREIndex As Int32
            Query = Query.ToUpper
            WHEREIndex = Query.IndexOf(" WHERE ")
            Found = 0
            TotalFnd = 0
            StrArr = Query.ToCharArray(0, Query.Length)
            If WHEREIndex <= 0 Then
                WHEREIndex = StrArr.Length - 1
            End If
            For i = 0 To WHEREIndex

                Select Case StrArr(i)
                    Case "'"
                        Select Case Found
                            Case 0
                                Found = 1
                                TotalFnd += 1
                                Pos(TotalFnd - 1) = i
                            Case 1
                                Found = 2
                                TotalFnd += 1
                                Pos(TotalFnd - 1) = i
                            Case 2
                                TotalFnd += 1
                                Pos(TotalFnd - 1) = i

                                For j = i + 1 To StrArr.Length - 1
                                    If StrArr(j) = "'" Then

                                    End If
                                    If StrArr(j) <> " " And StrArr(j) <> "," Then
                                        Query = Query.Insert(Pos(TotalFnd - 2) + PosCorr, "'")
                                        PosCorr += 1
                                        Found = 1
                                        Exit For
                                    End If
                                Next
                                'If StrArr(i - 1) = "'" Then
                                '    Found = 3
                                '    Pos(TotalFnd - 1) = i
                                'Else
                                '    Query = Query.Insert(Pos(TotalFnd - 2) + PosCorr, "'")
                                '    PosCorr += 1
                                '    Found = 2
                                'End If
                        End Select
                    Case ","
                        Select Case Found
                            Case 0
                                'Nothing
                            Case 1
                                ' Part of Text, Carry on
                            Case 2
                                Found = 0

                        End Select
                End Select
            Next

            WHEREIndex = Query.IndexOf(" WHERE ")
            If WHEREIndex > 0 Then
                Dim WhereParts() As String
                Dim WhereClause, xQuery As String
                Dim AndArr() As Char = {" ", "A", "N", "D", " "}
                'Dim j As Int32

                WhereClause = Query.Substring(WHEREIndex)
                WhereParts = WhereClause.Split(AndArr) '(" AND ")
                xQuery = Query.Substring(0, WHEREIndex)
                For j = 0 To WhereParts.Length - 1
                    StrArr = WhereParts(j).ToCharArray
                    Found = 0
                    TotalFnd = 0
                    'Pos.Clear(Pos, 0, Pos.Length)
                    Array.Clear(Pos, 0, Pos.Length)
                    PosCorr = 0
                    For i = 0 To StrArr.Length - 1

                        Select Case StrArr(i)
                            Case "'"
                                Select Case Found
                                    Case 0
                                        Found = 1
                                        TotalFnd += 1
                                        Pos(TotalFnd - 1) = i
                                    Case 1
                                        Found = 2
                                        TotalFnd += 1
                                        Pos(TotalFnd - 1) = i
                                    Case 2
                                        TotalFnd += 1
                                        Pos(TotalFnd - 1) = i
                                        If StrArr(i - 1) = "'" Then
                                            Found = 1
                                            'Pos(TotalFnd - 1) = i
                                        Else
                                            WhereParts(j) = WhereParts(j).Insert(Pos(TotalFnd - 2) + PosCorr, "'")
                                            PosCorr += 1
                                            Found = 2
                                        End If
                                End Select
                        End Select
                    Next i
                    If Found = 1 Then
                        WhereParts(j) = WhereParts(j).Insert(Pos(TotalFnd - 2) + PosCorr, "'")
                        PosCorr += 1
                        Found = 0
                    End If
                    xQuery = xQuery & WhereParts(j) & " AND "
                Next j
                xQuery = xQuery.Substring(0, xQuery.Length - 1 - Len(" AND "))
                Query = xQuery
            End If 'WHERE Clause

        Else

            While i <> -1
                i = Query.IndexOf("'S", i)
                'If i = 0 Then Exit While
                If i >= 1 Then
                    If Query.Substring(i - 1, 1) <> "'" Then
                        Query = Query.Insert(i, "'")
                        i += 2
                    Else
                    End If
                End If
            End While

            i = 0
            While i <> -1
                i = Query.IndexOf("S'", i)
                If i = -1 Then Exit While
                If i > 1 Then
                    If Query.Substring(i + 1, 1) <> "'" Then
                        Query = Query.Insert(i, "'")
                        i += 2
                    Else ' Could be '.... Jones'' (last ' for termination)
                        If IsSQLQry Then
                            If (i + 2) >= Query.Length Then
                                Query = Query.Insert(i, "'")
                                'i += 2
                                Exit While
                            End If
                        End If
                    End If
                End If
            End While
        End If


    End Sub

    '===================   Import Functions    =============================
    Private Function DailyEntryExists(ByVal p_sShortDate As String, ByVal p_iWeightPlanID As Integer) As Boolean

        Try
            ' Data Access Variables
            Dim oDataAdapter As SqlDataAdapter = Nothing
            Dim oDataSet As DataSet = Nothing
            'Dim oDailyEntry As DataRow
            Dim iRowCount As Integer

            ' Utility Variables
            Dim sb As New StringBuilder
            Dim sCmd As String

            ' Check to see if there is already a row for this record.  The primary key is {TranDate,ManifestID}.
            sb.Append("SELECT COUNT(*) AS RowsFound FROM ")
            sb.Append(WEIGHTTblPath)
            sb.Append("DAILYENTRY WHERE TranDate = '")
            sb.Append(p_sShortDate)
            sb.Append("' and ManifestID = '")
            sb.Append(p_iWeightPlanID)
            sb.Append("'")

            sCmd = sb.ToString()

            PopulateDataset2(oDataAdapter, oDataSet, sCmd)

            iRowCount = oDataSet.Tables(0).Rows(0).Item("RowsFound")

            If iRowCount = 1 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception

            ''MessageBox.Show(ex.Message)
            Console.WriteLine("[Module1.3558]" & ex.Message)
            Return False

        End Try

    End Function

    Private Function GetParentWeightPlan(ByVal p_iWeightPlanID As Integer) As Integer

        Try
            ' Data Access Variables
            Dim oDataAdapter As SqlDataAdapter = Nothing
            Dim oDataSet As DataSet = Nothing
            'Dim oDailyEntry As DataRow
            Dim iRowCount As Integer

            ' Utility Variables
            Dim sb As New StringBuilder
            Dim sCmd As String

            ' Check to see if this weight plan has a parent
            sb.Append("select case when ISNUMERIC(ParentID) = 1 then ParentID else 0 end as ParentId from ")
            sb.Append(WEIGHTTblPath)
            sb.Append("Manifests where [id] = ")
            sb.Append(p_iWeightPlanID)

            sCmd = sb.ToString()

            PopulateDataset2(oDataAdapter, oDataSet, sCmd)

            iRowCount = oDataSet.Tables(0).Rows.Count

            If iRowCount = 1 Then
                Return oDataSet.Tables(0).Rows(0).Item("ParentID")
            Else
                Return 0
            End If

        Catch ex As Exception

            ''MessageBox.Show(ex.Message)
            Console.WriteLine("[Module1.3599]" & ex.Message)
            Return False

        End Try

    End Function


    Private Function DailyEntryUpdateString(ByVal p_sShortDate As String, ByVal p_iWeightPlanID As Integer, ByVal p_fWeight As Decimal) As String

        Try
            ' Data Access Variables
            Dim oDataAdapter As SqlDataAdapter = Nothing
            Dim oDataSet As DataSet = Nothing
            Dim oDataRow As DataRow
            Dim iRowCount As Integer

            ' Command Component Variables
            Dim fOldWeight, fNewWeight, fWeightLimit, fOverCharge, fNewCharge, fChargableWeight As Decimal

            ' Utility Variables
            Dim sb As New StringBuilder
            Dim sCmd As String

            ' Get the current weight, limit and overcharge for this record
            sb.Append("SELECT Weight, WeightLimit, OWCharge FROM ")
            sb.Append(WEIGHTTblPath)
            sb.Append("DAILYENTRY WHERE TranDate = '")
            sb.Append(p_sShortDate)
            sb.Append("' and ManifestID = '")
            sb.Append(p_iWeightPlanID)
            sb.Append("'")

            sCmd = sb.ToString()
            sb.Remove(0, sb.Length)

            PopulateDataset2(oDataAdapter, oDataSet, sCmd)

            iRowCount = oDataSet.Tables(0).Rows.Count

            If iRowCount = 1 Then

                oDataRow = oDataSet.Tables(0).Rows(0)
                fOverCharge = oDataRow.Item("OWCharge")
                fOldWeight = oDataRow.Item("Weight")
                fWeightLimit = oDataRow.Item("WeightLimit")

                fNewWeight = fOldWeight + p_fWeight
                fChargableWeight = fNewWeight - fWeightLimit
                If fChargableWeight > 0 Then
                    fNewCharge = fChargableWeight * fOverCharge
                Else
                    fNewCharge = 0
                End If

                sb.Append("UPDATE ")
                sb.Append(WEIGHTTblPath)
                sb.Append("DailyEntry SET Weight = ")
                sb.Append(fNewWeight)
                sb.Append(", Charge = ")
                sb.Append(fNewCharge)
                sb.Append(" WHERE TranDate = '")
                sb.Append(p_sShortDate)
                sb.Append("' and ManifestID = '")
                sb.Append(p_iWeightPlanID)
                sb.Append("'")

                Return sb.ToString()

            Else
                Return String.Empty
            End If

        Catch ex As Exception

            ''MessageBox.Show(ex.Message)
            Console.WriteLine("[Module1.3675]" & ex.Message)
            Return String.Empty

        End Try

    End Function

    Private Function DailyEntryInsertString(ByVal p_sTranDate As String, ByVal p_iWeightPlanID As Integer, ByVal p_fWeight As Decimal) As String

        Dim sb As New StringBuilder

        ' Construct the Insert Statement
        sb.Append("INSERT INTO ")
        sb.Append(WEIGHTTblPath)
        sb.Append("DailyEntry (TranDate,ManifestID,OfficeID,AccountID,AccountName,ManifestName,[Weight],WeightLimit,OWCharge,Charge,Finalize,WeightPlanGroupID,WeightPlanGroup,ParentID,[Invoice No]) SELECT '")
        sb.Append(p_sTranDate)
        sb.Append("' as TranDate, ")
        sb.Append("m.[ID] as ManifestID,")
        sb.Append("m.OfficeID as OfficeID,")
        sb.Append("m.AccountID as AccountID,")
        sb.Append("c.[name] as AccountName,")
        sb.Append("m.[Name] as ManifestName,")
        sb.Append(p_fWeight)
        sb.Append(" as Weight,")
        sb.Append("wbd.WeightLimit as WeightLimit,")
        sb.Append("wbd.OWCharge as OWCharge,")
        sb.Append("ROUND(((")
        sb.Append(p_fWeight)
        sb.Append(" - wbd.WeightLimit) + ABS(")
        sb.Append(p_fWeight)
        sb.Append(" - wbd.WeightLimit)) / 2 * wbd.OWCharge,2) as Charge,")
        sb.Append("0 as Finalize,")
        sb.Append("wpg.[id] as WeightPlanGroupID,")
        sb.Append("wpg.[name] as WeightPlanGroup,")
        sb.Append("m.ParentID as ParentID,")
        sb.Append("0 as [Invoice No] ")
        sb.Append("from	")
        sb.Append(WEIGHTTblPath)
        sb.Append("manifests m,")
        sb.Append(WEIGHTTblPath)
        sb.Append("WeightBreakdown wbd,")
        sb.Append(WEIGHTTblPath)
        sb.Append("WeightPlanGroups wpg,")
        sb.Append(AppTblPath)
        sb.Append("Customer c ")
        sb.Append("where	m.[id] = ")
        sb.Append(p_iWeightPlanID)
        sb.Append(" and wbd.[id] = m.WeightID and wpg.[id] = m.GroupID and c.[id] = m.AccountID")

        Return sb.ToString()

    End Function

    'Private Sub ImportBluePackageManifests()

    '    Dim bResult As Boolean = False
    '    Dim oConn As SqlConnection = Nothing
    '    Dim oCmd As SqlCommand = Nothing
    '    Dim oTran As SqlTransaction = Nothing
    '    Dim oPickupManifest As PickupManifest = Nothing

    '    Dim sFiles() As String = Nothing
    '    Dim sFolders() As String = Nothing
    '    Dim sValidFiles() As String = Nothing
    '    Dim sValidFilesPath() As String = Nothing

    '    Dim sFileNameParts() As String = Nothing
    '    Dim iParts As Integer = 0

    '    Dim sFileName As String = Nothing
    '    Dim iFileName As Integer = 0
    '    Dim iFirstThree As Integer = 0

    '    Dim sFileExtension As String = Nothing

    '    Dim sTmp() As String = Nothing
    '    Dim i As Int32 = 0
    '    Dim j As Int32 = 0

    '    Try

    '        'Get List of Files that Must be Processed
    '        sBluePackagePath = sBluePackagePath.ToUpper
    '        If Directory.Exists(sBluePackagePath) Then
    '            sFiles = Directory.GetFiles(sBluePackagePath)
    '            For i = 0 To sFiles.Length - 1

    '                sFiles(i) = sFiles(i).ToUpper

    '                sFileNameParts = sFiles(i).Split(".")
    '                iParts = sFileNameParts.Length

    '                sFolders = sFileNameParts(iParts - 2).Split("\")

    '                sFileName = sFolders(sFolders.Length - 1)
    '                sFileExtension = sFileNameParts(iParts - 1)

    '                iFileName = sFileName.Length

    '                If IsNumeric(sFileName.Substring(0, 8)) And (iFileName = 11) And (sFileExtension.CompareTo("TXT") = 0) Then

    '                    ReDim Preserve sValidFiles(j)
    '                    ReDim Preserve sValidFilesPath(j)

    '                    Dim iNameLength As Integer = sFiles(i).Length
    '                    sValidFiles(j) = sFiles(i).Substring(iNameLength - 15, 15)
    '                    sValidFilesPath(j) = sFiles(i).Substring(0, iNameLength - 15)
    '                    j += 1

    '                End If
    '            Next
    '        Else
    '            Console.WriteLine("Path does not exist for Blue Package Manifests:" & sBluePackagePath)
    '            Exit Sub
    '        End If


    '        If Not sValidFiles Is Nothing Then

    '            Console.WriteLine("Blue Package Files Detected. Beginning Import.")

    '            ' We can use the same SqlConnection for the entire batch of files
    '            If oConn Is Nothing Then
    '                oConn = New SqlConnection(sScanListConnection)
    '                oConn.Open()
    '            End If

    '            ' Process Batch of Valid Files
    '            For i = 0 To sValidFiles.Length - 1

    '                ' Either the entire file is processed and renamed or the operation on that file fails
    '                oTran = oConn.BeginTransaction()
    '                oCmd = New SqlCommand("spInsertPuManifestRecordV2", oConn, oTran)
    '                oCmd.CommandType = CommandType.StoredProcedure

    '                ' Read in the next file
    '                oPickupManifest = New PickupManifestMapper(sValidFilesPath(i) & sValidFiles(i), "V3")

    '                If Not oPickupManifest.Records Is Nothing Then

    '                    bResult = InsertManifestRecords(oPickupManifest.Records, oCmd, oPickupManifest.FileVersion)
    '                    If bResult Then
    '                        sTmp = sValidFiles(i).Split(".")
    '                        sTmp(1) = sTmp(1) & "_ARC"

    '                        Dim sOldFileName As String = sValidFilesPath(i) & sValidFiles(i)
    '                        Dim sNewFileName As String = sValidFilesPath(i) & sTmp(0) & "." & sTmp(1)

    '                        'File.Move(sValidFiles(i), sTmp(0) & "." & sTmp(1))
    '                        File.Move(sOldFileName, sNewFileName)
    '                        sTmp = Nothing
    '                        Console.WriteLine(sValidFiles(i) & " was successfully processed")
    '                    Else
    '                        Console.WriteLine(sValidFiles(i) & " was NOT successfully processed")
    '                    End If

    '                Else

    '                    ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                    Console.WriteLine("[Module1.3834]" & oPickupManifest.ErrorMessage & ". " & sValidFiles(i) & "failed to process.")
    '                    bResult = False

    '                End If

    '                If bResult = True Then
    '                    oTran.Commit()
    '                Else
    '                    oTran.Rollback()
    '                End If

    '                oTran = Nothing
    '                oCmd = Nothing

    '            Next

    '        End If


    '    Catch ex As Exception

    '        Console.WriteLine(ex.Message, "ImportMedExpressManifests() Status")
    '        If Not oTran Is Nothing Then oTran.Rollback()

    '    Finally

    '        oCmd = Nothing
    '        oTran = Nothing

    '        If Not oConn Is Nothing Then
    '            oConn.Close()
    '            oConn.Dispose()
    '        End If

    '    End Try


    'End Sub

    Private Sub ImportAutoZoneManifests()

        Dim bResult As Boolean = False
        Dim oConn As SqlConnection = Nothing
        Dim oCmd As SqlCommand = Nothing
        Dim oTran As SqlTransaction = Nothing
        Dim oPickupManifest As PickupManifest = Nothing

        Dim sFiles() As String = Nothing
        Dim sFolders() As String = Nothing
        Dim sValidFiles() As String = Nothing
        Dim sValidFilesPath() As String = Nothing

        Dim sFileNameParts() As String = Nothing
        Dim iParts As Integer = 0

        Dim sFileName As String = Nothing
        Dim iFileName As Integer = 0
        Dim iFirstThree As Integer = 0

        Dim sFileExtension As String = Nothing

        Dim sTmp() As String = Nothing
        Dim i As Int32 = 0
        Dim j As Int32 = 0

        Try

            'Get List of Files that Must be Processed
            sAutoZonePath = sAutoZonePath.ToUpper
            If Directory.Exists(sAutoZonePath) Then
                sFiles = Directory.GetFiles(sAutoZonePath)
                For i = 0 To sFiles.Length - 1

                    sFiles(i) = sFiles(i).ToUpper

                    sFileNameParts = sFiles(i).Split(".")
                    iParts = sFileNameParts.Length

                    sFolders = sFileNameParts(iParts - 2).Split("\")

                    sFileName = sFolders(sFolders.Length - 1)
                    sFileExtension = sFileNameParts(iParts - 1)

                    iFileName = sFileName.Length

                    If IsNumeric(sFileName.Substring(0, 8)) And (iFileName = 11) And (sFileExtension.CompareTo("TXT") = 0) Then

                        ReDim Preserve sValidFiles(j)
                        ReDim Preserve sValidFilesPath(j)

                        Dim iNameLength As Integer = sFiles(i).Length
                        sValidFiles(j) = sFiles(i).Substring(iNameLength - 15, 15)
                        sValidFilesPath(j) = sFiles(i).Substring(0, iNameLength - 15)
                        j += 1

                    End If
                Next
            Else
                Console.WriteLine("Path does not exist for Blue Package Manifests:" & sAutoZonePath)
                Exit Sub
            End If


            If Not sValidFiles Is Nothing Then

                Console.WriteLine("Auto Zone Files Detected. Beginning Import.")

                ' We can use the same SqlConnection for the entire batch of files
                If oConn Is Nothing Then
                    oConn = New SqlConnection(sScanListConnection)
                    oConn.Open()
                End If

                ' Process Batch of Valid Files
                For i = 0 To sValidFiles.Length - 1

                    ' Either the entire file is processed and renamed or the operation on that file fails
                    oTran = oConn.BeginTransaction()
                    oCmd = New SqlCommand("spInsertPuManifestRecordV4", oConn, oTran)
                    oCmd.CommandType = CommandType.StoredProcedure

                    ' Read in the next file
                    oPickupManifest = New PickupManifestMapper(sValidFilesPath(i) & sValidFiles(i), "V4")

                    If Not oPickupManifest.Records Is Nothing Then

                        bResult = InsertManifestRecords(oPickupManifest.Records, oCmd, oPickupManifest.FileVersion)
                        If bResult Then
                            sTmp = sValidFiles(i).Split(".")
                            sTmp(1) = sTmp(1) & "_ARC"

                            Dim sOldFileName As String = sValidFilesPath(i) & sValidFiles(i)
                            Dim sNewFileName As String = sValidFilesPath(i) & sTmp(0) & "." & sTmp(1)

                            'File.Move(sValidFiles(i), sTmp(0) & "." & sTmp(1))
                            File.Move(sOldFileName, sNewFileName)
                            sTmp = Nothing
                            Console.WriteLine(sValidFiles(i) & " was successfully processed")
                        Else
                            Console.WriteLine(sValidFiles(i) & " was NOT successfully processed")
                        End If

                    Else

                        ''MessageBox.Show(oScanList.ErrorMessage, "Import ScanList Status", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Console.WriteLine("[Module1.3979]" & oPickupManifest.ErrorMessage & ". " & sValidFiles(i) & "failed to process.")
                        bResult = False

                    End If

                    If bResult = True Then
                        oTran.Commit()
                    Else
                        oTran.Rollback()
                    End If

                    oTran = Nothing
                    oCmd = Nothing

                Next

            End If

        Catch SqlEx As SqlException
            Dim myError As SqlError
            Console.WriteLine("Sql Errors Count:" & SqlEx.Errors.Count)
            For Each myError In SqlEx.Errors
                Console.WriteLine(myError.Number & " - " & myError.Message)
            Next
            If Not oTran Is Nothing Then oTran.Rollback()

        Catch ex As Exception

            Console.WriteLine(ex.Message, "ImportAutoZoneManifests() Status")
            If Not oTran Is Nothing Then oTran.Rollback()

        Finally

            oCmd = Nothing
            oTran = Nothing

            If Not oConn Is Nothing Then
                oConn.Close()
                oConn.Dispose()
            End If

        End Try


    End Sub

    Private Function FixLocationDatFile() As Boolean

        Dim sb As StringBuilder
        Dim oFileInfo As FileInfo
        Dim dtInspect As DateTime
        Dim oReader As StreamReader = Nothing
        Dim sReadFileName As String
        Dim oWriter As StreamWriter = Nothing
        Dim sWriteFileName As String
        Dim bUpdated As Boolean = False

        Try

            'Check to see if LOCATION.DAT file has been modified since last check
            sb = New StringBuilder
            sb.Append(sTrackingPath)
            sb.Append("\LOCATION.DAT")
            sReadFileName = sb.ToString()

            sb.Length = 0
            sb.Append(sTrackingPath)
            sb.Append("\LOCATION.DAT.TMP")
            sWriteFileName = sb.ToString()

            oFileInfo = New FileInfo(sReadFileName)
            dtInspect = oFileInfo.CreationTime

            'Replace "TPC,TPC" with ,TPC" if it has
            If dtInspect > dtLastLocationFix Then

                Console.Write("New version of Location.dat File Detected...")

                ' Open the files for reading & writing
                'oReader = New StreamReader(sReadFileName)
                oReader = File.OpenText(sReadFileName)
                'oWriter = New StreamWriter(sWriteFileName)
                oWriter = File.CreateText(sWriteFileName)

                Dim strCurrentLine As String

                Do
                    strCurrentLine = oReader.ReadLine()
                    If strCurrentLine Is Nothing Then
                        oWriter.Flush()
                        Exit Do
                    Else
                        strCurrentLine = strCurrentLine.Replace("TPC,TPC", ",TPC")
                        oWriter.WriteLine(strCurrentLine)
                    End If
                Loop

                'Replace DAT file with TMP file
                oReader.Close()
                oWriter.Flush()
                oWriter.Close()
                System.IO.File.Delete(sReadFileName)
                System.IO.File.Move(sWriteFileName, sReadFileName)

                dtLastLocationFix = Date.Now()
                bUpdated = True

            End If

            Return bUpdated

        Catch ex As Exception

        Finally

            If Not IsNothing(oReader) Then
                oReader.Close()
            End If

            If Not IsNothing(oWriter) Then
                oWriter.Close()
            End If

        End Try

    End Function

    Sub Main()

        Dim dt As DateTime
        Dim sParams As String() = Split(Command$, " ")

        If sParams(0).Length = 0 Then
            'MsgBox(sParams(0), MsgBoxStyle.OKOnly, "Command Line Parameters")
            Console.WriteLine("Please Specify the Partners Book Path to Monitor")
            Return
        Else
            sPartnerBooksPath = sParams(0)
            Console.WriteLine("Monitoring " & sPartnerBooksPath)
        End If

        If sParams(1).Length = 0 Then
            'MsgBox(sParams(1), MsgBoxStyle.OKOnly, "Command Line Parameters")
            Console.WriteLine("Please Specify the Partners Book Path to Monitor")
            Return
        Else
            sMedExpressPath = sParams(1)
            Console.WriteLine("Monitoring " & sMedExpressPath)
        End If

        If sParams(2).Length = 0 Then
            Console.WriteLine("Please Specify the Med Express Same-day Path to Monitor")
            Return
        Else
            sMedExpressSdPath = sParams(2)
            Console.WriteLine("Monitoring " & sMedExpressSdPath)
        End If

        If sParams(3).Length = 0 Then
            Console.WriteLine("Please Specify the AutoZone Path to Monitor")
            Return
        Else
            sAutoZonePath = sParams(3)
            Console.WriteLine("Monitoring " & sAutoZonePath)
        End If

        If sParams(4).Length = 0 Then
            Console.WriteLine("Please Specify the Tracking Path to Monitor")
            Return
        Else
            sTrackingPath = sParams(4)
            Console.WriteLine("Monitoring " & sTrackingPath)
        End If

        If sParams(5).Length = 0 Then
            Console.WriteLine("Please Specify the DIOCS Path to Monitor")
            Return
        Else
            sDio2CsPath = sParams(5)
            Console.WriteLine("Monitoring " & sDio2CsPath)
        End If

        'INACTIVE'If sParams(0).Length = 0 Then
        'INACTIVE'    'MsgBox(sParams(0), MsgBoxStyle.OKOnly, "Command Line Parameters")
        'INACTIVE'    Console.WriteLine("Please Specify the Ingram-Micro Path to Monitor")
        'INACTIVE'    Return
        'INACTIVE'Else
        'INACTIVE'    sIngramMicroPath = sParams(0)
        'INACTIVE'    Console.WriteLine("Monitoring " & sIngramMicroPath)
        'INACTIVE ''End If

        'INACTIVE'If sParams(4).Length = 0 Then
        'INACTIVE'    Console.WriteLine("Please Specify the Blue Package Path to Monitor")
        'INACTIVE'    Return
        'INACTIVE'Else
        'INACTIVE'    sBluePackagePath = sParams(4)
        'INACTIVE'    Console.WriteLine("Monitoring " & sBluePackagePath)
        'INACTIVE'End If
        'INACTIVE'If sParams(4).Length = 0 Then
        'INACTIVE'    Console.WriteLine("Please Specify the Pro Courier Path to Monitor")
        'INACTIVE'    Return
        'INACTIVE'Else
        'INACTIVE'    sProCourierPath = sParams(4)
        'INACTIVE'    Console.WriteLine("Monitoring " & sProCourierPath)
        'INACTIVE'End If

        'INACTIVE'If sParams(6).Length = 0 Then
        'INACTIVE'    Console.WriteLine("Please Specify the DSC Delivery Path to Monitor")
        'INACTIVE'    Return
        'INACTIVE'Else
        'INACTIVE'    sDCSDeliveryPath = sParams(6)
        'INACTIVE'    Console.WriteLine("Monitoring " & sDCSDeliveryPath)
        'INACTIVE'End If

        'INACTIVE'If sParams(8).Length = 0 Then
        'INACTIVE'    Console.WriteLine("Please Specify the PIA Path to Monitor")
        'INACTIVE'    Return
        'INACTIVE'Else
        'INACTIVE'    sPiaDeliveryPath = sParams(8)
        'INACTIVE'    Console.WriteLine("Monitoring " & sPiaDeliveryPath)
        'INACTIVE'End If


        Console.Write("Waiting...")

        While (1)

            If ImportScanList() = True Then
                dt = DateTime.Now
                Console.WriteLine("Import of Scan List Completed at {0}", dt.ToString("u"))
            Else
                '    dt = DateTime.Now
                Console.WriteLine("Problems Were Encontered During the Scan List Import at {0}", dt.ToString("u"))
            End If

            PurgeScanList() 'This will mark all non-importable records as Processed.  "non-importable records" are those scans with fatal errors such as 5 or 6, i.e. the scans are no good.

            If ImportEZScanListDD() = True Then
                dt = DateTime.Now
                Console.WriteLine("Import of Delivery EZ-Scan List Completed at {0}", dt.ToString("u"))
            Else
                '    dt = DateTime.Now
                Console.WriteLine("Problems Were Encontered During the Delivery EZ-Scan List Import at {0}", dt.ToString("u"))
            End If

            If ImportEZScanListAF() = True Then
                dt = DateTime.Now
                Console.WriteLine("Import of Pick-up EZ-Scan List Completed at {0}", dt.ToString("u"))
            Else
                '    dt = DateTime.Now
                Console.WriteLine("Problems Were Encontered During the Pick-up EZ-Scan List Import at {0}", dt.ToString("u"))
            End If

            'INACTIVE'If ImportSmsData() = True Then
            'INACTIVE'    dt = DateTime.Now
            'INACTIVE'    Console.WriteLine("Import of SMS Data Completed at (0)", dt.ToString("u"))
            'INACTIVE'Else
            'INACTIVE'    Console.WriteLine("Problems Were Encountered During the Scan List Import at (0)", dt.ToString("u"))
            'INACTIVE'End If

            'INACTIVE'PurgeSmsData()

            'INACTIVE'ImportIngramMicroManifests()
            'INACTIVE'ImportBluePackageManifests()
            'INACTIVE'ImportDSCDeliveryManifests()
            'INACTIVE'ImportPiaDeliveryManifests()
            'INACTIVE'ImportProCourierManifests()

            ImportPartnerBooksManifests()
            ImportMedExpressManifests()
            ImportMedExpressSdManifests()
            ImportAutoZoneManifests()
            ImportDIO2CSManifests()

            Console.Write("Waiting...")

            If FixLocationDatFile() = True Then
                dt = DateTime.Now
                Console.WriteLine("successfully updated at {0}", dt.ToString("u"))
            Else
                Console.WriteLine("")
            End If

            Sleep(30000)  'Change back to 30000

        End While


    End Sub

End Module
