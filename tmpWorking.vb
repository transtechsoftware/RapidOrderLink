Imports System.Data
Imports System.Data.SqlClient

Module tmpWorking

    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Sub tmpMain()

        ' Connection String
        Dim strConnection As String = "Server=12.158.238.68;Database=UN_ORDERS;Uid=RUSR1;Pwd=TOPPRI;"
        'Dim strConnection As String = "Server=67.112.189.227;Database=UN_ORDERS;Uid=RUSR1;Pwd=TOPPRI;"

        ' Connection Objects
        Dim cnSqlServer As SqlConnection ' For main loop
        Dim cnSqlInsert As SqlConnection ' For calling spInsertOrder
        Dim cnSqlUpdate As SqlConnection ' For calling spUpdateOrder
        Dim cnSqlDelete As SqlConnection ' For calling spDeleteOrder

        ' Command Objects
        Dim cmd As New SqlCommand
        Dim cmdInsert As SqlCommand
        Dim cmdUpdate As SqlCommand
        Dim cmdDelete As SqlCommand

        ' Parameter Object
        Dim prmSqlParameter As SqlParameter

        ' Data Readers
        Dim drReader As SqlDataReader = Nothing

        While (1)

            Console.Write("Processing...")

            ' Instantiate Connections
            cnSqlServer = New SqlConnection(strConnection)
            cnSqlInsert = New SqlConnection(strConnection)
            cnSqlUpdate = New SqlConnection(strConnection)
            cnSqlDelete = New SqlConnection(strConnection)

            ' Instantiate Commands
            cmd = New SqlCommand("SELECT * FROM NETSERVER3.RTMGMT.dbo.Orders_Q WHERE Processed = 'N' ORDER BY RowId ASC", cnSqlServer)
            cmdInsert = New SqlCommand("spInsertOrder", cnSqlInsert)
            cmdUpdate = New SqlCommand("spUpdateOrder", cnSqlUpdate)
            cmdDelete = New SqlCommand("spDeleteOrder", cnSqlDelete)

            ' Set Command Types
            cmdInsert.CommandType = CommandType.StoredProcedure
            cmdUpdate.CommandType = CommandType.StoredProcedure
            cmdDelete.CommandType = CommandType.StoredProcedure

            Try
                'open the connection
                If Not cnSqlServer.State = ConnectionState.Open Then cnSqlServer.Open()

                'populate the DataReader
                drReader = cmd.ExecuteReader

                'loop through the results and do something
                Do While drReader.Read

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
                            prmSqlParameter = cmdInsert.Parameters.AddWithValue("@RAPIDID", strRapidID)
                            prmSqlParameter.Direction = ParameterDirection.Input
                            prmSqlParameter.SqlDbType = SqlDbType.Int
                            ' @Q_ROWID
                            prmSqlParameter = cmdInsert.Parameters.AddWithValue("@Q_ROWID", strRowID)
                            prmSqlParameter.Direction = ParameterDirection.Input
                            prmSqlParameter.SqlDbType = SqlDbType.Int
                            ' @WEBORDERID
                            prmSqlParameter = cmdInsert.Parameters.AddWithValue("@WEBORDERID", Nothing)
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
                            MsgBox("Unknown Event Type - " & strEvent)
                    End Select
                Loop

            Catch ex As Exception
                'Error Handling Code Goes Here
                MsgBox(ex.Message)
            Finally
                'clean up code that need to run no matter what

                'close the data reader
                drReader.Close()

                'close the connections
                cnSqlServer.Close()
                cnSqlInsert.Close()
                cnSqlUpdate.Close()
                cnSqlDelete.Close()

                If Not cnSqlServer Is Nothing Then
                    cnSqlServer.Dispose()
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

            Console.WriteLine("Done!")

            Sleep(60000)

        End While

    End Sub

End Module
