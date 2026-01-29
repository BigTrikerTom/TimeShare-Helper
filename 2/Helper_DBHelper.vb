Imports MySql.Data.MySqlClient
Imports TimeShare_Helper
Imports TimeShare_Error

Friend Class HelperDB
    Friend Structure CountryDataDef
        Friend CountryEn As String
        Friend CountryDe As String
        Friend iso2 As String
        Friend iso3 As String
        Friend lieferland As String
        Friend isonum As String
        Friend language As String
        Sub New(ByVal Optional _CountryEn As String = "",
                ByVal Optional _CountryDe As String = "",
                ByVal Optional _iso2 As String = "",
                ByVal Optional _iso3 As String = "",
                ByVal Optional _lieferland As String = "",
                ByVal Optional _isonum As String = "",
                ByVal Optional _language As String = "")
            CountryEn = _CountryEn
            CountryDe = _CountryDe
            iso2 = _iso2
            iso3 = _iso3
            lieferland = _lieferland
            isonum = _isonum
            language = _language
        End Sub
    End Structure
    Friend Enum ReadWriteSettingFor
        all = 0
        Server = 1
        MailServer = 2
        Client = 3
        Other = 4
        Auswerter = 5
    End Enum
    Friend Shared Function GetCountryByISO(ByVal CountryIso As String, ByVal Database As clsDbConnectLocal.SelectDatabase, ByVal DbTable As String) As String
        Dim Country As String = ""
        Try
            Using mhgcbi1c As New clsDBconnect
                If mhgcbi1c.connect(Database) Then
                    Dim query As String = "SELECT * FROM `" & DbTable & "` WHERE "
                    query = query & "`iso2` LIKE ?iso2? "
                    query = query & ";"
                    mhgcbi1c.cmd.CommandText = query
                    mhgcbi1c.cmd.Parameters.Clear()
                    mhgcbi1c.cmd.Parameters.AddWithValue("?iso2?", CountryIso)
                    'Debug.Print(ParameterQuery(mhgcbi1c))
                    Using reader_mhgcbi1c As MySqlDataReader = mhgcbi1c.cmd.ExecuteReader
                        While reader_mhgcbi1c.Read()
                            Country = Helper_Convert.ConvertToString(reader_mhgcbi1c.Item("land"))
                            Exit While
                        End While
                    End Using
                End If
            End Using

        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return Country
    End Function
    Friend Shared Function GetISOByCountry(ByVal Country As String, ByVal Database As clsDbConnectLocal.SelectDatabase, ByVal DbTable As String) As String
        Dim ISO2 As String = ""
        Try
            Using mhgcbi1c As New clsDBconnect
                If mhgcbi1c.connect(Database) Then
                    Dim query As String = "SELECT * FROM `" & DbTable & "` WHERE "
                    query = query & "`country` LIKE ?land? OR "
                    query = query & "`land` LIKE ?land? "
                    query = query & ";"
                    mhgcbi1c.cmd.CommandText = query
                    mhgcbi1c.cmd.Parameters.Clear()
                    mhgcbi1c.cmd.Parameters.AddWithValue("?land?", Country)
                    'Debug.Print(ParameterQuery(mhgcbi1c))
                    Using reader_mhgcbi1c As MySqlDataReader = mhgcbi1c.cmd.ExecuteReader
                        While reader_mhgcbi1c.Read()
                            ISO2 = Helper_Convert.ConvertToString(reader_mhgcbi1c.Item("iso2"))
                            Exit While
                        End While
                    End Using
                End If
            End Using

        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return ISO2
    End Function
    Friend Shared Function GetISO3ByCountry(ByVal Country As String, ByVal Database As clsDbConnectLocal.SelectDatabase, ByVal DbTable As String) As String
        Dim ISO3 As String = ""
        Try
            Using mhgcbi1c As New clsDBconnect
                If mhgcbi1c.connect(Database) Then
                    Dim query As String = "SELECT * FROM `" & DbTable & "` WHERE "
                    query = query & "`country` LIKE ?land? OR "
                    query = query & "`land` LIKE ?land? "
                    query = query & ";"
                    mhgcbi1c.cmd.CommandText = query
                    mhgcbi1c.cmd.Parameters.Clear()
                    mhgcbi1c.cmd.Parameters.AddWithValue("?land?", Country)
                    'Debug.Print(ParameterQuery(mhgcbi1c))
                    Using reader_mhgcbi1c As MySqlDataReader = mhgcbi1c.cmd.ExecuteReader
                        While reader_mhgcbi1c.Read()
                            ISO3 = Helper_Convert.ConvertToString(reader_mhgcbi1c.Item("iso3"))
                        End While
                    End Using
                End If
            End Using

        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return ISO3
    End Function
    Friend Shared Function ParameterQuery(ByVal TempDBConnection As clsDBconnect) As String
        If Helper.IsIDE() Then
            Dim CommandText As String = ""
            Try
                If Not IsNothing(TempDBConnection.cmd) AndAlso TempDBConnection.DBIsOpen Then
                    CommandText = TempDBConnection.cmd.CommandText
                    For Each p As MySqlParameter In TempDBConnection.cmd.Parameters
                        Dim value As String = ""
                        If IsNothing(p.Value) OrElse IsDBNull(p.Value) Then
                            value = "null"
                        Else
                            value = p.Value.ToString
                        End If
                        If p.MySqlDbType.ToString.Contains("VarChar") Then
                            CommandText = CommandText.Replace(p.ParameterName, "'" & CStr(value) & "'")
                        ElseIf p.MySqlDbType.ToString.Contains("Blob") Then
                            CommandText = CommandText.Replace(p.ParameterName, "'BLOB'")
                        Else
                            CommandText = CommandText.Replace(CStr(p.ParameterName), CStr(value))
                        End If
                    Next
                End If
                Return CommandText
            Catch ex As Exception
                Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
                If Helper.IsIDE Then Stop
                Return ex.Message
            End Try
        Else
            Return ""
        End If

    End Function
    Friend Shared Function GetCountry(ByVal Country As String) As CountryDataDef
        Dim query_cpcwos2d2c As String = ""
        Dim CountryData As New CountryDataDef
        '!? Hole Länderdaten aus der DB 
        Using cpcwos2d2c As New clsDBconnect
            If cpcwos2d2c.connect(clsDbConnectLocal.SelectDatabase.ConDrop) Then
                query_cpcwos2d2c = "SELECT * FROM `" & clsDbConnectLocal.db_table_condrop_country & "` WHERE "
                query_cpcwos2d2c = query_cpcwos2d2c & "`iso3` LIKE ?Country? OR "
                query_cpcwos2d2c = query_cpcwos2d2c & "`iso2` LIKE ?Country? OR "
                query_cpcwos2d2c = query_cpcwos2d2c & "`country` LIKE ?Country? OR "
                query_cpcwos2d2c = query_cpcwos2d2c & "`land` LIKE ?Country? "
                query_cpcwos2d2c = query_cpcwos2d2c & "; "
                cpcwos2d2c.cmd.CommandText = query_cpcwos2d2c
                cpcwos2d2c.cmd.Parameters.Clear()
                cpcwos2d2c.cmd.Parameters.AddWithValue("?Country?", Helper_Convert.ConvertToString(Country, False, ""))
               'Debug.Print(HelperDB.ParameterQuery(cpcwos2d2c))
                Using reader_cpcwos2d2c As MySqlDataReader = cpcwos2d2c.cmd.ExecuteReader
                    While reader_cpcwos2d2c.Read()
                        CountryData.CountryEn = Helper_Convert.ConvertToString(reader_cpcwos2d2c("country"))
                        CountryData.CountryDe = Helper_Convert.ConvertToString(reader_cpcwos2d2c("land"))
                        CountryData.iso2 = Helper_Convert.ConvertToString(reader_cpcwos2d2c("iso2"))
                        CountryData.iso3 = Helper_Convert.ConvertToString(reader_cpcwos2d2c("iso3"))
                        CountryData.lieferland = Helper_Convert.ConvertToString(reader_cpcwos2d2c("lieferland"))
                        CountryData.isonum = Helper_Convert.ConvertToString(reader_cpcwos2d2c("isonum"))
                        CountryData.language = Helper_Convert.ConvertToString(reader_cpcwos2d2c("lang"))
                    End While
                End Using
            End If
        End Using
        Return CountryData
    End Function

#Region "Settings (DB)"
    Friend Shared Function ReadSettingsFromDB(ByVal SourceProgram As ReadWriteSettingFor,
                                              ByVal Key As String,
                                              ByVal Key_1 As String,
                                              ByVal Database As clsDbConnectLocal.SelectDatabase,
                                              ByVal DbTable As String,
                                              Optional ByVal ProgramOther As String = ""
                                              ) As Helper.ReadSettingReturn
        Dim Result As Helper.ReadSettingReturn = Nothing
        Dim Program As String = ""
        Try
            Select Case SourceProgram
                Case ReadWriteSettingFor.all
                    Program = "all"
                Case ReadWriteSettingFor.Server
                    Program = "Server"
                Case ReadWriteSettingFor.MailServer
                    Program = "MailServer"
                Case ReadWriteSettingFor.Client
                    Program = "Client"
                Case ReadWriteSettingFor.Other
                    Program = ProgramOther
            End Select
            If Not String.IsNullOrWhiteSpace(Helper_Convert.ConvertToString(DbTable)) Then
                Dim query As String = ""
                Using db_rsfdb As New clsDBconnect
                    If db_rsfdb.connect(Database) Then
                        query = "SELECT * FROM `" & DbTable & "` WHERE"
                        query = query & " `Programm` = ?Programm? AND "
                        query = query & " `Key` = ?Key? AND "
                        query = query & " `Key_1` = ?Key_1? "
                        query = query & ";"
                        db_rsfdb.cmd.CommandText = query
                        db_rsfdb.cmd.Parameters.Clear()
                        db_rsfdb.cmd.Parameters.AddWithValue("?Programm?", Program)
                        db_rsfdb.cmd.Parameters.AddWithValue("?Key?", Key)
                        db_rsfdb.cmd.Parameters.AddWithValue("?Key_1?", Key_1)
                        'Debug.Print(ParameterQuery(db_rsfdb))
                        Using reader_1t As MySqlDataReader = db_rsfdb.cmd.ExecuteReader
                            While reader_1t.Read()
                                Application.DoEvents()
                                Result.Value = Helper_Convert.ConvertToString(reader_1t.Item("Value"))
                                Result.Value_1 = Helper_Convert.ConvertToString(reader_1t.Item("Value_1"))
                                Result.Value_2 = Helper_Convert.ConvertToString(reader_1t.Item("Value_2"))
                            End While
                        End Using
                    End If
                End Using
            End If

        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return Result
    End Function
    Friend Shared Function WriteSettings2DB(ByVal Key As String,
                                           ByVal Key_1 As String,
                                           ByVal Value As String,
                                           ByVal Value_1 As String,
                                           ByVal Database As clsDbConnectLocal.SelectDatabase,
                                           ByVal DbTable As String,
                                           Optional ByVal Programm As String = "all"
                                           ) As Integer
        Dim query As String = ""
        Dim Value_Old As String = "-"
        Dim Value_Old1 As String = "-"
        Dim Value_Old2 As String = "-"
        dim ri As Integer=0
        Try
            If Not String.IsNullOrWhiteSpace(Helper_Convert.ConvertToString(DbTable)) Then
                Using db_ws2db1 As New clsDBconnect
                    If db_ws2db1.connect(Database) Then
                        query = "SELECT * FROM `" & DbTable & "` WHERE"
                        query = query & " `Programm` LIKE ?Programm? AND "
                        query = query & " `Key` LIKE ?Key? AND "
                        query = query & " `Key_1` LIKE ?Key_1? "
                        query = query & ";"
                        db_ws2db1.cmd.CommandText = query
                        db_ws2db1.cmd.Parameters.Clear()
                        db_ws2db1.cmd.Parameters.AddWithValue("?Programm?", Programm)
                        db_ws2db1.cmd.Parameters.AddWithValue("?Key?", Key)
                        db_ws2db1.cmd.Parameters.AddWithValue("?Key_1?", Key_1)
                        'Debug.Print(ParameterQuery(db_ws2db1))
                        Using reader_db_ws2db1 As MySqlDataReader = db_ws2db1.cmd.ExecuteReader
                            While reader_db_ws2db1.Read()
                                Application.DoEvents()
                                Value_Old1 = Helper_Convert.ConvertToString(reader_db_ws2db1.Item("Value"))
                                Value_Old2 = Helper_Convert.ConvertToString(reader_db_ws2db1.Item("Value_1"))
                                Value_Old = Value_Old1 & Value_Old2
                            End While
                        End Using
                        If String.IsNullOrEmpty(Value_Old1) AndAlso String.IsNullOrEmpty(Value_Old2) Then  ' insert
                            query = "INSERT INTO `" & DbTable & "` "
                            query = query & "("
                            query = query & "`ID`, "
                            query = query & "`TimeStamp`, "
                            query = query & "`Programm`, "
                            query = query & "`Key`, "
                            query = query & "`Key_1`, "
                            query = query & "`Value`, "
                            query = query & "`Value_1`"
                            query = query & ") VALUES ("
                            query = query & "null, "
                            query = query & "?value_1?, "
                            query = query & "?value_2?, "
                            query = query & "?value_3?, "
                            query = query & "?value_4?, "
                            query = query & "?value_5?, "
                            query = query & "?value_6? "
                            query = query & ")"
                            query = query & ";"
                        Else                    ' update
                            query = "UPDATE `" & DbTable & "` SET "
                            query = query & "`TimeStamp` = ?value_1?, "
                            query = query & "`Programm` = ?value_2?, "
                            query = query & "`Key` = ?value_3?, "
                            query = query & "`Key_1` = ?value_4?, "
                            query = query & "`Value` = ?value_5?, "
                            query = query & "`Value_1` = ?value_6? "
                            query = query & " WHERE "
                            query = query & "`Programm` LIKE ?value_2? AND "
                            query = query & "`Key` LIKE ?value_3? AND "
                            query = query & "`Key_1` LIKE ?value_4? "
                            query = query & ";"
                        End If
                        db_ws2db1.cmd.CommandText = query
                        db_ws2db1.cmd.Parameters.Clear()
                        db_ws2db1.cmd.Parameters.AddWithValue("?value_1?", Format(Now, "yyyy-MM-dd HH:mm:ss"))
                        db_ws2db1.cmd.Parameters.AddWithValue("?value_2?", Programm)
                        db_ws2db1.cmd.Parameters.AddWithValue("?value_3?", Key)
                        db_ws2db1.cmd.Parameters.AddWithValue("?value_4?", Key_1)
                        db_ws2db1.cmd.Parameters.AddWithValue("?value_5?", Value)
                        db_ws2db1.cmd.Parameters.AddWithValue("?value_6?", Value_1)
                        'Debug.Print(ParameterQuery(db_ws2db1))
                        ri = db_ws2db1.cmd.ExecuteNonQuery()
                    End If
                End Using
            End If
            Return ri
        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False)
            If Helper.IsIDE() Then Stop
            Return 0
        End Try
    End Function
    Friend Shared Function WriteSetting2DB(ByVal SourceProgram As ReadWriteSettingFor,
                                           ByVal Key As String,
                                           ByVal Key_1 As String,
                                           ByVal Value As Object,
                                           ByVal Value_1 As Object,
                                           ByVal Database As clsDbConnectLocal.SelectDatabase,
                                           ByVal DbTable As String,
                                           Optional ByVal ProgramOther As String = ""
                                           ) As Boolean
        Dim ReturnVal As Boolean = False
        Dim query_fmdipwfc1c As String = ""
        Dim ResCount As Integer = 0
        Dim Program As String = ""
        dim ri As Integer=0

        Try
            Select Case SourceProgram
                Case ReadWriteSettingFor.all
                    Program = "all"
                Case ReadWriteSettingFor.Server
                    Program = "Server"
                Case ReadWriteSettingFor.MailServer
                    Program = "MailServer"
                Case ReadWriteSettingFor.Client
                    Program = "Client"
                Case ReadWriteSettingFor.Other
                    Program = ProgramOther
            End Select
            If Not String.IsNullOrWhiteSpace(Helper_Convert.ConvertToString(DbTable)) Then
                Try
                    Using db_ws2db1 As New clsDBconnect
                        If db_ws2db1.connect(Database) Then
                            query_fmdipwfc1c = "SELECT `id` FROM `" & DbTable & "` WHERE "
                            query_fmdipwfc1c = query_fmdipwfc1c & "`Programm` LIKE ?Program? AND "
                            query_fmdipwfc1c = query_fmdipwfc1c & "`Key` LIKE ?Key? "
                            If Not String.IsNullOrWhiteSpace(Key_1) Then
                                query_fmdipwfc1c = query_fmdipwfc1c & "AND `Key_1` LIKE ?Key_1? "
                            End If
                            db_ws2db1.cmd.CommandText = query_fmdipwfc1c
                            db_ws2db1.cmd.Parameters.Clear()
                            db_ws2db1.cmd.Parameters.AddWithValue("?Program?", Program)
                            db_ws2db1.cmd.Parameters.AddWithValue("?Key?", Key)
                            db_ws2db1.cmd.Parameters.AddWithValue("?Key_1?", Key_1)
                            'Debug.Print(ParameterQuery(db_ws2db1))
                            Dim total As Integer = Helper_Convert.ConvertToInteger(db_ws2db1.cmd.ExecuteScalar)
                            If total > 0 Then
                                query_fmdipwfc1c = "DELETE FROM `" & DbTable & "` WHERE "
                                query_fmdipwfc1c = query_fmdipwfc1c & "`Programm` LIKE ?Program? AND "
                                query_fmdipwfc1c = query_fmdipwfc1c & "`Key` LIKE ?Key? "
                                If Not String.IsNullOrWhiteSpace(Key_1) Then
                                    query_fmdipwfc1c = query_fmdipwfc1c & "AND `Key_1` LIKE ?Key_1? "
                                End If
                                db_ws2db1.cmd.CommandText = query_fmdipwfc1c
                                db_ws2db1.cmd.Parameters.Clear()
                                db_ws2db1.cmd.Parameters.AddWithValue("?Program?", Program)
                                db_ws2db1.cmd.Parameters.AddWithValue("?Key?", Key)
                                db_ws2db1.cmd.Parameters.AddWithValue("?Key_1?", Key_1)
                               'Debug.Print(ParameterQuery(db_ws2db1))
                                ri = db_ws2db1.cmd.ExecuteNonQuery
                            End If

                            query_fmdipwfc1c = "INSERT INTO `" & DbTable & "` ("
                            query_fmdipwfc1c = query_fmdipwfc1c & "`ID`, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "`TimeStamp`, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "`Programm`, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "`Key`, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "`Key_1`, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "`Value`, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "`Value_1`, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "`uniqueKey` "
                            query_fmdipwfc1c = query_fmdipwfc1c & ") VALUES ("
                            query_fmdipwfc1c = query_fmdipwfc1c & "null, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "?TimeStamp?, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "?Program?, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "?Key?, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "?Key_1?, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "?Value?, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "?Value_1?, "
                            query_fmdipwfc1c = query_fmdipwfc1c & "?uniqueKey? "
                            query_fmdipwfc1c = query_fmdipwfc1c & ")"
                            query_fmdipwfc1c = query_fmdipwfc1c & ";"

                            db_ws2db1.cmd.CommandText = query_fmdipwfc1c
                            db_ws2db1.cmd.Parameters.Clear()
                            db_ws2db1.cmd.Parameters.AddWithValue("?TimeStamp?", Format(Now, "yyyy-MM-dd HH:mm:ss"))
                            db_ws2db1.cmd.Parameters.AddWithValue("?Program?", Program)
                            db_ws2db1.cmd.Parameters.AddWithValue("?Key?", Key)
                            db_ws2db1.cmd.Parameters.AddWithValue("?Key_1?", Key_1)
                            db_ws2db1.cmd.Parameters.AddWithValue("?Value?", Value)
                            db_ws2db1.cmd.Parameters.AddWithValue("?Value_1?", Value_1)
                            db_ws2db1.cmd.Parameters.AddWithValue("?uniqueKey?", Program & Key & Key_1)
                            'Debug.Print(ParameterQuery(db_ws2db1))
                            ResCount = db_ws2db1.cmd.ExecuteNonQuery()

                        End If
                    End Using
                Catch ex As MySqlException
                    Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False, False)
                    If Helper.IsIDE() Then Stop
                End Try
            End If
            If ResCount > 0 Then
                ReturnVal = True
            Else
                ReturnVal = False
            End If
        Catch ex As Exception
            Call ErrorHandling.HandleErrorCatch(ex, Error_Helper.GetCallingProc(), System.Reflection.MethodBase.GetCurrentMethod().Name, Environment.CurrentManagedThreadId, False)
            If Helper.IsIDE() Then Stop
        End Try
        Return ReturnVal

    End Function
#End Region
End Class
