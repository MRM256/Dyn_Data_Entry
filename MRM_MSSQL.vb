Imports System.Data.Sql
Imports System.Data.SqlClient
Partial Public Class MRM_MSSQL
    Inherits System.Windows.Forms.Control

#Region "MS-SQL Server Connections"

    Public Function TrustedConnection(ByVal svrName As String,
                                      ByVal svrCatalog As String) _
                                        As String
        'Purpose:       Creates a Windows Trusted Connection string
        'Parameters:    svrName As String - Name of the SQL Server
        '               to which I am connecting.
        '               svrCatalog As String - The SQL Server Catalog
        'Returns:       A properly build Trusted connection string
        Dim cnnSQL As New SqlConnectionStringBuilder
        With cnnSQL
            .DataSource = svrName
            .InitialCatalog = svrCatalog
            .IntegratedSecurity = True
            .TrustServerCertificate = True
            .PersistSecurityInfo = True
        End With
        Return cnnSQL.ConnectionString
    End Function

    Public Function StandardSecurity(ByVal svrName As String,
                                    ByVal svrCatalog As String,
                                    ByVal svrUser As String,
                                    ByVal svrUserPW As String) _
                                    As String
        'Purpose:       Creates a Standard Security Connection string
        'Parameters:    svrName As String - Name of the SQL Server
        '               to which I am connecting.
        '               svrCatalog As String - The SQL Server Catalog.
        '               svrUser As string - The User name for SQL Server.
        '               svrUserPW As String - The User's Password.
        'Returns:       A properly build Trusted connection string
        Dim cnnSQL As New SqlConnectionStringBuilder
        With cnnSQL
            .DataSource = svrName
            .InitialCatalog = svrCatalog
            .UserID = svrUser
            .Password = svrUserPW
            .PersistSecurityInfo = True
        End With
        Return cnnSQL.ConnectionString
    End Function
#End Region

#Region "MS-SQL Server Database Operations"

    Public Function LongestStrInCol(ByRef strCnn As String,
                               ByVal strTable As String,
                               ByVal strColName As String) _
                               As String
        'Purpose:       Find the longest record in the column for '
        '               this table.
        'Parameters:    strCnn - A properly built connection string
        '               strTable - Name of the table whose records
        '                          We are interested in.
        '               strColName - Column we are searching
        'Returns:       The Longest Record as a string
        Dim strSQL As String
        Using cnn As New SqlConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    'Get Column Schema Information from 
                    'selected Server's Database
                    strSQL = "SELECT TOP 1 WITH TIES " & strColName &
                         ", DATALENGTH(" & strColName & " ) " &
                         " AS TextLength FROM [" & strTable & "]" &
                         "ORDER BY TextLength DESC;"
                    Dim da As New SqlDataAdapter(strSQL, cnn)
                    Dim ds As New DataSet
                    da.Fill(ds)
                    Dim dt As DataTable = ds.Tables(0)
                    'lngMaxStrLen = dt.Rows(0).Item(0)
                    'Debug.Print("Longest String in table: " & strTable & _
                    '            " from Column: " & strColName & _
                    '            " is - " & vbCrLf & dt.Rows(0).Item(0).ToString)
                    If dt.Rows.Count > 0 Then
                        Return dt.Rows(0).Item(0).ToString
                    Else
                        Return Nothing
                    End If
                End If
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                Return Nothing
            End Try
            Return Nothing
            cnn.Dispose()
        End Using
    End Function
    Public Function MSSQL_DBs(ByVal strCnn As String) _
                             As List(Of String)
        'Purpose:       Fill cboSQLDB ComboBox list with available 
        '               database names
        'Parameters:    strCnn As String - The SQL Server Connection 
        '               string
        'Returns:       A list of databases contained on this 
        '               SQL Server Instance
        Dim listDataBases As List(Of String) = New List(Of String),
            strSQL As String
        Using cnn As New SqlConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    'Get database names from selected server
                    strSQL = "SELECT Name FROM sys.databases;"
                    Dim cmd As SqlCommand = New SqlCommand(strSQL, cnn),
                        dr As SqlDataReader = cmd.ExecuteReader()
                    While (dr.Read())
                        listDataBases.Add(dr(0).ToString())
                    End While
                    'Set databases list as combobox’s datasource
                    dr.Close()
                    cmd.Dispose()
                End If
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
            End Try
            'cnn.Dispose()
        End Using
        Return listDataBases
    End Function
    Public Function MSSQL_DB_Tbls(ByVal strCnn As String) _
                                    As List(Of String)
        'Purpose:       Creates a list of Database Tables,
        '               list of tables from the selected database.
        'Parameters:    strCnn as string
        'Returns:       Creates a list of database tables
        Dim listTables As List(Of String) = New List(Of String),
            strSQL As String
        Using cnn As New SqlConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    'Get database table names from selected server's
                    'database.
                    strSQL = "SELECT TABLE_NAME FROM " &
                             "INFORMATION_SCHEMA.TABLES"
                    Dim cmd As SqlCommand = New SqlCommand(strSQL, cnn)
                    Dim dr As SqlDataReader = cmd.ExecuteReader()
                    While (dr.Read())
                        listTables.Add(dr(0).ToString())
                    End While
                    'Set table list as combobox’s datasource
                    dr.Close()
                    cmd.Dispose()
                End If
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
            End Try
            'cnn.Dispose()
        End Using
        Return listTables
    End Function
    Public Function Lst_SQLServers() As List(Of String)
        'Purpose:       List all instances of MS-SQL Servies on the
        '               the Network
        'Parameters:    None
        'Returns:       A List of SQL Server instances.
        'declare variables
        Dim dt As Data.DataTable = Nothing,
            dr As Data.DataRow = Nothing,
            listSQL = New List(Of String)

        Try
            'get sql server instances in to DataTable object
            listSQL = SqlDataSourceEnumerator.
                Instance.GetDataSources().
                    AsEnumerable().
                        Select(Function(row) row.Field(Of String)("ServerName")).
                            ToList()

        Catch ex As System.Data.SqlClient.SqlException
            'Log error 
            Dim el As New Log.ErrorLogger
            el.WriteToErrorLog(ex.Message,
                           ex.StackTrace,
                           "Error - lst_SQLServers")

        Catch ex As Exception
            'Log error 
            Dim el As New Log.ErrorLogger
            el.WriteToErrorLog(ex.Message,
                           ex.StackTrace,
                           "Error - lst_SQLServers")
        Finally
            'clean up ;)
            dr = Nothing
            'dt = Nothing
        End Try
        Return listSQL
    End Function
    Public Function Dt_MSSQL_Tbl_Relationships(ByVal strCnn As String,
                                               ByVal strTable As String) _
                                       As DataTable
        'Purpose:       Creates a DataTable of table relationships.
        'Parameters:    strCnn - Connection string to SQL Server;
        '               strTable - Name of the table we want to see
        '               if any other table(s) have some relationship
        '               to it.
        'Returns:       DataTable of related tables
        Dim strSQL As String
        Using cnn As New SqlConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    'SQL to return all child tables, the Foreign key
                    'Column, Primary Key column and the constraint name
                    'based on the given parent table.
                    strSQL = "SELECT Child_Table = FK.TABLE_NAME, " &
                         "FK_Column = CU.COLUMN_NAME, " &
                         "PK_Column = PT.COLUMN_NAME, " &
                         "Constraint_Name = C.CONSTRAINT_NAME " &
                         "FROM " &
                         "INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS C " &
                         "INNER JOIN " &
                         "INFORMATION_SCHEMA.TABLE_CONSTRAINTS " &
                         "FK ON C.CONSTRAINT_NAME = " &
                         "FK.CONSTRAINT_NAME " &
                         "INNER JOIN " &
                         "INFORMATION_SCHEMA.TABLE_CONSTRAINTS " &
                         "PK ON C.UNIQUE_CONSTRAINT_NAME = " &
                         "PK.CONSTRAINT_NAME " &
                         "INNER JOIN " &
                         "INFORMATION_SCHEMA.KEY_COLUMN_USAGE " &
                         "CU ON C.CONSTRAINT_NAME = " &
                         "CU.CONSTRAINT_NAME " &
                         "INNER JOIN " &
                         "(SELECT i1.TABLE_NAME, i2.COLUMN_NAME FROM " &
                         "INFORMATION_SCHEMA.TABLE_CONSTRAINTS i1 " &
                         "INNER JOIN " &
                         "INFORMATION_SCHEMA.KEY_COLUMN_USAGE " &
                         "i2 ON i1.CONSTRAINT_NAME = i2.CONSTRAINT_NAME " &
                         "WHERE i1.CONSTRAINT_TYPE = 'PRIMARY KEY') " &
                         "PT ON PT.TABLE_NAME = PK.TABLE_NAME " &
                         "WHERE PK.TABLE_NAME = '" & strTable & "';"
                    Dim da As New SqlDataAdapter(strSQL, cnn),
                        ds As New DataSet
                    da.Fill(ds)
                    Dim dt As DataTable = ds.Tables(0)
                    'Debug Purposes - Display Data
                    'Dim I As Long
                    'Debug.Print(dt.Columns.Item(0).ToString & " " & _
                    '            dt.Columns.Item(1).ToString & " " & _
                    '            dt.Columns.Item(2).ToString & " " & _
                    '            dt.Columns.Item(3).ToString)
                    'For I = 0 To dt.Rows.Count - 1
                    'Debug.Print(dt.Rows(I).Item(0).ToString & " " & _
                    '            dt.Rows(I).Item(1).ToString & " " & _
                    '            dt.Rows(I).Item(2).ToString & " " & _
                    '            dt.Rows(I).Item(3).ToString)
                    'Next
                    Return dt
                End If
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                Return Nothing
            End Try
        End Using
        Return Nothing
    End Function
    Public Function Dt_MSSQL_Tbl_Schema(ByVal strCnn As String,
                                    ByVal strTable As String) _
                                    As DataTable
        'Purpose:       Creates a list of table columns(fields).
        '               We use the generated list to make the 
        '               TabPages to insert, or edit data for the table.
        'Parameters:    strCnn - Connection string to SQL Server;
        '               strTable - Name of the table we are using
        'Returns:       DataTable of Column Schema Information
        Dim strSQL As String
        Using cnn As New SqlConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    'Get Column Schema Information from 
                    'selected Server's Database
                    strSQL = "SELECT Ordinal_Position, Column_Name, " &
                         "Data_Type, Character_Maximum_Length, " &
                         "Is_Nullable, Table_Name " &
                         "FROM Information_Schema.Columns " &
                         "WHERE Table_Name = '" & strTable & "'" &
                         "ORDER BY Ordinal_Position ASC;"
                    Dim da As New SqlDataAdapter(strSQL, cnn),
                        ds As New DataSet
                    da.Fill(ds)
                    Dim dt As DataTable = ds.Tables(0)
                    'Call dt_Viewer(dt)
                    'Debug Purposes - Display Data
                    'Dim I As Long
                    'For I = 0 To dt.Rows.Count - 1
                    'Debug.Print(dt.Rows(I).Item(1).ToString)
                    'Next
                    Return dt
                End If
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                Return Nothing
                'Finally
                '    cnn.Dispose()
            End Try
        End Using
        Return Nothing
    End Function
    Public Function Dt_MSSQL_Retrive_Data(ByVal strCnn As String,
                                          ByVal strTable As String) _
                                          As DataTable
        'Purpose:       Creates a DataTable containing Rows(Records)
        '               for the desired table name
        'Parameters:    strCnn - A properly built connection string
        '               strTable - Name of the table whose records
        '                          We are interested in.
        'Returns:       A DataTable of Records.
        Dim strSQL As String,
            dt As New DataTable
        Using cnn As New SqlConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    'Get all Information from 
                    'selected Server's Database table
                    strSQL = "SELECT * FROM [" & strTable & "];"
                    Using dad As New SqlDataAdapter(strSQL, cnn)
                        dad.Fill(dt)
                        'Call Dt_Viewer(dt)
                    End Using

                    cnn.Close()
                    'Debug Purposes - Display Data
                    'Dim I As Long
                    'For I = 0 To dt.Rows.Count - 1
                    'Debug.Print(dt.Rows(I).Item(1).ToString)
                    'Next
                    dt.TableName = strTable
                    Return dt
                End If
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                Return Nothing
            End Try
            Return Nothing
            'cnn.Dispose()
        End Using
    End Function
    Public Function Det_Tbl_FK(ByVal strCnn As String,
                           ByVal strTbl As String,
                           ByVal strCol As String) As Boolean
        'Purpose:       Checks the Column name with the table name 
        '               to determine if the selected column name
        '               contains the Foreign Key for this table.
        'Parameters:    strCnn - Connection string to MSSQL
        '               strTbl - Name of the table
        '               strCol - Column name we are checking
        'Returns:       True, if strCol is the Foreign Key column;
        '               False otherwise
        Dim strSQL As String,
            b_FK As Boolean = False
        Using cnn As New SqlConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    'Get all Information from 
                    'selected Server's Database table                   
                    strSQL = "SELECT COLUMN_NAME FROM " &
                         "INFORMATION_SCHEMA.KEY_COLUMN_USAGE " &
                         "WHERE OBJECTPROPERTY(OBJECT_ID(CONSTRAINT_SCHEMA+" &
                         "'.'+CONSTRAINT_NAME), 'IsForeignKey') = 1 " &
                         "AND TABLE_NAME = '" & strTbl &
                         "' AND COLUMN_NAME = '" & strCol & "';"

                    Dim dt As New DataTable
                    Using dad As New SqlDataAdapter(strSQL, cnn)
                        dad.Fill(dt)
                    End Using
                    cnn.Close()
                    If dt.Rows.Count > 0 Then
                        b_FK = True
                    Else
                        b_FK = False
                    End If
                End If
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                Return Nothing
            End Try
        End Using
        Return b_FK
    End Function
    Public Function Det_Tbl_PK(ByVal strCnn As String,
                           ByVal strTbl As String,
                           ByVal strCol As String) As Boolean
        'Purpose:       Checks the Column name with the table name 
        '               to determine if the selected column name
        '               contains the Primary Key for this table.
        'Parameters:    strCnn - Connection string to MSSQL
        '               strTbl - Name of the table
        '               strCol - Column name we are checking
        'Returns:       True, if strCol is the Primary Key column;
        '               False otherwise
        Dim strSQL As String,
            b_PK As Boolean = False
        Using cnn As New SqlConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    'Get all Information from 
                    'selected Server's Database table                   
                    strSQL = "SELECT COLUMN_NAME FROM " &
                         "INFORMATION_SCHEMA.KEY_COLUMN_USAGE " &
                         "WHERE OBJECTPROPERTY(OBJECT_ID(CONSTRAINT_SCHEMA+" &
                         "'.'+CONSTRAINT_NAME), 'IsPrimaryKey') = 1 " &
                         "AND TABLE_NAME = '" & strTbl &
                         "' AND COLUMN_NAME = '" & strCol & "';"

                    Dim dt As New DataTable
                    Using dad As New SqlDataAdapter(strSQL, cnn)
                        dad.Fill(dt)
                    End Using
                    cnn.Close()
                    If dt.Rows.Count > 0 Then
                        b_PK = True
                    Else
                        b_PK = False
                    End If
                End If
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                Return Nothing
            End Try
        End Using
        Return b_PK
    End Function

#End Region

#Region "DataTable Viewer"
    Public Sub Dt_Viewer(ByVal dt As DataTable)
        'Purpose:       Lets the Developer view DataTable contents
        'Parameters:    dt As DataTable
        'Returns:       A Log file
        Dim rowData As String = ""
        Dim el As New Log.ErrorLogger
        'dt in the name of the data table
        For Each row As DataRow In dt.Rows
            For Each column As DataColumn In dt.Columns
                rowData = rowData & column.ColumnName & "=" & row(column) & " "
            Next
            rowData = rowData & vbCrLf & vbCrLf
        Next
        el.WriteToErrorLog(rowData, " ", "DataTable Viewer")
        'MessageBox.Show(rowData)
    End Sub

#End Region
End Class


