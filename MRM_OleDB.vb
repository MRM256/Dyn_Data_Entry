Imports System.Data.OleDb
Imports System.IO
Imports System.Windows.Forms
Public Class MRM_OleDB
    'Purpose:       Where we keep all OleDB operations 
    'Property:      GetOleDBName - Get/Set the OleDB 
    '                              file name.
    '               getCnn   - Returns the prebuilt connection 
    '                          string.
    'Methods:       cnnOLEDB - Determines the .Provider
    '                          part of the connection string
    '               cnnOLEDB - Allows the user to select the
    '                          OleDB(Access) database.

    Private _strOleDBName As String,
            _cnn As String
    Public dt_Rel As New DataTable
    Public Sub Dt_Remove_Rel()
        'Remove DataTRable Columns from dt_Rel
        Dim columncount As Integer = dt_Rel.Columns.Count - 1
        For i = columncount To 0 Step -1
            dt_Rel.Columns.RemoveAt(i)
        Next
    End Sub

#Region "Debug Tool - DataTable Viewer Subroutine"
    Public Sub Dt_Read(ByVal dt As DataTable)
        'Purpose:       Lets the programmer see the data
        '               contained in the DataTable
        'Parameters:    dt As DataTable
        'Returns:       Nothing - Just a Debug routine

        Dim myRow As DataRow,
            myCol As DataColumn
        'For each field in the table...
        For Each myRow In dt.Rows
            'For each property of the field...
            For Each myCol In dt.Columns
                'Display the field name and value.
                Debug.Print(myCol.ColumnName & vbTab &
                            myRow(myCol).ToString())
            Next
            Debug.Print(vbCrLf)
        Next
    End Sub
#End Region
    Public Property GetOleDBName As String
        Get
            Return _strOleDBName
        End Get
        Set(value As String)
            _strOleDBName = value
            _cnn = CnnOLEDB(_strOleDBName)
        End Set
    End Property
    Public ReadOnly Property GetAccessCnn As String
        Get
            Return _cnn
        End Get
    End Property
    Public Function FindFK(ByVal dt As DataTable,
                           ByVal strFK As String) As Boolean
        'Purpose:       Search the dt DataTable where Column 
        '               FK_Tbl_Name equals variable strTbl and
        '               FK_Col_Name equals variable strFK.
        '               If a both values exists then  Return to TRUE.
        'Parameters:    dt As DataTable; The DataTable we are searching
        '               strTbl As String; The table containing the
        '               foreign key;
        '               strFK As String; The foreign key we are searching for
        'Returns:       True if conditions exist; False otherwise
        Dim expression As String,
            foundRows() As DataRow,
            i As Integer,
            FK_Found As Boolean

        FK_Found = False

        'Call dt_Read(dt)
        'SELECT expression 
        'expression = "FK_Tbl_Name = '" & _
        '             strTbl & "' AND FK_Col_Name = '" & _
        '             strFK & "'"
        expression = "FK_Col_Name = '" &
                     strFK & "'"
        ' Use the Select method to find all rows matching the filter.
        foundRows = dt.Select(expression)

        For i = 0 To foundRows.GetUpperBound(0)
            If (foundRows(i)(3).ToString = strFK) Then
                'Debug.Print(foundRows(i)(2).ToString & " " & _
                '            foundRows(i)(3).ToString)
                FK_Found = True
            Else
                FK_Found = False
            End If
        Next i
        Return FK_Found
    End Function
    Private Function CnnOLEDB(ByVal _strOleDBName As String) As String
        'Purpose:       To create a OleDB Connection string.
        'Parameters:    strDb As String - directory path to the database
        '               to which we are connecting
        'Returns:       A properly built OLEDB connection string
        Dim cnBuilder As New OleDbConnectionStringBuilder,
            strSplit() As String,
            strTest As String
        Try
            If Not String.IsNullOrEmpty(_strOleDBName) Then
                'Determine which file extension is in use
                strTest = Path.GetExtension(_strOleDBName)
                strSplit = Split(strTest, ".")
                With cnBuilder
                    Select Case UCase(strSplit(1))
                        Case "ACCDB"
                            'MS Access 2007 - 2010
                            .Provider = "Microsoft.ACE.OLEDB.12.0"
                        Case "MDB"
                            'MS Access 2003 and before
                            .Provider = "Microsoft.Jet.OLEDB.4.0"
                    End Select
                    .DataSource = _strOleDBName
                    .PersistSecurityInfo = False
                End With
            End If
            Return cnBuilder.ConnectionString
        Catch ex As Exception
            'Log error 
            Dim el As New Log.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Function cnnOLDEB Error")
            Return Nothing
        End Try
    End Function
    Public Function OfdAccess() As String
        'Purpose:       Allows the user to select which
        '               OleDb(MS - Access) database. 
        '               they wish to use.
        'Parameters:    None
        'Returns:       The complete path to the database.

        ' Create an instance of the open file dialog box.
        Dim openFileDialog1 As OpenFileDialog = New OpenFileDialog,
            strFileName As String
        'Set filter options for MS-Access
        With openFileDialog1
            .Filter = "MS Access 2005 and above (*.accdb)|*.accdb|" &
                                 "MS Access <=2003(*.mdb)|*.mdb|" &
                                 "All Files(*.*)|*.*"

            .FilterIndex = 1
            .Multiselect = False
        End With
        ' Call the ShowDialog method to show the dialogbox.
        Dim UserClickedOK = openFileDialog1.ShowDialog

        ' Process input if the user clicked OK.
        If (UserClickedOK = DialogResult.OK) Then
            'Get the filename for connection string
            strFileName = openFileDialog1.FileName
            Return strFileName
        Else
            Return Nothing
        End If
    End Function
    Public Function LongestStrInCol(ByRef strCnn As String,
                                    ByVal strTable As String,
                                    ByVal strColName As String) _
                                    As String
        'Purpose:       Find the longest record in the column for
        '               this table.
        'Parameters:    strCnn - A properly built connection string
        '               strTable - Name of the table whose records
        '                          We are interested in.
        '               strColName - Column we are searching
        'Returns:       The Longest Record as a string
        Dim strSQL As String
        Using cnn As New OleDbConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    'Find longest string in the [strColName] Column
                    strSQL = "SELECT [" & strColName & "] FROM " &
                        strTable & " WHERE len([" & strColName &
                        "]) = (SELECT max(len([" & strColName _
                        & "])) FROM " & strTable & ");"
                    'strSQL = "SELECT TOP 1 WITH TIES " & strColName & _
                    '         ", DATALENGTH(" & strColName & " ) " & _
                    '         " AS TextLength FROM [" & strTable & "]" & _
                    '         "ORDER BY TextLength DESC;"
                    Dim da As New OleDbDataAdapter(strSQL, cnn),
                        ds As New DataSet
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
    Public Function Lst_Access_Tables(ByVal strCnn As String) _
                                      As List(Of String)
        'Purpose:       Lists the tables inside the selected database
        'Parameters:    strCnn As String
        'Returns:       Creates a list of tables for this Access DB
        Dim listTables As List(Of String) = New List(Of String),
            userTables As DataTable = Nothing,
            I As Long
        Using cnn As New OleDbConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    '---------------------Routine Unusable OLEDB permissions--------------------
                    'Get database table names from selected MS-Access
                    'database.
                    'strSQL = "SELECT MSysObjects.[Name] AS table_name " & _
                    '         "FROM MSysObjects WHERE (((Left([Name],1))<>'~') " & _
                    '         "AND ((Left([Name],4))<>'MSys') " & _
                    '         "AND ((MSysObjects.[Type]) In (1,4,6)) AND " & _
                    '         "((MSysObjects.[Flags])=0)) order by MSysObjects.[Name];"
                    'Dim cmd As OleDbCommand = New OleDbCommand(strSQL, cnn)
                    'Dim dr As OleDbDataReader = cmd.ExecuteReader(CommandBehavior.Default)
                    'While (dr.Read())
                    'listTables.Add(dr(0).ToString())
                    'End While
                    'Set table list as combobox’s datasource
                    'dr.Close()
                    'cmd.Dispose()
                    '---------------------------------------------------------------------------
                    '-------------------------------Routine Works-------------------------------
                    userTables = cnn.GetSchema("Tables",
                                               New String() _
                                               {Nothing, Nothing,
                                                Nothing, "TABLE"})
                    cnn.Close()

                    For I = 0 To userTables.Rows.Count - 1
                        listTables.Add(userTables.Rows(I)(2).ToString)
                    Next
                    '---------------------------------------------------------------------------
                End If
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
            End Try
        End Using
        Return listTables
    End Function
    Public Function Get_OleDB_Tbl_Rel(ByVal strCnn As String,
                                      ByVal strTbl As String) _
                                      As DataTable
        'Purpose:       Find all relationship information for 
        '               the selected database table.
        'Parameters:    strCnn As String - A properly built 
        '                                  OleDB connection string.
        '               strTbl As String - The table we are finding 
        '                                  the relationships for.
        'Returns:       A DataTable of Relationships.  

        Dim schemaTable As DataTable,
            restrictions() As String,
            dr As DataRow
        'Remove DataTable Columns from dt_Rel
        Dim columncount As Integer = dt_Rel.Columns.Count - 1
        For i = columncount To 0 Step -1
            dt_Rel.Columns.RemoveAt(i)
        Next
        dt_Rel.Clear()
        'Creating New columns for dt_Rel
        dt_Rel.Columns.Add("PK_Tbl_Name",
                         Type.GetType("System.String"))
        dt_Rel.Columns.Add("PK_Col_Name",
        Type.GetType("System.String"))
        dt_Rel.Columns.Add("FK_Tbl_Name",
        Type.GetType("System.String"))
        dt_Rel.Columns.Add("FK_Col_Name",
        Type.GetType("System.String"))
        Using cnn As New OleDbConnection(strCnn)
            Try
                With cnn
                    .Open()
                    'This restriction set only returns the relationships
                    'for strTbl.
                    restrictions = {Nothing, Nothing, Nothing,
                                    Nothing, Nothing, strTbl}

                    'No restrictions returns all table relationships
                    'in the database.
                    'restrictions = {}
                    schemaTable = .GetOleDbSchemaTable(
                                  OleDbSchemaGuid.Foreign_Keys,
                                  restrictions)
                    .Close()
                    'Loop through the schemaTable 
                    For RowCount = 0 To schemaTable.Rows.Count - 1
                        If InStr(UCase(schemaTable.Rows(RowCount) ! _
                                       PK_TABLE_NAME.ToString), "MSYS") Then
                            Continue For
                        Else
                            dr = dt_Rel.NewRow
                            dr.Item("PK_Tbl_Name") =
                                schemaTable.Rows(RowCount) ! _
                                PK_TABLE_NAME.ToString()
                            dr.Item("PK_Col_Name") =
                            schemaTable.Rows(RowCount) ! _
                            PK_COLUMN_NAME.ToString()
                            dr.Item("FK_Tbl_Name") =
                            schemaTable.Rows(RowCount) ! _
                            FK_TABLE_NAME.ToString()
                            dr.Item("FK_Col_Name") =
                            schemaTable.Rows(RowCount) ! _
                            FK_COLUMN_NAME.ToString()
                            dt_Rel.Rows.Add(dr)
                        End If
                    Next RowCount
                    'Debug Purposes
                    'Call Dt_Read(dt_Rel)
                End With
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                Return Nothing
            End Try
        End Using
        Return dt_Rel
    End Function
    Public Function Get_OleDB_Tbl_Schema(ByVal strCnn As String,
                                         ByVal strTbl As String) _
                                         As DataTable
        'Purpose:       Creates a list of table columns(fields).
        '               We use the generated list to make the 
        '               TabPages to insert, or edit data for the table.
        'Parameters:    strCnn As String - A properly build OLEDB 
        '                                  connection string.
        '               strTbl As String - The name of the database
        '                                  database table whose schema
        '                                  we need to use.
        'Returns:       DataTable of Column Schema Information
        Dim cmd As New OleDbCommand(),
            cnn As New OleDbConnection(),
            myReader As OleDbDataReader,
            schemaTable As DataTable,
            strSQL As String

        strSQL = "SELECT * FROM [" & strTbl & "];"

        With cnn
            Try
                .ConnectionString = strCnn
                .Open()
                With cmd
                    .Connection = cnn
                    .CommandText = strSQL
                    'myReader = .ExecuteReader(CommandBehavior.KeyInfo)
                    myReader = .ExecuteReader(CommandBehavior.SchemaOnly)
                    schemaTable = myReader.GetSchemaTable()
                    'Debug Purposes
                    'Call dt_Read(schemaTable)
                    myReader.Close()
                    Return schemaTable
                End With
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                Return Nothing
            End Try
            .Close()
        End With
    End Function
    Public Function Get_OleDB_Tbl_Records(ByVal strCnn As String,
                                          ByVal strTable As String) _
                                      As DataTable
        'Purpose:       Creates a DataTable containing Rows(Records)
        '               for the desired table name.
        'Parameters:    strCnn As String - A properly build OLEDB 
        '                                  connection string.
        '               strTbl As String - The name of the table whose
        '                                  data we wish to use.
        'Returns:       A DataTable of Records.
        Dim strSQL As String
        Using cnn As New OleDbConnection(strCnn)
            Try
                If cnn.State = ConnectionState.Open Then
                    cnn.Close()
                Else
                    cnn.Open()
                    'Get all Information from 
                    'selected Server's Database table
                    strSQL = "SELECT * FROM [" & strTable & "];"
                    Dim dt As New DataTable
                    Using dad As New OleDbDataAdapter(strSQL, cnn)
                        dad.Fill(dt)
                    End Using
                    cnn.Close()
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
        End Using
    End Function
    Public Function Dt_AccessTblRelationships(ByVal strCnn As String,
                                          ByVal strTbl As String) _
                                          As DataTable
        'Purpose:       Find all relationships for the database table entered
        'Parameters:    strCnn As String - Properly built MS Access 
        '                                  connection string
        '               strTbl As String - Database relationships for the entered
        '                                  table.
        'Returns:       A DataTable od Relationships
        Dim schemaTable As DataTable,
            restrictions() As String,
            dr As DataRow
        Call Dt_Remove_Rel()
        dt_Rel.Clear()
        'Creating New columns for dt_Rel
        dt_Rel.Columns.Add("PK_Tbl_Name",
                         Type.GetType("System.String"))
        dt_Rel.Columns.Add("PK_Col_Name",
                           Type.GetType("System.String"))
        dt_Rel.Columns.Add("FK_Tbl_Name",
                           Type.GetType("System.String"))
        dt_Rel.Columns.Add("FK_Col_Name",
                           Type.GetType("System.String"))

        Using cnn As New OleDbConnection(strCnn)
            Try
                With cnn
                    .Open()
                    restrictions = {Nothing, Nothing, Nothing,
                                    Nothing, Nothing, strTbl}
                    schemaTable = .GetOleDbSchemaTable(
                                  OleDbSchemaGuid.Foreign_Keys,
                                  restrictions)
                    .Close()
                    'Call dt_Read(schemaTable)
                    'Loop through the schemaTable 
                    For RowCount = 0 To schemaTable.Rows.Count - 1
                        If InStr(UCase(schemaTable.Rows(RowCount) ! _
                                       PK_TABLE_NAME.ToString), "MSYS") Then
                            Continue For
                        Else
                            dr = dt_Rel.NewRow
                            dr.Item("PK_Tbl_Name") =
                                schemaTable.Rows(RowCount) ! _
                                PK_TABLE_NAME.ToString()
                            dr.Item("PK_Col_Name") =
                                schemaTable.Rows(RowCount) ! _
                                PK_COLUMN_NAME.ToString()
                            dr.Item("FK_Tbl_Name") =
                                schemaTable.Rows(RowCount) ! _
                                FK_TABLE_NAME.ToString()
                            dr.Item("FK_Col_Name") =
                                schemaTable.Rows(RowCount) ! _
                                FK_COLUMN_NAME.ToString()
                            dt_Rel.Rows.Add(dr)
                        End If
                    Next RowCount
                    'Debug Purposes
                    'Call dt_Read(dt_Rel)
                End With
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                Return Nothing
            End Try
        End Using
        Return dt_Rel
    End Function
End Class
