Imports System.Data.OleDb

Public Class MRM_Ctrls
    Inherits System.Windows.Forms.Control
    Private objAccessDB As New MRM_OleDB
    Private ReadOnly Msg As String
    Public m_dt As New DataTable

#Region "Public Properties"
    Public Property dtTbls() As DataTable
        Get
            Return m_dt
        End Get
        Set(ByVal value As DataTable)
            m_dt = value
        End Set
    End Property
#End Region
#Region "Explanation of Class"
    'Purpose:       Where I keep all dynamically
    '               created windows form controls, Methods,
    '               and so forth.
    'Methods:
    '   Private:    
    '               Functions:
    '                   CtrlPic
    '                   Det_CboBox_Width
    '                   Det_T_Box_Width
    '                   Ins_Btn_Ctrl
    '                   Ins_Lbl_Ctrl
    '                   Ins_Pic_Ctrl
    '                   Ins_CboBox_Ctrl
    '                   Ins_T_Box_Ctrl
    '   Public:
    '       Properties: 
    '           DtTbls
    '       Functions:
    '           Create_Access_DE_Form
    '           Create_MSSQL_DE_Form

    'Schema Indexes:
    '       0 - ColumnName
    '       1 - ColumnOrdinal	
    '       2 - ColumnSize	
    '       3 - NumericPrecision
    '       4 - NumericScale
    '       5 - DataType	
    '       6 - ProviderType	
    '       7 - IsLong
    '       8 - AllowDBNull
    '       9 - IsReadOnly
    '      10 - IsRowVersion
    '      11 - IsUnique
    '      12 - IsKey
    '      13 - IsAutoIncrement
    '      14 - BaseSchemaName	
    '      15 - BaseCatalogName	
    '      16 - BaseTableName
    '      17 - BaseColumnName
#End Region

#Region "Class only Accessible Methods"

    Private Function CtrlPic(ByVal Pic As PictureBox,
                             ByVal T_Box As TextBox) _
                             As PictureBox
        Dim strPath = T_Box.Name.ToString,
            ofdImage As New System.Windows.Forms.OpenFileDialog
        'Set the OpenFileDialog properties
        With ofdImage
            .Filter = "Bitmap Files (*.bmp;*'dib)|*.bmp;*.dib|" &
                      "JPEG (*.jpg;*.jpeg;*.jpe;*.jfif)" &
                            "|*.jpg;*.jpeg;*.jpe;*.jfif|" &
                      "GIF (*.gif)|*.gif|" &
                      "TIFF (*.tif; *.tiff)|*.tif;*.tiff|" &
                      "PNG (*.png)|*.png|" &
                      "ICO (*.ico)|*.ico|" &
                      "All Picture Files |*.gif;*.jpg;*.jpeg;" &
                            "*.jpe;*.jfif;*.png;*.bmp;*.dib;" &
                            "*.wmf;*.art;*.ico|" &
                      "All Files (*.*)|*.*"
            .FilterIndex = 1
            .FileName = Nothing
            .InitialDirectory = "C:\"
            .Title = "Browse Picture Files"
        End With
        With Pic
            If ofdImage.ShowDialog() = Windows.Forms.DialogResult.OK Then
                .Image = Image.FromFile(ofdImage.FileName)
                For Each ctrl As Control In FrmMain.tbctrlTblPages.Controls.Find(strPath,
                                                                  True)
                    If InStr(T_Box.Name.ToString, "PhotoPath") Then
                        ctrl.Select()
                        ctrl.Text = ofdImage.FileName
                    End If
                Next
            End If
        End With
        Return Pic
    End Function
    Private Function Det_CboBox_Width(ByVal ctrl As ComboBox, s As String) As Long
        Dim g As Graphics = ctrl.CreateGraphics
        ctrl.Text = s
        ctrl.Width = (g.MeasureString(ctrl.Text,
                                      ctrl.Font).Width) + 15
        'Debug.Print("DetTextWidth by Record:" & _
        'ctrl.Width.ToString)
        Return ctrl.Width
    End Function
    Private Function Det_T_Box_Width(ByVal ctrl As TextBox,
                                     ByVal s As String) _
                                     As Long
        Dim g As Graphics = ctrl.CreateGraphics
        ctrl.Text = s
        ctrl.Width = (g.MeasureString(ctrl.Text,
                                      ctrl.Font).Width) + 15
        'Debug.Print("DetTextWidth by Record:" & _
        'ctrl.Width.ToString)
        Return ctrl.Width
    End Function
    Private Function Ins_Btn_Ctrl(ByVal ctrl As Button) As Control
        With ctrl
            .Anchor = AnchorStyles.Top
            .Anchor = AnchorStyles.Left
            .AutoSize = True
            .BackColor = Color.FromKnownColor(KnownColor.Control)
            .ForeColor = Color.Black
            .FlatStyle = FlatStyle.Standard
            .Name = "btnPhoto"
            .Text = "Browse Photo"
        End With
        Return ctrl
    End Function
    Private Function Ins_Lbl_Ctrl(ByVal ctrl As Label,
                                  ByVal strColName As String,
                                  ByVal strKey As String,
                                  ByVal strColDT As String,
                                  ByVal strLen As String,
                                  ByVal strIsNull As String) _
                                  As Control
        Dim strText As String
        If strKey = String.Empty Then
            'Create Label
            strText = strColName & ": [ " &
                    strColDT & "(" &
                    strLen & "), " &
                    strIsNull & " ]"
        Else
            strText = strColName & ": [" &
                    strKey & ", " &
                    strColDT & "(" &
                    strLen & "), " &
                    strIsNull & "]"
        End If

        With ctrl
            .Anchor = AnchorStyles.Top
            .Anchor = AnchorStyles.Left
            .AutoSize = True
            .BackColor = Color.FromKnownColor(KnownColor.Control)
            .Font = New Drawing.Font("Comic Sans MS",
                               9, FontStyle.Regular)
            .ForeColor = Color.Black
            .FlatStyle = FlatStyle.Standard
            .Name = "lbl" & strColName
            .Text = strText
        End With
        Return ctrl
    End Function
    Private Function Ins_Pic_Ctrl(ByVal ctrl As PictureBox, _
                                  ByVal strColName As String, _
                                  ByVal bs As BindingSource) _
                                  As Control
        With ctrl
            .Anchor = AnchorStyles.Top
            .Anchor = AnchorStyles.Left
            .Size = New Size(431, 272)
            .SizeMode = PictureBoxSizeMode.Zoom
            .BackColor = Color.FromKnownColor(KnownColor.Control)
            .BorderStyle = BorderStyle.Fixed3D
            .ForeColor = Color.Black
            .Name = "img" & strColName
            .DataBindings.Add(New Binding("Image", bs, strColName, True))
        End With
        Return ctrl
    End Function
    Private Function Ins_CboBox_Ctrl(ByVal ctrl As ComboBox,
                                     ByVal strColName As String,
                                     ByVal strCnn As String,
                                     ByVal strPKTbl As String,
                                     ByVal strPKCol As String,
                                     ByVal strFKTbl As String,
                                     ByVal strFKCol As String,
                                     ByVal bs As BindingSource,
                                     ByVal s As String) _
                                     As Control
        Dim dtaSource As DataTable,
        lngWidth As Long
        dtaSource = objAccessDB.Get_OleDB_Tbl_Records(strCnn, strPKTbl)
        With ctrl
            .Font = New Drawing.Font("Comic Sans MS",
                                     10, FontStyle.Regular)
            If String.IsNullOrEmpty(s) Then
                lngWidth = 121
            Else
                lngWidth = Det_CboBox_Width(ctrl, s)
            End If
            .Anchor = AnchorStyles.Top
            .Anchor = AnchorStyles.Left
            .AutoSize = True
            .BackColor = Color.FromKnownColor(KnownColor.Control)
            .ForeColor = Color.Black
            .Name = "cbo" & strColName
            'AddHandler .Paint,
            '    Function(sender, e) _
            'dtaSource
            .DataBindings.Add(New Binding("SelectedValue",
                        bs, strColName,
                        True))
            .ValueMember = strPKCol
            .DisplayMember = dtaSource.Columns.Item(1).ToString
            .DataSource = dtaSource
            .Width = lngWidth
        End With
        Return ctrl
    End Function
    Private Function Ins_T_Box_Ctrl(ByVal ctrl As TextBox,
                                    ByVal strColName As String,
                                    ByVal bs As BindingSource,
                                    ByVal s As String) _
                                    As Control
        Dim lngWidth As Long
        With ctrl
            .Font = New Drawing.Font("Comic Sans MS",
                                     10, FontStyle.Regular)
            lngWidth = Det_T_Box_Width(ctrl, s)
            If lngWidth > (Screen.PrimaryScreen.WorkingArea.Width) Then
                .Width = lngWidth / 3
                .Multiline = True
                .Height = 40
                .ScrollBars = ScrollBars.Vertical
            Else
                .Width = lngWidth
                .Multiline = False
                .Height = 20
                .Size = New System.Drawing.Size(lngWidth, 40)
            End If
            .Anchor = AnchorStyles.Top
            .Anchor = AnchorStyles.Left
            .AutoSize = True
            .BackColor = Color.FromKnownColor(KnownColor.Control)
            .BorderStyle = BorderStyle.None
            .ForeColor = Color.Black
            .Name = "txt" & strColName
            .DataBindings.Add(New Binding("Text", bs, strColName, True))
        End With
        Return ctrl
    End Function
    Private Function Schema_Col_Info(ByVal dt_Schema As DataTable,
                                    ByVal dt_Rel As DataTable,
                                    ByVal i_Row As Long) As String
        'Purpose:       Creates an Information label using the schema DataTable
        'Parameters:    dt_Schema As DataTable - DataTable containing the table's schema
        '               dt_Rel As DataTable - DataTable showing relationships
        '               i_Row As Long - The DataTable's Row number 
        '                               for which we are retrieving schema data
        'Returns:       An Information lable as a String

        Dim strLabel As String = "",
            strKey As String = "",
            fRow As DataRow,
            expression As String,
            searchedValue As String = Nothing

        'IsKey?
        Select Case dt_Schema.Rows(i_Row).Item(12)
            Case True
                'IsKey an AutoNumber?
                Select Case dt_Schema.Rows(i_Row).Item(13)
                    Case True
                        strKey = "PK(Auto)"
                    Case False
                        strKey = "PK"
                End Select
                strLabel = dt_Schema.Rows(i_Row).Item(0).ToString & ": [" &
                           strKey & ", " &
                           Mid(dt_Schema.Rows(i_Row).Item(5).ToString, 8) & "(" &
                           dt_Schema.Rows(i_Row).Item(2).ToString & ")]"
            Case False
                'Is Foreign Key?
                expression = "FK_Col_Name = '" &
                     dt_Schema.Rows(i_Row).Item(0).ToString & "'"
                fRow = dt_Rel.Select(expression).FirstOrDefault
                If Not fRow Is Nothing Then
                    searchedValue = fRow.Item("FK_Col_Name")
                End If
                Select Case searchedValue = dt_Schema.Rows(i_Row).Item(0).ToString()
                    Case True
                        strKey = "FK"
                        strLabel = dt_Schema.Rows(i_Row).Item(0).ToString & ": [" &
                                    strKey & ", " &
                                    Mid(dt_Schema.Rows(i_Row).Item(5).ToString, 8) & "(" &
                                    dt_Schema.Rows(i_Row).Item(2).ToString & ")]"
                    Case False
                        strKey = ""
                        'No Primary or Foreign Keys
                        strLabel = dt_Schema.Rows(i_Row).Item(0).ToString & ": [" &
                                   Mid(dt_Schema.Rows(i_Row).Item(5).ToString, 8) & "(" &
                                   dt_Schema.Rows(i_Row).Item(2).ToString & ")]"
                End Select
        End Select
        Return strLabel
    End Function
    Private Function Schema_Ctrl_Type(ByVal strCnn As String,
                                      ByVal dt_Schema As DataTable,
                                      ByVal dt_Rel As DataTable,
                                      ByVal bs As BindingSource,
                                      ByVal i_Row As Long) As Control
        'Purpose:       Determines which control to use based on the 
        '               column DataType 
        'Parameters:    dt_Schema As DataTable - DataTable containing
        '                                        the table's schema
        '               dt_Rel As DataTable - DataTable showing relationships
        '               bs As BindingSource - BindingSource for the control
        'Returns:       The proper control.

        Dim strLongestRec As String,
            b_FK As Boolean,
            ctrl As Control,
            NewColor As Color = Color.FromKnownColor(KnownColor.Control),
            strColName As String,
            strColType As String,
            strFKTbl As String = Nothing,
            strFKCol As String = Nothing,
            strPKTbl As String = Nothing,
            strPKCol As String = Nothing

        'Get Column data type
        strColType = Mid(dt_Schema.Rows(i_Row).Item(5).ToString, 8)
        strColName = dt_Schema.Rows(i_Row).Item(0).ToString

        'Look for Foreign Key data
        If objAccessDB.FindFK(dt_Rel, strColName) Then
            b_FK = True
        Else
            b_FK = False
        End If

        'Create Form Controls
        strLongestRec = objAccessDB.LongestStrInCol(strCnn,
                                    dt_Schema.Rows(i_Row).Item(16).ToString,
                                    dt_Schema.Rows(i_Row).Item(17).ToString)
        Select Case UCase(strColType)
            Case "INT16", "INT32", "INT64", "STRING"
                'Foreign Key?
                If objAccessDB.FindFK(dt_Rel, strColName) Then
                    strPKTbl = dt_Rel.Rows(0).Item("PK_Tbl_Name").ToString
                    strPKCol = dt_Rel.Rows(0).Item("PK_Col_Name").ToString
                    strFKTbl = dt_Rel.Rows(0).Item("FK_Tbl_Name").ToString
                    strFKCol = dt_Rel.Rows(0).Item("FK_Col_Name").ToString
                    'Foreign key; use ComboBox - DEACTIVATED: 18_May-2023 
                    strLongestRec =
                        objAccessDB.LongestStrInCol(strCnn, strPKTbl, "Country")
                    'ctrl = New ComboBox
                    'ctrl = Ins_CboBox_Ctrl(ctrl, strColName,
                    'strCnn, strPKTbl,
                    'strPKCol, strFKTbl,
                    'strFKCol, bs, strLongestRec)
                    ctrl = New TextBox
                    ctrl = Ins_T_Box_Ctrl(ctrl, strColName,
                                            bs, strLongestRec)
                    Return ctrl
                Else
                    'Use TextBox
                    ctrl = New TextBox
                    ctrl = Ins_T_Box_Ctrl(ctrl, strColName,
                                            bs, strLongestRec)
                    Return ctrl
                End If
            Case "IMAGE"
                ctrl = New PictureBox
                ctrl = Ins_Pic_Ctrl(ctrl, strColName, bs)
            Case Else
                'Use TextBox
                ctrl = New TextBox
                ctrl = Ins_T_Box_Ctrl(ctrl, strColName,
                                        bs, strLongestRec)
        End Select
        Return ctrl
    End Function

#End Region

#Region "Publically Accessible Methods"
    Public Function Ins_DGV_Ctrl(ByVal ctrl As DataGridView,
                                 ByVal dt As DataTable) _
                                 As Control
        With ctrl
            .AutoGenerateColumns = True
            .AutoSizeColumnsMode =
                DataGridViewAutoSizeColumnsMode.AllCells
            .ColumnHeadersHeightSizeMode =
                DataGridViewColumnHeadersHeightSizeMode.AutoSize
            .Dock = DockStyle.Fill
            .MultiSelect = False
            .ReadOnly = True
            .DataSource = dt
            .ScrollBars = ScrollBars.Both
        End With
        Return ctrl
    End Function
    Public Function Ins_TabPage_Ctrl(ByVal ctrlTabPg As TabPage,
                                     ByVal ctrlNav As Control,
                                     ByVal ctrlTLP As Control,
                                     ByVal cntxMenu As ContextMenuStrip,
                                     ByVal strTbl As String) _
                                     As Control
        With ctrlTabPg
            .AutoScroll = True
            .Name = strTbl
            .Text = strTbl
            .ContextMenuStrip = cntxMenu
            .Controls.Add(ctrlNav)
            .Controls.Add(ctrlTLP)
        End With
        Return ctrlTabPg
    End Function
#End Region

#Region "Dynamically Created Data Entry Forms"
    Public Function Create_Access_DE_Form(ByVal ctrlNav As BindingNavigator,
                                          ByVal bindSrc As BindingSource,
                                          ByVal strCnn As String,
                                          ByVal strTable As String) _
                                          As TableLayoutPanel
        'Purpose:       Builds a data entry from based on 
        '               the Access database table's schema
        'Parameters:    ctrlNav As BindingNavigator
        '               bindSrc As BindingSource
        '               strCnn As String
        '               strTable As String
        'Returns:       A Data Entry Form containing all the
        '               controls needed.
        Dim dynTLP As New TableLayoutPanel,
            dtSchema As New DataTable,
            objAccessDB As New MRM_OleDB,
            dtRecords As New DataTable,
            dt_Rel As DataTable,
            I As Long,
            szWidth As Long,
            strPath As String

        'Variables for table schema
        Dim strColName As String,
            strColDT As String,
            strLen As String,
            strIs_Null As String = "Null",
            strKey As String,
            strLongestRec As String,
            ctrlLabel As New Label,
            ctrlButton As New Button,
            ctrlImage As New PictureBox,
            ctrlTB As New TextBox,
            NewColor As Color = Color.FromKnownColor(KnownColor.Control)

        'Get database Table's Relationships
        dt_Rel = objAccessDB.Dt_AccessTblRelationships(strCnn, strTable)
        'Get database table schema
        dtSchema = objAccessDB.Get_OleDB_Tbl_Schema(strCnn, strTable)
        'Get database table's data
        dtRecords = objAccessDB.Get_OleDB_Tbl_Records(strCnn, strTable)
        'Set bindSrc DataSource
        bindSrc.DataSource = dtRecords
        dtTbls = dtRecords
        'frmLogIn.dsTbls.Tables.Add(dtRecords)
        'Set BindingNavigator Control to BindingSource
        ctrlNav.BindingSource = bindSrc
        'Build TableLayoutPanel
        With dynTLP
            .Name = strTable
            .SuspendLayout()
            .Controls.Clear()
            .AutoSize = True
            .AutoSizeMode = AutoSizeMode.GrowAndShrink
            .BackColor = NewColor
            .CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
            .ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            .Location = New Point(0, 25)
            .RowCount = dtSchema.Rows.Count - 1
            'Loop throught the Table schema and
            'extract the data we need.
            For I = 0 To dtSchema.Rows.Count - 1
                'MS-Access table's ColumnName 
                strColName = dtSchema.Rows(I).Item(0).ToString
                'IsKey?
                If UCase(dtSchema.Rows(I).Item(12).ToString) = "TRUE" Then
                    'IsKey an AutoNumber
                    If UCase(dtSchema.Rows(I).Item(13).ToString) = "TRUE" Then
                        strKey = "PK(Auto)"
                        strIs_Null = "Not Null"
                    Else
                        strKey = "PK"
                        strIs_Null = "Not Null"
                    End If
                Else
                    'Is the Column a Foreign Key?
                    If objAccessDB.FindFK(objAccessDB.dt_Rel, strColName) Then
                        strKey = "FK"
                        strIs_Null = "Null"
                    Else
                        strKey = ""
                    End If
                End If
                'MS-Access ColumnSize
                strLen = dtSchema.Rows(I).Item(2).ToString
                'MS-Access Column DataType
                strColDT = Mid(dtSchema.Rows(I).Item(5).ToString, 8)
                strLongestRec = objAccessDB.LongestStrInCol(strCnn,
                                                            strTable,
                                                            strColName)
                szWidth = Len(strColName + strColDT + strLen)
                'Placing controls on Column DataType
                Select Case strColDT
                    Case "int16", "Int32", "Int64", "String",
                         "DateTime", "Decimal", "Int16", "Double"
                        'Place Label control with Column Name, DataType,
                        'DataType Length and are Nulls allowed
                        ctrlLabel = New Label
                        ctrlLabel = Ins_Lbl_Ctrl(ctrlLabel, strColName,
                                                 strKey, strColDT,
                                                 strLen, strIs_Null)
                        'Add control to dynTLP
                        .Controls.Add(ctrlLabel, 1, I)
                        'TextBox to display the column's data
                        ctrlTB = New TextBox
                        ctrlTB = Ins_T_Box_Ctrl(ctrlTB, strColName,
                                                bindSrc, strLongestRec)
                        If InStr(strColName, "PhotoPath") Then
                            strPath = strColName
                        End If
                        'Adding control to dynTLP
                        .Controls.Add(ctrlTB, 2, I)
                    Case "image"
                        'Place Label control with Column Name, DataType,
                        'DataType Length and are Nulls allowed
                        ctrlLabel = New Label
                        'ctrlLabel = Ins_Lbl_Ctrl(ctrlLabel, strColName, _
                        '                        strColDT, strLen, _
                        '                        strIs_Null)
                        'Add control to dynTLP
                        .Controls.Add(ctrlLabel, 1, I)
                        'Picture box to display image
                        ctrlImage = New PictureBox
                        ctrlImage = Ins_Pic_Ctrl(ctrlImage,
                                                 strColName,
                                                 bindSrc)
                        'Adding control to dynTLP
                        .Controls.Add(ctrlImage, 2, I)
                        'Button to update image
                        ctrlButton = New Button
                        ctrlButton = Ins_Btn_Ctrl(ctrlButton)
                        AddHandler ctrlButton.Click,
                                     Function(sender, e) _
                                        CtrlPic(ctrlImage, ctrlTB)
                        'Adding control to dynTLP
                        .Controls.Add(ctrlButton, 3, I)
                End Select
            Next
            .ResumeLayout()
        End With
        Return dynTLP
    End Function
    Public Function Create_MSSQL_DE_Form(ByVal ctrlNav As BindingNavigator,
                                         ByVal bindSrc As BindingSource,
                                         ByVal strSQLCnn As String,
                                         ByVal strTbl As String) _
                                         As TableLayoutPanel
        'Purpose:       Builds a data entry form based on MSSQL database
        '               tables schema
        'Parameters:    ctrlNav As BindingNavigator
        '               bindSrc As BindingSource
        '               strCnn As String
        '               strTable As String
        'Returns:       A Data Entry Form containing all the
        '               controls needed.
        Dim dynTLP As New TableLayoutPanel,
            dtSchema As New DataTable,
            objAccessDB As New MRM_OleDB,
            objMSSQL_DB As New MRM_MSSQL,
            dtRecords As New DataTable,
            dt_SQL_Rel As DataTable,
            I As Long,
            szWidth As Long,
            strPath As String
        'Variables for table schema
        Dim strColName As String,
            strColDT As String,
            strLen As String,
            strIs_Null As String,
            strKey As String,
            strLongestRec As String,
            ctrlLabel As New Label,
            ctrlButton As New Button,
            ctrlImage As New PictureBox,
            ctrlTB As New TextBox,
            NewColor As Color = Color.FromKnownColor(KnownColor.Control)

        'Get Table Relationships
        dt_SQL_Rel = objMSSQL_DB.Dt_MSSQL_Tbl_Relationships(strSQLCnn, strTbl)
        'Get database table schema
        dtSchema = objMSSQL_DB.Dt_MSSQL_Tbl_Schema(strSQLCnn, strTbl)
        'Read MS_SQL Server Table Schema
        'Call objMSSQL_DB.dt_Viewer(dtSchema)
        'Get database table's data
        dtRecords = objMSSQL_DB.Dt_MSSQL_Retrive_Data(strSQLCnn, strTbl)
        'Read database table's reords
        'Call objMSSQL_DB.dt_Viewer(dtRecords)
        'Set bindSrc DataSource
        bindSrc.DataSource = dtRecords
        'dtTbls = dtRecords
        'Call objMSSQL_DB.dt_Viewer(dtTbls)
        'Set BindingNavigator Control to BindingSource
        ctrlNav.BindingSource = bindSrc
        'Build TableLayoutPanel
        With dynTLP
            .Name = strTbl
            .SuspendLayout()
            .Controls.Clear()
            .AutoSize = True
            .AutoSizeMode = AutoSizeMode.GrowAndShrink
            .BackColor = NewColor
            .CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
            .ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            .Location = New Point(0, 25)
            .RowCount = dtSchema.Rows.Count - 1
            'Loop throught the Table schema and
            'extract the data we need.
            For I = 0 To dtSchema.Rows.Count - 1
                strColName = dtSchema.Rows(I).Item(1).ToString
                'Is this a Primary Key Column?
                If objMSSQL_DB.Det_Tbl_PK(strSQLCnn,
                                          strTbl,
                                          strColName) Then
                    strKey = "PK"
                    strIs_Null = "Not Null"
                    'Is this a Foreign Key Column?
                ElseIf objMSSQL_DB.Det_Tbl_FK(strSQLCnn,
                                          strTbl,
                                          strColName) Then
                    strKey = "FK"
                    strIs_Null = "Null"
                Else
                    strKey = ""
                End If
                strColDT = dtSchema.Rows(I).Item(2).ToString
                strLen = dtSchema.Rows(I).Item(3).ToString
                'Find longest string in column
                strLongestRec = objMSSQL_DB.LongestStrInCol(strSQLCnn,
                                                            strTbl,
                                                            strColName)
                If (dtSchema.Rows(I).Item(4).ToString) = "YES" Then
                    strIs_Null = "Null"
                Else
                    strIs_Null = "Not Null"
                End If
                szWidth = Len(strColName + strColDT + strLen)
                'Placing controls on Column DataType
                Select Case strColDT
                    Case "int", "nvarchar", "ntext", "nchar", "varchar",
                         "datetime", "money", "smallint", "real", "bit"

                        'Place Label control with Column Name, DataType,
                        'DataType Length and are Nulls allowed
                        ctrlLabel = New Label
                        ctrlLabel = Ins_Lbl_Ctrl(ctrlLabel, strColName,
                                                 strKey, strColDT,
                                                 strLen, strIs_Null)

                        'Add control to dynTLP
                        .Controls.Add(ctrlLabel, 1, I)
                        'TextBox to display the column's data
                        ctrlTB = New TextBox
                        ctrlTB = Ins_T_Box_Ctrl(ctrlTB, strColName,
                                                bindSrc, strLongestRec)
                        If InStr(strColName, "PhotoPath") Then
                            strPath = strColName
                        End If
                        'Adding control to dynTLP
                        .Controls.Add(ctrlTB, 2, I)
                    Case "image"
                        'Place Label control with Column Name, DataType,
                        'DataType Length and are Nulls allowed
                        ctrlLabel = New Label
                        ctrlLabel = Ins_Lbl_Ctrl(ctrlLabel, strColName,
                                                 strKey, strColDT,
                                                 strLen, strIs_Null)
                        'Add control to dynTLP
                        .Controls.Add(ctrlLabel, 1, I)
                        'Picture box to display image
                        ctrlImage = New PictureBox
                        ctrlImage = Ins_Pic_Ctrl(ctrlImage,
                                                 strColName,
                                                 bindSrc)
                        'Adding control to dynTLP
                        .Controls.Add(ctrlImage, 2, I)
                        'Button to update image
                        ctrlButton = New Button
                        ctrlButton = Ins_Btn_Ctrl(ctrlButton)
                        AddHandler ctrlButton.Click,
                                        Function(sender, e) _
                                            CtrlPic(ctrlImage, ctrlTB)
                        'Adding control to dynTLP
                        .Controls.Add(ctrlButton, 3, I)
                End Select
            Next I
            .ResumeLayout()
        End With
        Return dynTLP
    End Function
#End Region
End Class
