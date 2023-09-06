Option Explicit On
Imports System.Data.OleDb
Public Class FrmMain
    'OleDB Connection string for MS_Access
    Public cnnOleDB As String,
           strSQLCnn As String = Nothing
    Private strAppPath As String = Application.StartupPath,
            strTbl As String,
            strCatalog As String,
            strCatTbl As String,
            objOleDB As New MRM_OleDB,
            objCtrls As New MRM_Ctrls,
            objSQL As New MRM_MSSQL,
            getOleDB As New MRM_OleDB,
            ctrlNav As New BindingNavigator(True),
            bindSrc As New BindingSource,
            dt_Rel As DataTable,
            dt_Schema As DataTable,
            dt_Records As DataTable,
            dsTbls As New DataSet,
            strMsg As String

    Private Function Ee_TabPage(ByVal ctrlTabPg As TabPage,
                                ByVal cntxMenu As ContextMenuStrip,
                                ByVal ctrlDGV As Control,
                                ByVal strTitle As String,
                                ByVal strTbl As String) _
                                As Control
        'Purpose:       Function to create Easter Egg TabPages
        'Parameters:    ctrlTabPg As TabPage  
        '               cntxMenu As ContextMenu
        '               ctrlDGV As Control
        '               strTitle as String
        '               strTbl As String
        'Returns:       A TabPage containing a populated Data Grid View

        With ctrlTabPg
            .AutoScroll = True
            .Name = strTitle & "-" & strTbl
            .Text = strTitle & "-" & strTbl
            .ContextMenuStrip = cntxOLEDB
            .Controls.Add(ctrlDGV)
        End With
        Return ctrlTabPg
    End Function

    Private Sub BindingNavigatorSaveItem_Click(ByVal sender As System.Object,
                                              ByVal e As System.EventArgs)
        Dim strTab As String,
            strSQL As String,
            da As OleDbDataAdapter,
            sb As OleDbCommandBuilder,
            Msg As String

        strTab = tbctrlTblPages.SelectedTab.Name

        'Search the current TabPage for a BindingNavigator control
        For Each c In Me.tbctrlTblPages.Controls.Item(strTab).Controls
            Dim oNav As BindingNavigator = TryCast(c, BindingNavigator)
            'If the control is a BindingNavigator
            If oNav IsNot Nothing Then
                oNav.BindingSource.EndEdit()
                Exit For
            End If
        Next
        'Connect to the database
        Using objConn As New OleDbConnection(cnnOleDB)
            Try
                If objConn.State = ConnectionState.Open Then
                    objConn.Close()
                Else
                    objConn.Open()
                    strSQL = "SELECT * FROM " & strTab 'CA2100 Warning- Need Parameterized query
                    da = New OleDbDataAdapter(strSQL, objConn)
                    'Build Commands for add, delete, and update
                    sb = New OleDbCommandBuilder(da)
                    'Update the selected table
                    da.Update(dt_Records)
                    Msg = "The " & strTab & " table from the " & cboOleDBTbls.Text &
                          " database has been successfully updated."
                    MessageBox.Show(Msg)
                End If
            Catch ex As Exception
                'Log error 
                Dim el As New Log.ErrorLogger
                el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                'Finally
                'objConn.Close()
            End Try
        End Using
    End Sub

    Private Function CtrlNavigator(ByVal ctrl As BindingNavigator,
                               ByVal strImgPath As String) _
                               As BindingNavigator
        'Purpose:       Places a new ToolStripButton on the 
        '               BindingNavigator ToolStrip Control
        'Parameters:    ctrl As BindingNavigator;
        '               strImgPath As String
        'Returns:       Modified BindingNavigator ToolStrip 
        Dim image1 As Bitmap =
                       CType(Image.FromFile(strImgPath, True), Bitmap), BindingNavigatorSaveItem As New ToolStripButton

        'Set the background color of the image to be transparent.
        BindingNavigatorSaveItem =
                                    New ToolStripButton(Nothing, image1,
                                    AddressOf BindingNavigatorSaveItem_Click) _
                                    With
                                        {
                                            .Image = New Bitmap(image1),
                                            .ImageScaling = ToolStripItemImageScaling.SizeToFit,
                                            .ImageTransparentColor = Color.Magenta
                                        }

        ' Show ToolTip text, set custom ToolTip text, and turn
        ' off the automatic ToolTips.
        ctrl.ShowItemToolTips = True
        BindingNavigatorSaveItem.ToolTipText = "Saves changes"
        BindingNavigatorSaveItem.AutoToolTip = False
        ' Add the button to the ToolStrip.
        ctrl.Items.Add(BindingNavigatorSaveItem)
        Return ctrl
    End Function

    Private Sub BtnAccessDB_Click(sender As Object,
                                  e As EventArgs) _
                                  Handles BtnAccessDB.Click
        'Purpose:       Allows the user to select which MS Access
        '               database they wish to use.
        'Parameters:    None
        'Returns:       Nothing

        getOleDB.GetOleDBName = objOleDB.OfdAccess()
        Me.txtDB.Text = getOleDB.GetOleDBName
        cnnOleDB = getOleDB.GetAccessCnn
        Me.cboOleDBTbls.DataSource = Nothing
        Me.cboOleDBTbls.DataSource = getOleDB.Lst_Access_Tables(cnnOleDB)
    End Sub

    Private Sub BtnFrmTbl_Click(sender As Object,
                                e As EventArgs) _
                                Handles BtnFrmTbl.Click
        'Create an instance of BindingNavigator control
        Dim ctrlNav As New BindingNavigator(True) _
                With {.Dock = DockStyle.Top},
            bindSrc As New BindingSource,
            dynTP As New TabPage,
            el As New Log.ErrorLogger,
            dynTLP As New TableLayoutPanel,
            frmDE As New MRM_Ctrls
        Cursor = Cursors.WaitCursor
        ctrlNav = CtrlNavigator(ctrlNav, strAppPath & "\Save.bmp")
        strTbl = cboOleDBTbls.SelectedValue
        'Get table schema
        dt_Schema = objOleDB.Get_OleDB_Tbl_Schema(cnnOleDB, strTbl)
        'Get Table Relationships
        dt_Rel = objOleDB.Get_OleDB_Tbl_Rel(cnnOleDB, strTbl)
        'Get table records
        dt_Records = objOleDB.Get_OleDB_Tbl_Records(cnnOleDB, strTbl)
        'Call objOleDB.dt_Read(dt_Records)
        bindSrc.DataSource = dt_Records
        dynTLP = frmDE.Create_Access_DE_Form(ctrlNav, bindSrc,
                                          cnnOleDB, strTbl)

        Try
            dynTP = objCtrls.Ins_TabPage_Ctrl(dynTP, ctrlNav,
                                          dynTLP, cntxOLEDB, strTbl)
            'Adding TabPage to Form
            Me.tbctrlTblPages.TabPages.Add(dynTP)
        Catch ex As Exception
            'Log error 
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
        End Try
        Cursor = Cursors.Default
    End Sub

    Private Sub FrmMain_KeyDown(sender As Object,
                                e As KeyEventArgs) _
                                Handles Me.KeyDown
        'Purpose:   Programmer's Tool    
        'Triggers:  Ctrl + Alt + S - Show Schema for the selected table.
        '           Ctrl + Alt + D - Show Records in the selected table.
        '           Ctrl + Alt + R - Shows Relationships to the selected table
        Dim tp As New TabPage,
                dgvCtrl As New DataGridView,
                dt_Log As New Log.ErrorLogger,
                strTitle As String = ""

        If Not String.IsNullOrEmpty(strTbl) Then
            Cursor = Cursors.WaitCursor
            Select Case e.KeyCode
                Case Keys.S And (e.Control And e.Alt)
                    'Ctrl + Alt + S
                    'dt_Schema = objOleDB.Get_OleDB_Tbl_Schema(cnnOleDB, strTbl) MODIFIED
                    If dt_Schema.Rows.Count <> 0 Then
                        dgvCtrl = New DataGridView
                        dgvCtrl = objCtrls.Ins_DGV_Ctrl(dgvCtrl, dt_Schema)
                        strTitle = "Schema"
                        Try
                            tp = Ee_TabPage(tp, cntxOLEDB, dgvCtrl,
                                            strTitle, strTbl)
                            Me.tbctrlTblPages.TabPages.Add(tp)
                        Catch ex As Exception
                            'Log error 
                            Dim el As New Log.ErrorLogger
                            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                        End Try
                    Else
                        strMsg = "No Schema information for table " & strTbl
                        MessageBox.Show(strMsg, "Database table Schema.",
                                        MessageBoxButtons.OK)
                    End If
                Case Keys.D And (e.Control And e.Alt)
                    'Ctrl + Alt + D
                    'dt_Records = objOleDB.Get_OleDB_Tbl_Records(cnnOleDB, strTbl) MODIFIED
                    If dt_Records.Rows.Count <> 0 Then
                        dgvCtrl = New DataGridView
                        dgvCtrl = objCtrls.Ins_DGV_Ctrl(dgvCtrl, dt_Records)
                        strTitle = "Data"
                        Try
                            tp = Ee_TabPage(tp, cntxOLEDB, dgvCtrl,
                                           strTitle, strTbl)
                            Me.tbctrlTblPages.TabPages.Add(tp)
                        Catch ex As Exception
                            'Log error 
                            Dim el As New Log.ErrorLogger
                            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                        End Try
                    Else
                        strMsg = "No records contained in table " & strTbl
                        MessageBox.Show(strMsg, "Database table Data.",
                                        MessageBoxButtons.OK)
                    End If
                Case Keys.R And (e.Control And e.Alt)
                    'Ctrl + Alt + R
                    'strTbl = cboOleDBTbls.SelectedValue
                    'dt_Rel = objOleDB.Get_OleDB_Tbl_Rel(cnnOleDB, strTbl)
                    If dt_Rel.Rows.Count <> 0 Then
                        dgvCtrl = New DataGridView
                        dgvCtrl = objCtrls.Ins_DGV_Ctrl(dgvCtrl, dt_Rel)
                        strTitle = "Relationships"
                        Try
                            tp = Ee_TabPage(tp, cntxOLEDB, dgvCtrl,
                                           strTitle, strTbl)
                            Me.tbctrlTblPages.TabPages.Add(tp)
                        Catch ex As Exception
                            'Log error 
                            Dim el As New Log.ErrorLogger
                            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
                        End Try
                    Else
                        strMsg = "No Relationships found for table " & strTbl
                        MessageBox.Show(strMsg, "Database table relationships.",
                                        MessageBoxButtons.OK)
                    End If
            End Select
        End If
        Cursor = Cursors.Default
    End Sub

    Private Sub FrmMain_Load(sender As Object,
                             e As EventArgs) _
                             Handles Me.Load
        Me.KeyPreview = True
        Me.tbctrlTblPages.TabPages.Clear()
    End Sub

#Region "Cntx Control Click Events"
    Private Sub AccessRemoveAllToolStripMenuItem_Click(sender As Object,
                                                           e As EventArgs) _
                                                           Handles AccessRemoveAllToolStripMenuItem.Click
        Cursor = Cursors.WaitCursor
        tbctrlTblPages.TabPages.Clear()
        Cursor = Cursors.Default
    End Sub
    Private Sub AccessRemoveToolStripMenuItem_Click(sender As Object,
                                                        e As EventArgs) _
                                                        Handles AccessRemoveToolStripMenuItem.Click
        Dim strTable As String
        strTable = tbctrlTblPages.SelectedTab.Name
        'Remove the named table from the DataSet
        Cursor = Cursors.WaitCursor
        tbctrlTblPages.TabPages.Remove(tbctrlTblPages.SelectedTab)
        Cursor = Cursors.Default
    End Sub
    Private Sub RemoveAllSQLToolStripMenuItem_Click(sender As Object,
                                                    e As EventArgs) _
                                                    Handles RemoveAllSQLToolStripMenuItem.Click
        Cursor = Cursors.WaitCursor
        tbctrlTblPages.TabPages.Clear()
        Cursor = Cursors.Default
    End Sub
    Private Sub RemoveSQLToolStripMenuItem_Click(sender As Object,
                                                 e As EventArgs) _
                                                 Handles RemoveSQLToolStripMenuItem.Click
        Dim strTable As String
        strTable = tbctrlTblPages.SelectedTab.Name
        'Remove the named table from the DataSet
        Cursor = Cursors.WaitCursor
        tbctrlTblPages.TabPages.Remove(tbctrlTblPages.SelectedTab)
        Cursor = Cursors.Default
    End Sub
#End Region

    Private Sub BtnCnnSQL_Click(sender As Object,
                                e As EventArgs) _
                                Handles BtnCnnSQL.Click
        Select Case Me.cboSQLAuth.SelectedIndex
            Case 0
                strSQLCnn = objSQL.TrustedConnection(txtServer.Text, "master")
            Case 1
                strSQLCnn =
                    objSQL.StandardSecurity(txtServer.Text,
                                            "master", txtSQLLogin.Text,
                                            txtSQL_PW.Text)
        End Select
        CboSQLDB.DataSource = objSQL.MSSQL_DBs(strSQLCnn)
        Me.CboSQLTbls.DataSource = objSQL.MSSQL_DB_Tbls(strSQLCnn)
    End Sub

    Private Sub CboSQLDB_SelectedIndexChanged(sender As Object,
                                              e As EventArgs) _
                                              Handles CboSQLDB.SelectedIndexChanged
        Dim strDB As String = Me.CboSQLDB.Text
        Select Case Me.cboSQLAuth.SelectedIndex
            Case 0
                strSQLCnn = objSQL.TrustedConnection(txtServer.Text, strDB)
            Case 1
                strSQLCnn =
                    objSQL.StandardSecurity(txtServer.Text,
                                            strDB, txtSQLLogin.Text,
                                            txtSQL_PW.Text)
        End Select
        Me.CboSQLTbls.DataSource = objSQL.MSSQL_DB_Tbls(strSQLCnn)
        Me.BtnSQLTabPg.Enabled = True
    End Sub

    Private Sub BtnSQLTabPg_Click(sender As Object,
                                  e As EventArgs) _
                                  Handles BtnSQLTabPg.Click
        'Create an instance of BindingNavigator control
        Dim ctrlNav As New BindingNavigator(True) _
                With {.Dock = DockStyle.Top},
            bindSrc As New BindingSource,
            tp As New TabPage,
            TLP As New MRM_MSSQL,
            dynTLP As New TableLayoutPanel,
            exam As MRM_Ctrls = New MRM_Ctrls
        strTbl = Me.CboSQLTbls.SelectedValue
        Cursor = Cursors.WaitCursor
        'Get table schema
        dt_Schema = objSQL.Dt_MSSQL_Tbl_Schema(strSQLCnn, strTbl)
        'Get Table Relationships
        dt_Rel = objSQL.Dt_MSSQL_Tbl_Relationships(strSQLCnn, strTbl)
        'Get table records
        dt_Records = objSQL.Dt_MSSQL_Retrive_Data(strSQLCnn, strTbl)
        bindSrc.DataSource = dt_Records
        'Place BindingNavigator control at top of the TabPage 
        ctrlNav.Dock = DockStyle.Top
        'Adding Save Button to BindingNavigator 
        ctrlNav = CtrlNavigator(ctrlNav, strAppPath & "\Save.bmp")
        'Create TableLayoutPanel for TabPage
        dynTLP = exam.Create_MSSQL_DE_Form(ctrlNav, bindSrc,
                                     strSQLCnn, Me.CboSQLTbls.Text)
        Me.dsTbls.Tables.Add(exam.dtTbls)
        'Create TabPage
        Try
            With tp
                .AutoScroll = True
                .Name = Me.CboSQLTbls.Text
                .Text = Me.CboSQLTbls.Text
                .ContextMenuStrip = cntxSQL
                .Name = Me.CboSQLTbls.Text
                .Text = Me.CboSQLTbls.Text
                .Controls.Add(ctrlNav)
                .Controls.Add(dynTLP)
            End With
            Me.tbctrlTblPages.TabPages.Add(tp)
        Catch ex As Exception
            'Log error 
            Dim el As New Log.ErrorLogger
            el.WriteToErrorLog(ex.Message, ex.StackTrace, "Error")
        End Try
        Cursor = Cursors.Default
    End Sub

    Private Sub tabCtrl_Upper_SelectedIndexChanged(sender As Object,
                                                   e As EventArgs) _
                                                   Handles tabCtrl_Upper.SelectedIndexChanged
        Me.tabCtrl_Upper.Refresh()
        Select Case tabCtrl_Upper.SelectedIndex
            Case 0
                'Access Tab page
                Me.tbctrlTblPages.TabPages.Clear()
                If Me.cboOleDBTbls.Items.Count = 0 Then
                    strMsg = "Could not retrieve MS_Access Database tables." &
                          vbCrLf & "Try to reconnect to the database."
                    MessageBox.Show(strMsg, "Database Connection Error.",
                                    MessageBoxButtons.OK)
                End If
            Case 1
                'MSSQL Server tab page
                Me.tbctrlTblPages.TabPages.Clear()
        End Select
    End Sub

    Private Sub BtnFind_Click(sender As Object,
                              e As EventArgs) _
                              Handles BtnFind.Click
        'Purpose:   Searches the network for instances of MS SQL Server
        'Disable Me.txtServer and search for avaliable servers
        Cursor = Cursors.WaitCursor
        Me.txtServer.Enabled = False
        Me.lblMsg.ForeColor = Color.CornflowerBlue
        Me.lblMsg.Text = "Searching for all MS-SQL Server instances ..."
        Me.tabCtrl_Upper.Refresh()
        'List all SQL Servers on the Network
        If objSQL.Lst_SQLServers.Count > 0 Then
            Me.LstServers.DataSource = objSQL.Lst_SQLServers
            Me.txtSQLLogin.Enabled = False
            Me.txtSQL_PW.Enabled = False
            Me.BtnSQLTabPg.Enabled = False
        Else
            Me.lblMsg.ForeColor = Color.Red
            'Add my MRM-WINX-002 as default SQl Server Name
            Me.LstServers.Items.Add("MRM-WINX-002")
            strMsg = "No MS-SQL Servers detected on network." &
                  vbCrLf & "You must manually enter in MS-SQL Server's name." &
                  vbCrLf & "NOTE: The server name in the list box is a default."
            lblMsg.Text = strMsg
            With Me.txtServer
                .Enabled = True
                .Select()
            End With
        End If
        Cursor = Cursors.Default
    End Sub

    Private Sub LstServers_SelectedValueChanged(sender As Object,
                                                e As EventArgs) _
                                                Handles LstServers.SelectedValueChanged
        Me.txtServer.Text =
           Me.LstServers.Items(LstServers.SelectedValue)
    End Sub
End Class