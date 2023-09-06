<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMain))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.tabCtrl_Upper = New System.Windows.Forms.TabControl()
        Me.MS_Access = New System.Windows.Forms.TabPage()
        Me.BtnFrmTbl = New System.Windows.Forms.Button()
        Me.BtnAccessDB = New System.Windows.Forms.Button()
        Me.lblTbls = New System.Windows.Forms.Label()
        Me.cboOleDBTbls = New System.Windows.Forms.ComboBox()
        Me.txtDB = New System.Windows.Forms.TextBox()
        Me.SQL_Server = New System.Windows.Forms.TabPage()
        Me.BtnFind = New System.Windows.Forms.Button()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.lblDetSvrs = New System.Windows.Forms.Label()
        Me.LstServers = New System.Windows.Forms.ListBox()
        Me.lblMsg = New System.Windows.Forms.Label()
        Me.BtnSQLTabPg = New System.Windows.Forms.Button()
        Me.BtnCnnSQL = New System.Windows.Forms.Button()
        Me.CboSQLTbls = New System.Windows.Forms.ComboBox()
        Me.lblSQLTbls = New System.Windows.Forms.Label()
        Me.CboSQLDB = New System.Windows.Forms.ComboBox()
        Me.lblSQLDB = New System.Windows.Forms.Label()
        Me.txtSQL_PW = New System.Windows.Forms.TextBox()
        Me.txtSQLLogin = New System.Windows.Forms.TextBox()
        Me.lblSQLPW = New System.Windows.Forms.Label()
        Me.lblSQLLogin = New System.Windows.Forms.Label()
        Me.cboSQLAuth = New System.Windows.Forms.ComboBox()
        Me.lblAuthentication = New System.Windows.Forms.Label()
        Me.lblServer = New System.Windows.Forms.Label()
        Me.tbctrlTblPages = New System.Windows.Forms.TabControl()
        Me.cntxSQL = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.RemoveAllSQLToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RemoveSQLToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.cntxOLEDB = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.AccessRemoveAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AccessRemoveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ofdAccess = New System.Windows.Forms.OpenFileDialog()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.tabCtrl_Upper.SuspendLayout()
        Me.MS_Access.SuspendLayout()
        Me.SQL_Server.SuspendLayout()
        Me.cntxSQL.SuspendLayout()
        Me.cntxOLEDB.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.BackColor = System.Drawing.SystemColors.Control
        Me.SplitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.tabCtrl_Upper)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.tbctrlTblPages)
        Me.SplitContainer1.Size = New System.Drawing.Size(1157, 625)
        Me.SplitContainer1.SplitterDistance = 244
        Me.SplitContainer1.TabIndex = 0
        '
        'tabCtrl_Upper
        '
        Me.tabCtrl_Upper.Appearance = System.Windows.Forms.TabAppearance.Buttons
        Me.tabCtrl_Upper.Controls.Add(Me.MS_Access)
        Me.tabCtrl_Upper.Controls.Add(Me.SQL_Server)
        Me.tabCtrl_Upper.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabCtrl_Upper.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabCtrl_Upper.Location = New System.Drawing.Point(0, 0)
        Me.tabCtrl_Upper.Name = "tabCtrl_Upper"
        Me.tabCtrl_Upper.SelectedIndex = 0
        Me.tabCtrl_Upper.Size = New System.Drawing.Size(1153, 240)
        Me.tabCtrl_Upper.TabIndex = 0
        Me.tabCtrl_Upper.Tag = ""
        '
        'MS_Access
        '
        Me.MS_Access.BackColor = System.Drawing.SystemColors.Control
        Me.MS_Access.BackgroundImage = CType(resources.GetObject("MS_Access.BackgroundImage"), System.Drawing.Image)
        Me.MS_Access.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.MS_Access.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.MS_Access.Controls.Add(Me.BtnFrmTbl)
        Me.MS_Access.Controls.Add(Me.BtnAccessDB)
        Me.MS_Access.Controls.Add(Me.lblTbls)
        Me.MS_Access.Controls.Add(Me.cboOleDBTbls)
        Me.MS_Access.Controls.Add(Me.txtDB)
        Me.MS_Access.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MS_Access.ForeColor = System.Drawing.Color.Black
        Me.MS_Access.Location = New System.Drawing.Point(4, 31)
        Me.MS_Access.Name = "MS_Access"
        Me.MS_Access.Padding = New System.Windows.Forms.Padding(3)
        Me.MS_Access.Size = New System.Drawing.Size(1145, 205)
        Me.MS_Access.TabIndex = 0
        Me.MS_Access.Tag = "1"
        Me.MS_Access.Text = "MS-Access"
        '
        'BtnFrmTbl
        '
        Me.BtnFrmTbl.ForeColor = System.Drawing.Color.Black
        Me.BtnFrmTbl.Location = New System.Drawing.Point(219, 112)
        Me.BtnFrmTbl.Name = "BtnFrmTbl"
        Me.BtnFrmTbl.Size = New System.Drawing.Size(110, 27)
        Me.BtnFrmTbl.TabIndex = 31
        Me.BtnFrmTbl.Text = "&Access Table"
        Me.BtnFrmTbl.UseVisualStyleBackColor = True
        '
        'BtnAccessDB
        '
        Me.BtnAccessDB.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.BtnAccessDB.Location = New System.Drawing.Point(19, 78)
        Me.BtnAccessDB.Name = "BtnAccessDB"
        Me.BtnAccessDB.Size = New System.Drawing.Size(129, 27)
        Me.BtnAccessDB.TabIndex = 30
        Me.BtnAccessDB.Text = "&Access Database"
        Me.BtnAccessDB.UseVisualStyleBackColor = True
        '
        'lblTbls
        '
        Me.lblTbls.AutoSize = True
        Me.lblTbls.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTbls.ForeColor = System.Drawing.Color.Black
        Me.lblTbls.Location = New System.Drawing.Point(160, 83)
        Me.lblTbls.Name = "lblTbls"
        Me.lblTbls.Size = New System.Drawing.Size(53, 18)
        Me.lblTbls.TabIndex = 29
        Me.lblTbls.Text = "Tables:"
        '
        'cboOleDBTbls
        '
        Me.cboOleDBTbls.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboOleDBTbls.FormattingEnabled = True
        Me.cboOleDBTbls.Location = New System.Drawing.Point(219, 80)
        Me.cboOleDBTbls.Name = "cboOleDBTbls"
        Me.cboOleDBTbls.Size = New System.Drawing.Size(163, 26)
        Me.cboOleDBTbls.TabIndex = 28
        '
        'txtDB
        '
        Me.txtDB.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDB.Location = New System.Drawing.Point(19, 27)
        Me.txtDB.Multiline = True
        Me.txtDB.Name = "txtDB"
        Me.txtDB.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDB.Size = New System.Drawing.Size(363, 45)
        Me.txtDB.TabIndex = 25
        '
        'SQL_Server
        '
        Me.SQL_Server.BackgroundImage = CType(resources.GetObject("SQL_Server.BackgroundImage"), System.Drawing.Image)
        Me.SQL_Server.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.SQL_Server.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.SQL_Server.Controls.Add(Me.BtnFind)
        Me.SQL_Server.Controls.Add(Me.txtServer)
        Me.SQL_Server.Controls.Add(Me.lblDetSvrs)
        Me.SQL_Server.Controls.Add(Me.LstServers)
        Me.SQL_Server.Controls.Add(Me.lblMsg)
        Me.SQL_Server.Controls.Add(Me.BtnSQLTabPg)
        Me.SQL_Server.Controls.Add(Me.BtnCnnSQL)
        Me.SQL_Server.Controls.Add(Me.CboSQLTbls)
        Me.SQL_Server.Controls.Add(Me.lblSQLTbls)
        Me.SQL_Server.Controls.Add(Me.CboSQLDB)
        Me.SQL_Server.Controls.Add(Me.lblSQLDB)
        Me.SQL_Server.Controls.Add(Me.txtSQL_PW)
        Me.SQL_Server.Controls.Add(Me.txtSQLLogin)
        Me.SQL_Server.Controls.Add(Me.lblSQLPW)
        Me.SQL_Server.Controls.Add(Me.lblSQLLogin)
        Me.SQL_Server.Controls.Add(Me.cboSQLAuth)
        Me.SQL_Server.Controls.Add(Me.lblAuthentication)
        Me.SQL_Server.Controls.Add(Me.lblServer)
        Me.SQL_Server.Location = New System.Drawing.Point(4, 31)
        Me.SQL_Server.Name = "SQL_Server"
        Me.SQL_Server.Padding = New System.Windows.Forms.Padding(3)
        Me.SQL_Server.Size = New System.Drawing.Size(1145, 205)
        Me.SQL_Server.TabIndex = 1
        Me.SQL_Server.Tag = "2"
        Me.SQL_Server.Text = "MS-SQL Server"
        Me.SQL_Server.UseVisualStyleBackColor = True
        '
        'BtnFind
        '
        Me.BtnFind.Location = New System.Drawing.Point(10, 167)
        Me.BtnFind.Name = "BtnFind"
        Me.BtnFind.Size = New System.Drawing.Size(116, 27)
        Me.BtnFind.TabIndex = 39
        Me.BtnFind.Text = "Network SVRs"
        Me.BtnFind.UseVisualStyleBackColor = True
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(155, 25)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(141, 26)
        Me.txtServer.TabIndex = 38
        '
        'lblDetSvrs
        '
        Me.lblDetSvrs.AutoSize = True
        Me.lblDetSvrs.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDetSvrs.Location = New System.Drawing.Point(7, 6)
        Me.lblDetSvrs.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDetSvrs.Name = "lblDetSvrs"
        Me.lblDetSvrs.Size = New System.Drawing.Size(86, 18)
        Me.lblDetSvrs.TabIndex = 37
        Me.lblDetSvrs.Text = "SQL Servers"
        '
        'LstServers
        '
        Me.LstServers.FormattingEnabled = True
        Me.LstServers.ItemHeight = 19
        Me.LstServers.Location = New System.Drawing.Point(10, 27)
        Me.LstServers.Name = "LstServers"
        Me.LstServers.Size = New System.Drawing.Size(116, 137)
        Me.LstServers.TabIndex = 36
        '
        'lblMsg
        '
        Me.lblMsg.AutoSize = True
        Me.lblMsg.Font = New System.Drawing.Font("Comic Sans MS", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMsg.Location = New System.Drawing.Point(151, 109)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(78, 23)
        Me.lblMsg.TabIndex = 35
        Me.lblMsg.Text = "Messages"
        '
        'BtnSQLTabPg
        '
        Me.BtnSQLTabPg.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnSQLTabPg.Location = New System.Drawing.Point(681, 75)
        Me.BtnSQLTabPg.Name = "BtnSQLTabPg"
        Me.BtnSQLTabPg.Size = New System.Drawing.Size(125, 32)
        Me.BtnSQLTabPg.TabIndex = 33
        Me.BtnSQLTabPg.Text = "Create TabPage"
        Me.BtnSQLTabPg.UseVisualStyleBackColor = True
        '
        'BtnCnnSQL
        '
        Me.BtnCnnSQL.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.BtnCnnSQL.Location = New System.Drawing.Point(551, 75)
        Me.BtnCnnSQL.Name = "BtnCnnSQL"
        Me.BtnCnnSQL.Size = New System.Drawing.Size(124, 32)
        Me.BtnCnnSQL.TabIndex = 32
        Me.BtnCnnSQL.Text = "Connect"
        Me.BtnCnnSQL.UseVisualStyleBackColor = True
        '
        'CboSQLTbls
        '
        Me.CboSQLTbls.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboSQLTbls.FormattingEnabled = True
        Me.CboSQLTbls.Items.AddRange(New Object() {"<default>"})
        Me.CboSQLTbls.Location = New System.Drawing.Point(353, 79)
        Me.CboSQLTbls.Name = "CboSQLTbls"
        Me.CboSQLTbls.Size = New System.Drawing.Size(192, 27)
        Me.CboSQLTbls.TabIndex = 31
        '
        'lblSQLTbls
        '
        Me.lblSQLTbls.AutoSize = True
        Me.lblSQLTbls.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSQLTbls.Location = New System.Drawing.Point(350, 58)
        Me.lblSQLTbls.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSQLTbls.Name = "lblSQLTbls"
        Me.lblSQLTbls.Size = New System.Drawing.Size(53, 18)
        Me.lblSQLTbls.TabIndex = 30
        Me.lblSQLTbls.Text = "Tables:"
        '
        'CboSQLDB
        '
        Me.CboSQLDB.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CboSQLDB.FormattingEnabled = True
        Me.CboSQLDB.Items.AddRange(New Object() {"<default>"})
        Me.CboSQLDB.Location = New System.Drawing.Point(155, 79)
        Me.CboSQLDB.Name = "CboSQLDB"
        Me.CboSQLDB.Size = New System.Drawing.Size(192, 27)
        Me.CboSQLDB.TabIndex = 29
        '
        'lblSQLDB
        '
        Me.lblSQLDB.AutoSize = True
        Me.lblSQLDB.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSQLDB.Location = New System.Drawing.Point(152, 58)
        Me.lblSQLDB.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSQLDB.Name = "lblSQLDB"
        Me.lblSQLDB.Size = New System.Drawing.Size(68, 18)
        Me.lblSQLDB.TabIndex = 28
        Me.lblSQLDB.Text = "Database:"
        '
        'txtSQL_PW
        '
        Me.txtSQL_PW.Location = New System.Drawing.Point(702, 27)
        Me.txtSQL_PW.Name = "txtSQL_PW"
        Me.txtSQL_PW.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSQL_PW.Size = New System.Drawing.Size(125, 26)
        Me.txtSQL_PW.TabIndex = 27
        '
        'txtSQLLogin
        '
        Me.txtSQLLogin.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSQLLogin.Location = New System.Drawing.Point(555, 27)
        Me.txtSQLLogin.Name = "txtSQLLogin"
        Me.txtSQLLogin.Size = New System.Drawing.Size(141, 26)
        Me.txtSQLLogin.TabIndex = 26
        '
        'lblSQLPW
        '
        Me.lblSQLPW.AutoSize = True
        Me.lblSQLPW.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSQLPW.Location = New System.Drawing.Point(699, 6)
        Me.lblSQLPW.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSQLPW.Name = "lblSQLPW"
        Me.lblSQLPW.Size = New System.Drawing.Size(69, 18)
        Me.lblSQLPW.TabIndex = 25
        Me.lblSQLPW.Text = "Password:"
        '
        'lblSQLLogin
        '
        Me.lblSQLLogin.AutoSize = True
        Me.lblSQLLogin.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSQLLogin.Location = New System.Drawing.Point(552, 5)
        Me.lblSQLLogin.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSQLLogin.Name = "lblSQLLogin"
        Me.lblSQLLogin.Size = New System.Drawing.Size(44, 18)
        Me.lblSQLLogin.TabIndex = 24
        Me.lblSQLLogin.Text = "Login:"
        '
        'cboSQLAuth
        '
        Me.cboSQLAuth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSQLAuth.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSQLAuth.FormattingEnabled = True
        Me.cboSQLAuth.Items.AddRange(New Object() {"Windows Authentication", "SQL Server Authentication"})
        Me.cboSQLAuth.Location = New System.Drawing.Point(302, 26)
        Me.cboSQLAuth.Name = "cboSQLAuth"
        Me.cboSQLAuth.Size = New System.Drawing.Size(247, 26)
        Me.cboSQLAuth.TabIndex = 23
        '
        'lblAuthentication
        '
        Me.lblAuthentication.AutoSize = True
        Me.lblAuthentication.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAuthentication.Location = New System.Drawing.Point(299, 5)
        Me.lblAuthentication.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblAuthentication.Name = "lblAuthentication"
        Me.lblAuthentication.Size = New System.Drawing.Size(100, 18)
        Me.lblAuthentication.TabIndex = 22
        Me.lblAuthentication.Text = "Authentication:"
        '
        'lblServer
        '
        Me.lblServer.AutoSize = True
        Me.lblServer.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblServer.Location = New System.Drawing.Point(152, 6)
        Me.lblServer.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblServer.Name = "lblServer"
        Me.lblServer.Size = New System.Drawing.Size(90, 18)
        Me.lblServer.TabIndex = 20
        Me.lblServer.Text = "Server Name:"
        '
        'tbctrlTblPages
        '
        Me.tbctrlTblPages.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbctrlTblPages.Font = New System.Drawing.Font("Comic Sans MS", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbctrlTblPages.Location = New System.Drawing.Point(0, 0)
        Me.tbctrlTblPages.Name = "tbctrlTblPages"
        Me.tbctrlTblPages.SelectedIndex = 0
        Me.tbctrlTblPages.Size = New System.Drawing.Size(1153, 373)
        Me.tbctrlTblPages.TabIndex = 0
        '
        'cntxSQL
        '
        Me.cntxSQL.BackColor = System.Drawing.SystemColors.Control
        Me.cntxSQL.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.cntxSQL.Font = New System.Drawing.Font("Comic Sans MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cntxSQL.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RemoveAllSQLToolStripMenuItem, Me.RemoveSQLToolStripMenuItem})
        Me.cntxSQL.Name = "cntxSQL"
        Me.cntxSQL.Size = New System.Drawing.Size(136, 48)
        '
        'RemoveAllSQLToolStripMenuItem
        '
        Me.RemoveAllSQLToolStripMenuItem.Name = "RemoveAllSQLToolStripMenuItem"
        Me.RemoveAllSQLToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.RemoveAllSQLToolStripMenuItem.Text = "Remove All"
        '
        'RemoveSQLToolStripMenuItem
        '
        Me.RemoveSQLToolStripMenuItem.Name = "RemoveSQLToolStripMenuItem"
        Me.RemoveSQLToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.RemoveSQLToolStripMenuItem.Text = "Remove"
        '
        'cntxOLEDB
        '
        Me.cntxOLEDB.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.cntxOLEDB.Font = New System.Drawing.Font("Comic Sans MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cntxOLEDB.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AccessRemoveAllToolStripMenuItem, Me.AccessRemoveToolStripMenuItem})
        Me.cntxOLEDB.Name = "cntxMenu"
        Me.cntxOLEDB.Size = New System.Drawing.Size(136, 48)
        '
        'AccessRemoveAllToolStripMenuItem
        '
        Me.AccessRemoveAllToolStripMenuItem.Name = "AccessRemoveAllToolStripMenuItem"
        Me.AccessRemoveAllToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.AccessRemoveAllToolStripMenuItem.Text = "Remove All"
        '
        'AccessRemoveToolStripMenuItem
        '
        Me.AccessRemoveToolStripMenuItem.Name = "AccessRemoveToolStripMenuItem"
        Me.AccessRemoveToolStripMenuItem.Size = New System.Drawing.Size(135, 22)
        Me.AccessRemoveToolStripMenuItem.Text = "Remove"
        '
        'ofdAccess
        '
        Me.ofdAccess.FileName = "OpenFileDialog1"
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1157, 625)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "FrmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Dynamic Data Entry "
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.tabCtrl_Upper.ResumeLayout(False)
        Me.MS_Access.ResumeLayout(False)
        Me.MS_Access.PerformLayout()
        Me.SQL_Server.ResumeLayout(False)
        Me.SQL_Server.PerformLayout()
        Me.cntxSQL.ResumeLayout(False)
        Me.cntxOLEDB.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents tabCtrl_Upper As System.Windows.Forms.TabControl
    Friend WithEvents MS_Access As System.Windows.Forms.TabPage
    Friend WithEvents SQL_Server As System.Windows.Forms.TabPage
    Friend WithEvents lblTbls As System.Windows.Forms.Label
    Friend WithEvents cboOleDBTbls As System.Windows.Forms.ComboBox
    Friend WithEvents txtDB As System.Windows.Forms.TextBox
    Friend WithEvents BtnAccessDB As System.Windows.Forms.Button
    Friend WithEvents tbctrlTblPages As System.Windows.Forms.TabControl
    Friend WithEvents BtnFrmTbl As System.Windows.Forms.Button
    Friend WithEvents cntxSQL As ContextMenuStrip
    Friend WithEvents RemoveAllSQLToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RemoveSQLToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents cntxOLEDB As ContextMenuStrip
    Friend WithEvents AccessRemoveAllToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AccessRemoveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents BtnSQLTabPg As Button
    Friend WithEvents BtnCnnSQL As Button
    Friend WithEvents CboSQLTbls As ComboBox
    Friend WithEvents lblSQLTbls As Label
    Friend WithEvents CboSQLDB As ComboBox
    Friend WithEvents lblSQLDB As Label
    Friend WithEvents txtSQL_PW As TextBox
    Friend WithEvents txtSQLLogin As TextBox
    Friend WithEvents lblSQLPW As Label
    Friend WithEvents lblSQLLogin As Label
    Friend WithEvents cboSQLAuth As ComboBox
    Friend WithEvents lblAuthentication As Label
    Friend WithEvents lblServer As Label
    Friend WithEvents lblMsg As Label
    Friend WithEvents ofdAccess As OpenFileDialog
    Friend WithEvents txtServer As TextBox
    Friend WithEvents lblDetSvrs As Label
    Friend WithEvents LstServers As ListBox
    Friend WithEvents BtnFind As Button
End Class
