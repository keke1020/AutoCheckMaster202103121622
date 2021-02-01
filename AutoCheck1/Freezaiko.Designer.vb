<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Freezaiko
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Freezaiko))
        Me.DGV1 = New System.Windows.Forms.DataGridView()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.処理ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.エクスポートToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.DGV2 = New System.Windows.Forms.DataGridView()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.行を削除ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.列を削除ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.TabControl2 = New System.Windows.Forms.TabControl()
        Me.TabPage4 = New System.Windows.Forms.TabPage()
        Me.DGV4 = New System.Windows.Forms.DataGridView()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TabPage5 = New System.Windows.Forms.TabPage()
        Me.DGV6 = New System.Windows.Forms.DataGridView()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.DGV5 = New System.Windows.Forms.DataGridView()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DGV3 = New System.Windows.Forms.DataGridView()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripDropDownButton6 = New System.Windows.Forms.ToolStripDropDownButton()
        Me.フォームのリセットToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.DGV2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        CType(Me.DGV4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage5.SuspendLayout()
        CType(Me.DGV6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGV5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DGV3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DGV1
        '
        Me.DGV1.AllowUserToAddRows = False
        Me.DGV1.AllowUserToOrderColumns = True
        Me.DGV1.BackgroundColor = System.Drawing.Color.White
        Me.DGV1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV1.Location = New System.Drawing.Point(2, 173)
        Me.DGV1.Name = "DGV1"
        Me.DGV1.RowTemplate.Height = 21
        Me.DGV1.Size = New System.Drawing.Size(454, 129)
        Me.DGV1.TabIndex = 0
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1, Me.処理ToolStripMenuItem, Me.エクスポートToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(807, 26)
        Me.MenuStrip1.TabIndex = 1
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(68, 22)
        Me.ToolStripMenuItem1.Text = "スタート"
        '
        '処理ToolStripMenuItem
        '
        Me.処理ToolStripMenuItem.Name = "処理ToolStripMenuItem"
        Me.処理ToolStripMenuItem.Size = New System.Drawing.Size(44, 22)
        Me.処理ToolStripMenuItem.Text = "処理"
        '
        'エクスポートToolStripMenuItem
        '
        Me.エクスポートToolStripMenuItem.Name = "エクスポートToolStripMenuItem"
        Me.エクスポートToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.エクスポートToolStripMenuItem.Text = "エクスポート"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(0, 29)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(807, 551)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.GroupBox2)
        Me.TabPage1.Controls.Add(Me.Label13)
        Me.TabPage1.Controls.Add(Me.Label12)
        Me.TabPage1.Controls.Add(Me.DGV2)
        Me.TabPage1.Controls.Add(Me.TextBox2)
        Me.TabPage1.Controls.Add(Me.Label11)
        Me.TabPage1.Controls.Add(Me.Label9)
        Me.TabPage1.Controls.Add(Me.TextBox1)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.GroupBox1)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.DGV1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(799, 525)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Ⅰ"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioButton2)
        Me.GroupBox2.Controls.Add(Me.RadioButton1)
        Me.GroupBox2.Controls.Add(Me.CheckBox1)
        Me.GroupBox2.Controls.Add(Me.CheckBox3)
        Me.GroupBox2.Controls.Add(Me.CheckBox2)
        Me.GroupBox2.Location = New System.Drawing.Point(42, 6)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(200, 98)
        Me.GroupBox2.TabIndex = 26
        Me.GroupBox2.TabStop = False
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.RadioButton2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.RadioButton2.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.RadioButton2.Location = New System.Drawing.Point(5, 77)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(166, 15)
        Me.RadioButton2.TabIndex = 27
        Me.RadioButton2.Text = "フリー在庫＞引当数*3場合除外"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.RadioButton1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.RadioButton1.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.RadioButton1.Location = New System.Drawing.Point(5, 60)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(166, 15)
        Me.RadioButton1.TabIndex = 26
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "フリー在庫＞引当数*2場合除外"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Enabled = False
        Me.CheckBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CheckBox1.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(5, 11)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(121, 15)
        Me.CheckBox1.TabIndex = 23
        Me.CheckBox1.Text = "1.引当数は0場合除外"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Checked = True
        Me.CheckBox3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CheckBox3.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.CheckBox3.Location = New System.Drawing.Point(5, 43)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(118, 15)
        Me.CheckBox3.TabIndex = 25
        Me.CheckBox3.Text = "3.フリー在庫<=引当数"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CheckBox2.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(5, 27)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(116, 15)
        Me.CheckBox2.TabIndex = 24
        Me.CheckBox2.Text = "2.フリー在庫は10以下"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("ＭＳ ゴシック", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label13.Location = New System.Drawing.Point(5, 316)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(29, 11)
        Me.Label13.TabIndex = 21
        Me.Label13.Text = "結果"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("ＭＳ ゴシック", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label12.Location = New System.Drawing.Point(5, 158)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(53, 11)
        Me.Label12.TabIndex = 20
        Me.Label12.Text = "元データ"
        '
        'DGV2
        '
        Me.DGV2.AllowUserToAddRows = False
        Me.DGV2.AllowUserToOrderColumns = True
        Me.DGV2.BackgroundColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.DGV2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV2.ContextMenuStrip = Me.ContextMenuStrip1
        Me.DGV2.Location = New System.Drawing.Point(2, 331)
        Me.DGV2.Name = "DGV2"
        Me.DGV2.RowTemplate.Height = 21
        Me.DGV2.Size = New System.Drawing.Size(454, 191)
        Me.DGV2.TabIndex = 19
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.行を削除ToolStripMenuItem1, Me.列を削除ToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(165, 48)
        '
        '行を削除ToolStripMenuItem1
        '
        Me.行を削除ToolStripMenuItem1.Name = "行を削除ToolStripMenuItem1"
        Me.行を削除ToolStripMenuItem1.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.D), System.Windows.Forms.Keys)
        Me.行を削除ToolStripMenuItem1.Size = New System.Drawing.Size(164, 22)
        Me.行を削除ToolStripMenuItem1.Text = "行を削除"
        '
        '列を削除ToolStripMenuItem
        '
        Me.列を削除ToolStripMenuItem.Name = "列を削除ToolStripMenuItem"
        Me.列を削除ToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.列を削除ToolStripMenuItem.Text = "列を削除"
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.Color.LightYellow
        Me.TextBox2.Font = New System.Drawing.Font("ＭＳ ゴシック", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(60, 133)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(367, 18)
        Me.TextBox2.TabIndex = 18
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label11.Location = New System.Drawing.Point(6, 136)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(48, 11)
        Me.Label11.TabIndex = 17
        Me.Label11.Text = "ファイル名:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Red
        Me.Label9.Location = New System.Drawing.Point(5, 114)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(16, 11)
        Me.Label9.TabIndex = 14
        Me.Label9.Text = "※"
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.LightYellow
        Me.TextBox1.Font = New System.Drawing.Font("ＭＳ ゴシック", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(98, 110)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(329, 18)
        Me.TextBox1.TabIndex = 13
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label8.Location = New System.Drawing.Point(19, 114)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(73, 11)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "ダウンロードパス:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.WebBrowser1)
        Me.GroupBox1.Location = New System.Drawing.Point(462, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(334, 516)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "カレンダー"
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WebBrowser1.Location = New System.Drawing.Point(3, 15)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(328, 498)
        Me.WebBrowser1.TabIndex = 0
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 5)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(31, 12)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "条件:"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.TabControl2)
        Me.TabPage3.Controls.Add(Me.Button3)
        Me.TabPage3.Controls.Add(Me.DGV5)
        Me.TabPage3.Controls.Add(Me.Label14)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(799, 525)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Ⅱ"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.TabPage4)
        Me.TabControl2.Controls.Add(Me.TabPage5)
        Me.TabControl2.Location = New System.Drawing.Point(3, 3)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(788, 242)
        Me.TabControl2.TabIndex = 8
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.DGV4)
        Me.TabPage4.Controls.Add(Me.Label4)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage4.Size = New System.Drawing.Size(780, 216)
        Me.TabPage4.TabIndex = 0
        Me.TabPage4.Text = "処理データ"
        Me.TabPage4.UseVisualStyleBackColor = True
        '
        'DGV4
        '
        Me.DGV4.AllowDrop = True
        Me.DGV4.AllowUserToAddRows = False
        DataGridViewCellStyle1.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.DGV4.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DGV4.BackgroundColor = System.Drawing.Color.White
        Me.DGV4.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV4.Location = New System.Drawing.Point(6, 21)
        Me.DGV4.Name = "DGV4"
        DataGridViewCellStyle2.Font = New System.Drawing.Font("ＭＳ ゴシック", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.DGV4.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.DGV4.RowTemplate.Height = 21
        Me.DGV4.Size = New System.Drawing.Size(768, 189)
        Me.DGV4.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(305, 12)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "データ元(商品コード、商品名、引当数、フリー在庫、予約時期)"
        '
        'TabPage5
        '
        Me.TabPage5.Controls.Add(Me.DGV6)
        Me.TabPage5.Controls.Add(Me.Button2)
        Me.TabPage5.Location = New System.Drawing.Point(4, 22)
        Me.TabPage5.Name = "TabPage5"
        Me.TabPage5.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage5.Size = New System.Drawing.Size(780, 216)
        Me.TabPage5.TabIndex = 1
        Me.TabPage5.Text = "発注管理"
        Me.TabPage5.UseVisualStyleBackColor = True
        '
        'DGV6
        '
        Me.DGV6.AllowUserToAddRows = False
        Me.DGV6.AllowUserToDeleteRows = False
        DataGridViewCellStyle3.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.DGV6.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.DGV6.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV6.Location = New System.Drawing.Point(3, 3)
        Me.DGV6.Name = "DGV6"
        Me.DGV6.RowTemplate.Height = 21
        Me.DGV6.Size = New System.Drawing.Size(774, 184)
        Me.DGV6.TabIndex = 7
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.PaleGreen
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Location = New System.Drawing.Point(3, 190)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(774, 23)
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "取得"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.MintCream
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Button3.Location = New System.Drawing.Point(10, 251)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(781, 23)
        Me.Button3.TabIndex = 7
        Me.Button3.Text = "処理"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'DGV5
        '
        Me.DGV5.BackgroundColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.DGV5.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV5.Location = New System.Drawing.Point(10, 292)
        Me.DGV5.Name = "DGV5"
        Me.DGV5.RowTemplate.Height = 21
        Me.DGV5.Size = New System.Drawing.Size(781, 222)
        Me.DGV5.TabIndex = 3
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(8, 277)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(336, 12)
        Me.Label14.TabIndex = 2
        Me.Label14.Text = "結果(syohin_code、yoyaku_zaiko_su、syohin_kbn、nyusyukko_riyu)"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label2)
        Me.TabPage2.Controls.Add(Me.Label7)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.Button1)
        Me.TabPage2.Controls.Add(Me.Label10)
        Me.TabPage2.Controls.Add(Me.Label3)
        Me.TabPage2.Controls.Add(Me.NumericUpDown1)
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Controls.Add(Me.DGV3)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(799, 525)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "テスト"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(193, 247)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(102, 11)
        Me.Label2.TabIndex = 31
        Me.Label2.Text = "3.フリー在庫<=引当数"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label7.Location = New System.Drawing.Point(193, 236)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(105, 11)
        Me.Label7.TabIndex = 30
        Me.Label7.Text = "2.引当数は0場合除外"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("MS UI Gothic", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label6.Location = New System.Drawing.Point(193, 224)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 11)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "1.フリー在庫は10以下"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(218, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 28
        Me.Button1.Text = "test"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label10.Font = New System.Drawing.Font("メイリオ", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label10.Location = New System.Drawing.Point(185, 7)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(22, 26)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "1"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(160, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(17, 12)
        Me.Label3.TabIndex = 26
        Me.Label3.Text = "日"
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.BackColor = System.Drawing.Color.Yellow
        Me.NumericUpDown1.Enabled = False
        Me.NumericUpDown1.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.NumericUpDown1.ForeColor = System.Drawing.Color.Red
        Me.NumericUpDown1.Location = New System.Drawing.Point(101, 15)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.NumericUpDown1.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.ReadOnly = True
        Me.NumericUpDown1.Size = New System.Drawing.Size(59, 19)
        Me.NumericUpDown1.TabIndex = 25
        Me.NumericUpDown1.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 12)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "予測(判断必要):"
        '
        'DGV3
        '
        Me.DGV3.AllowUserToAddRows = False
        Me.DGV3.AllowUserToOrderColumns = True
        Me.DGV3.BackgroundColor = System.Drawing.Color.Gray
        Me.DGV3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV3.ContextMenuStrip = Me.ContextMenuStrip1
        Me.DGV3.Location = New System.Drawing.Point(8, 40)
        Me.DGV3.Name = "DGV3"
        Me.DGV3.RowTemplate.Height = 21
        Me.DGV3.Size = New System.Drawing.Size(285, 181)
        Me.DGV3.TabIndex = 23
        '
        'StatusStrip1
        '
        Me.StatusStrip1.BackColor = System.Drawing.Color.GhostWhite
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripDropDownButton6, Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 583)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(807, 24)
        Me.StatusStrip1.TabIndex = 9
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripDropDownButton6
        '
        Me.ToolStripDropDownButton6.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ToolStripDropDownButton6.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.フォームのリセットToolStripMenuItem})
        Me.ToolStripDropDownButton6.Image = CType(resources.GetObject("ToolStripDropDownButton6.Image"), System.Drawing.Image)
        Me.ToolStripDropDownButton6.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripDropDownButton6.Name = "ToolStripDropDownButton6"
        Me.ToolStripDropDownButton6.Size = New System.Drawing.Size(69, 22)
        Me.ToolStripDropDownButton6.Text = "リセット"
        '
        'フォームのリセットToolStripMenuItem
        '
        Me.フォームのリセットToolStripMenuItem.Name = "フォームのリセットToolStripMenuItem"
        Me.フォームのリセットToolStripMenuItem.Size = New System.Drawing.Size(184, 22)
        Me.フォームのリセットToolStripMenuItem.Text = "フォームのリセット"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 19)
        '
        'Freezaiko
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(807, 607)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Freezaiko"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "予約変更"
        CType(Me.DGV1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.DGV2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.TabControl2.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.TabPage4.PerformLayout()
        CType(Me.DGV4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage5.ResumeLayout(False)
        CType(Me.DGV6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGV5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DGV3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DGV1 As DataGridView
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents Label5 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripDropDownButton6 As ToolStripDropDownButton
    Friend WithEvents フォームのリセットToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents WebBrowser1 As WebBrowser
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents 処理ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents エクスポートToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
    Friend WithEvents 行を削除ToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents 列を削除ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DGV2 As DataGridView
    Friend WithEvents Label13 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents DGV3 As DataGridView
    Friend WithEvents Label10 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents NumericUpDown1 As NumericUpDown
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents Label14 As Label
    Friend WithEvents DGV5 As DataGridView
    Friend WithEvents Label15 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents TabControl2 As TabControl
    Friend WithEvents TabPage4 As TabPage
    Friend WithEvents DGV4 As DataGridView
    Friend WithEvents Label4 As Label
    Friend WithEvents TabPage5 As TabPage
    Friend WithEvents DGV6 As DataGridView
End Class
