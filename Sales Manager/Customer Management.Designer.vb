<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCustomer_Management
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomer_Management))
        Me.grbQuanLyKhachHang = New System.Windows.Forms.GroupBox()
        Me.cbbGioiTinh = New System.Windows.Forms.ComboBox()
        Me.btnclose = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnedit = New System.Windows.Forms.Button()
        Me.btnadd = New System.Windows.Forms.Button()
        Me.dgvDataGridView = New System.Windows.Forms.DataGridView()
        Me.cbbTonGiao = New System.Windows.Forms.ComboBox()
        Me.txtSoDienThoai = New System.Windows.Forms.TextBox()
        Me.txtDiaChi = New System.Windows.Forms.TextBox()
        Me.txtTenKH = New System.Windows.Forms.TextBox()
        Me.txtMaKH = New System.Windows.Forms.TextBox()
        Me.lblGioiTinh = New System.Windows.Forms.Label()
        Me.lblTonGiao = New System.Windows.Forms.Label()
        Me.lblSoDienThoai = New System.Windows.Forms.Label()
        Me.LblDiaChi = New System.Windows.Forms.Label()
        Me.lblTenKhachHang = New System.Windows.Forms.Label()
        Me.lblMaKhachHang = New System.Windows.Forms.Label()
        Me.grbQuanLyKhachHang.SuspendLayout()
        CType(Me.dgvDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grbQuanLyKhachHang
        '
        Me.grbQuanLyKhachHang.Controls.Add(Me.cbbGioiTinh)
        Me.grbQuanLyKhachHang.Controls.Add(Me.btnclose)
        Me.grbQuanLyKhachHang.Controls.Add(Me.btnDelete)
        Me.grbQuanLyKhachHang.Controls.Add(Me.btnedit)
        Me.grbQuanLyKhachHang.Controls.Add(Me.btnadd)
        Me.grbQuanLyKhachHang.Controls.Add(Me.dgvDataGridView)
        Me.grbQuanLyKhachHang.Controls.Add(Me.cbbTonGiao)
        Me.grbQuanLyKhachHang.Controls.Add(Me.txtSoDienThoai)
        Me.grbQuanLyKhachHang.Controls.Add(Me.txtDiaChi)
        Me.grbQuanLyKhachHang.Controls.Add(Me.txtTenKH)
        Me.grbQuanLyKhachHang.Controls.Add(Me.txtMaKH)
        Me.grbQuanLyKhachHang.Controls.Add(Me.lblGioiTinh)
        Me.grbQuanLyKhachHang.Controls.Add(Me.lblTonGiao)
        Me.grbQuanLyKhachHang.Controls.Add(Me.lblSoDienThoai)
        Me.grbQuanLyKhachHang.Controls.Add(Me.LblDiaChi)
        Me.grbQuanLyKhachHang.Controls.Add(Me.lblTenKhachHang)
        Me.grbQuanLyKhachHang.Controls.Add(Me.lblMaKhachHang)
        Me.grbQuanLyKhachHang.Location = New System.Drawing.Point(0, 0)
        Me.grbQuanLyKhachHang.Name = "grbQuanLyKhachHang"
        Me.grbQuanLyKhachHang.Size = New System.Drawing.Size(685, 502)
        Me.grbQuanLyKhachHang.TabIndex = 1
        Me.grbQuanLyKhachHang.TabStop = False
        Me.grbQuanLyKhachHang.Text = "Quản Lý Khách Hàng"
        '
        'cbbGioiTinh
        '
        Me.cbbGioiTinh.FormattingEnabled = True
        Me.cbbGioiTinh.Items.AddRange(New Object() {"Nam", "Nữ", "3D"})
        Me.cbbGioiTinh.Location = New System.Drawing.Point(361, 117)
        Me.cbbGioiTinh.Name = "cbbGioiTinh"
        Me.cbbGioiTinh.Size = New System.Drawing.Size(75, 27)
        Me.cbbGioiTinh.TabIndex = 5
        '
        'btnclose
        '
        Me.btnclose.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnclose.Location = New System.Drawing.Point(575, 205)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.Size = New System.Drawing.Size(75, 27)
        Me.btnclose.TabIndex = 10
        Me.btnclose.Text = "Đóng"
        Me.btnclose.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(494, 205)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 27)
        Me.btnDelete.TabIndex = 9
        Me.btnDelete.Text = "Xóa"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnedit
        '
        Me.btnedit.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnedit.Location = New System.Drawing.Point(413, 205)
        Me.btnedit.Name = "btnedit"
        Me.btnedit.Size = New System.Drawing.Size(75, 27)
        Me.btnedit.TabIndex = 8
        Me.btnedit.Text = "Sửa"
        Me.btnedit.UseVisualStyleBackColor = True
        '
        'btnadd
        '
        Me.btnadd.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnadd.Location = New System.Drawing.Point(332, 205)
        Me.btnadd.Name = "btnadd"
        Me.btnadd.Size = New System.Drawing.Size(75, 27)
        Me.btnadd.TabIndex = 7
        Me.btnadd.Text = "Thêm"
        Me.btnadd.UseVisualStyleBackColor = True
        '
        'dgvDataGridView
        '
        Me.dgvDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvDataGridView.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.dgvDataGridView.Location = New System.Drawing.Point(3, 248)
        Me.dgvDataGridView.Name = "dgvDataGridView"
        Me.dgvDataGridView.Size = New System.Drawing.Size(679, 251)
        Me.dgvDataGridView.TabIndex = 12
        '
        'cbbTonGiao
        '
        Me.cbbTonGiao.FormattingEnabled = True
        Me.cbbTonGiao.Items.AddRange(New Object() {"Thiên Chúa Giáo", "Phật Giáo", "Hồi Giáo", "Cao Đài", "Tin Lành"})
        Me.cbbTonGiao.Location = New System.Drawing.Point(361, 168)
        Me.cbbTonGiao.Name = "cbbTonGiao"
        Me.cbbTonGiao.Size = New System.Drawing.Size(208, 27)
        Me.cbbTonGiao.TabIndex = 6
        '
        'txtSoDienThoai
        '
        Me.txtSoDienThoai.Location = New System.Drawing.Point(361, 63)
        Me.txtSoDienThoai.MaxLength = 50
        Me.txtSoDienThoai.Name = "txtSoDienThoai"
        Me.txtSoDienThoai.Size = New System.Drawing.Size(289, 27)
        Me.txtSoDienThoai.TabIndex = 4
        '
        'txtDiaChi
        '
        Me.txtDiaChi.Location = New System.Drawing.Point(32, 168)
        Me.txtDiaChi.MaxLength = 50
        Me.txtDiaChi.Name = "txtDiaChi"
        Me.txtDiaChi.Size = New System.Drawing.Size(289, 27)
        Me.txtDiaChi.TabIndex = 3
        '
        'txtTenKH
        '
        Me.txtTenKH.Location = New System.Drawing.Point(32, 117)
        Me.txtTenKH.MaxLength = 30
        Me.txtTenKH.Name = "txtTenKH"
        Me.txtTenKH.Size = New System.Drawing.Size(289, 27)
        Me.txtTenKH.TabIndex = 2
        '
        'txtMaKH
        '
        Me.txtMaKH.Location = New System.Drawing.Point(32, 63)
        Me.txtMaKH.MaxLength = 20
        Me.txtMaKH.Name = "txtMaKH"
        Me.txtMaKH.Size = New System.Drawing.Size(289, 27)
        Me.txtMaKH.TabIndex = 1
        '
        'lblGioiTinh
        '
        Me.lblGioiTinh.AutoSize = True
        Me.lblGioiTinh.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGioiTinh.Location = New System.Drawing.Point(357, 95)
        Me.lblGioiTinh.Name = "lblGioiTinh"
        Me.lblGioiTinh.Size = New System.Drawing.Size(83, 20)
        Me.lblGioiTinh.TabIndex = 11
        Me.lblGioiTinh.Text = "Giới Tính:"
        '
        'lblTonGiao
        '
        Me.lblTonGiao.AutoSize = True
        Me.lblTonGiao.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTonGiao.Location = New System.Drawing.Point(357, 146)
        Me.lblTonGiao.Name = "lblTonGiao"
        Me.lblTonGiao.Size = New System.Drawing.Size(80, 20)
        Me.lblTonGiao.TabIndex = 10
        Me.lblTonGiao.Text = "Tôn Giáo:"
        '
        'lblSoDienThoai
        '
        Me.lblSoDienThoai.AutoSize = True
        Me.lblSoDienThoai.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSoDienThoai.Location = New System.Drawing.Point(357, 41)
        Me.lblSoDienThoai.Name = "lblSoDienThoai"
        Me.lblSoDienThoai.Size = New System.Drawing.Size(115, 20)
        Me.lblSoDienThoai.TabIndex = 9
        Me.lblSoDienThoai.Text = "Số Điện Thoại:"
        '
        'LblDiaChi
        '
        Me.LblDiaChi.AutoSize = True
        Me.LblDiaChi.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblDiaChi.Location = New System.Drawing.Point(28, 146)
        Me.LblDiaChi.Name = "LblDiaChi"
        Me.LblDiaChi.Size = New System.Drawing.Size(70, 20)
        Me.LblDiaChi.TabIndex = 8
        Me.LblDiaChi.Text = "Địa Chỉ:"
        '
        'lblTenKhachHang
        '
        Me.lblTenKhachHang.AutoSize = True
        Me.lblTenKhachHang.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTenKhachHang.Location = New System.Drawing.Point(28, 95)
        Me.lblTenKhachHang.Name = "lblTenKhachHang"
        Me.lblTenKhachHang.Size = New System.Drawing.Size(135, 20)
        Me.lblTenKhachHang.TabIndex = 7
        Me.lblTenKhachHang.Text = "Tên Khách Hàng:"
        '
        'lblMaKhachHang
        '
        Me.lblMaKhachHang.AutoSize = True
        Me.lblMaKhachHang.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaKhachHang.Location = New System.Drawing.Point(28, 41)
        Me.lblMaKhachHang.Name = "lblMaKhachHang"
        Me.lblMaKhachHang.Size = New System.Drawing.Size(132, 20)
        Me.lblMaKhachHang.TabIndex = 6
        Me.lblMaKhachHang.Text = "Mã Khách Hàng:"
        '
        'frmCustomer_Management
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(684, 501)
        Me.Controls.Add(Me.grbQuanLyKhachHang)
        Me.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmCustomer_Management"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Customer Management"
        Me.grbQuanLyKhachHang.ResumeLayout(False)
        Me.grbQuanLyKhachHang.PerformLayout()
        CType(Me.dgvDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grbQuanLyKhachHang As System.Windows.Forms.GroupBox
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnedit As System.Windows.Forms.Button
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents dgvDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents cbbTonGiao As System.Windows.Forms.ComboBox
    Friend WithEvents txtSoDienThoai As System.Windows.Forms.TextBox
    Friend WithEvents txtDiaChi As System.Windows.Forms.TextBox
    Friend WithEvents txtTenKH As System.Windows.Forms.TextBox
    Friend WithEvents txtMaKH As System.Windows.Forms.TextBox
    Friend WithEvents lblGioiTinh As System.Windows.Forms.Label
    Friend WithEvents lblTonGiao As System.Windows.Forms.Label
    Friend WithEvents lblSoDienThoai As System.Windows.Forms.Label
    Friend WithEvents LblDiaChi As System.Windows.Forms.Label
    Friend WithEvents lblTenKhachHang As System.Windows.Forms.Label
    Friend WithEvents lblMaKhachHang As System.Windows.Forms.Label
    Friend WithEvents cbbGioiTinh As System.Windows.Forms.ComboBox
End Class
