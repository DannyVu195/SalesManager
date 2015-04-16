<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProduct_Management
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProduct_Management))
        Me.lblMaSP = New System.Windows.Forms.Label()
        Me.lblTenSanPham = New System.Windows.Forms.Label()
        Me.lblDongia = New System.Windows.Forms.Label()
        Me.lblSoLuong = New System.Windows.Forms.Label()
        Me.lblChiTiet = New System.Windows.Forms.Label()
        Me.lblLoaispMaLoai = New System.Windows.Forms.Label()
        Me.lblNgayNhap = New System.Windows.Forms.Label()
        Me.grbThemSanPham = New System.Windows.Forms.GroupBox()
        Me.nudSoluong = New System.Windows.Forms.NumericUpDown()
        Me.dtpNgayNhapHang = New System.Windows.Forms.DateTimePicker()
        Me.dgvConnect = New System.Windows.Forms.DataGridView()
        Me.btnclose = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnedit = New System.Windows.Forms.Button()
        Me.btnadd = New System.Windows.Forms.Button()
        Me.txtLoaiMaLoai = New System.Windows.Forms.TextBox()
        Me.txtChitiet = New System.Windows.Forms.TextBox()
        Me.txtDongia = New System.Windows.Forms.TextBox()
        Me.txtTensp = New System.Windows.Forms.TextBox()
        Me.txtMasp = New System.Windows.Forms.TextBox()
        Me.grbThemSanPham.SuspendLayout()
        CType(Me.nudSoluong, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvConnect, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblMaSP
        '
        Me.lblMaSP.AutoSize = True
        Me.lblMaSP.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaSP.Location = New System.Drawing.Point(29, 37)
        Me.lblMaSP.Name = "lblMaSP"
        Me.lblMaSP.Size = New System.Drawing.Size(62, 20)
        Me.lblMaSP.TabIndex = 0
        Me.lblMaSP.Text = "Mã SP:"
        '
        'lblTenSanPham
        '
        Me.lblTenSanPham.AutoSize = True
        Me.lblTenSanPham.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTenSanPham.Location = New System.Drawing.Point(29, 90)
        Me.lblTenSanPham.Name = "lblTenSanPham"
        Me.lblTenSanPham.Size = New System.Drawing.Size(118, 20)
        Me.lblTenSanPham.TabIndex = 1
        Me.lblTenSanPham.Text = "Tên Sản Phẩm:"
        '
        'lblDongia
        '
        Me.lblDongia.AutoSize = True
        Me.lblDongia.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDongia.Location = New System.Drawing.Point(29, 142)
        Me.lblDongia.Name = "lblDongia"
        Me.lblDongia.Size = New System.Drawing.Size(75, 20)
        Me.lblDongia.TabIndex = 2
        Me.lblDongia.Text = "Đơn Giá:"
        '
        'lblSoLuong
        '
        Me.lblSoLuong.AutoSize = True
        Me.lblSoLuong.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSoLuong.Location = New System.Drawing.Point(29, 191)
        Me.lblSoLuong.Name = "lblSoLuong"
        Me.lblSoLuong.Size = New System.Drawing.Size(83, 20)
        Me.lblSoLuong.TabIndex = 3
        Me.lblSoLuong.Text = "Số Lượng:"
        '
        'lblChiTiet
        '
        Me.lblChiTiet.AutoSize = True
        Me.lblChiTiet.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblChiTiet.Location = New System.Drawing.Point(370, 37)
        Me.lblChiTiet.Name = "lblChiTiet"
        Me.lblChiTiet.Size = New System.Drawing.Size(94, 20)
        Me.lblChiTiet.TabIndex = 4
        Me.lblChiTiet.Text = "Chi Tiết SP:"
        '
        'lblLoaispMaLoai
        '
        Me.lblLoaispMaLoai.AutoSize = True
        Me.lblLoaispMaLoai.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLoaispMaLoai.Location = New System.Drawing.Point(370, 90)
        Me.lblLoaispMaLoai.Name = "lblLoaispMaLoai"
        Me.lblLoaispMaLoai.Size = New System.Drawing.Size(196, 20)
        Me.lblLoaispMaLoai.TabIndex = 5
        Me.lblLoaispMaLoai.Text = "Loại Sản Phẩm_ Mã Loại:"
        '
        'lblNgayNhap
        '
        Me.lblNgayNhap.AutoSize = True
        Me.lblNgayNhap.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNgayNhap.Location = New System.Drawing.Point(370, 142)
        Me.lblNgayNhap.Name = "lblNgayNhap"
        Me.lblNgayNhap.Size = New System.Drawing.Size(138, 20)
        Me.lblNgayNhap.TabIndex = 6
        Me.lblNgayNhap.Text = "Ngày Nhập Hàng:"
        '
        'grbThemSanPham
        '
        Me.grbThemSanPham.BackColor = System.Drawing.SystemColors.Window
        Me.grbThemSanPham.Controls.Add(Me.nudSoluong)
        Me.grbThemSanPham.Controls.Add(Me.dtpNgayNhapHang)
        Me.grbThemSanPham.Controls.Add(Me.dgvConnect)
        Me.grbThemSanPham.Controls.Add(Me.btnclose)
        Me.grbThemSanPham.Controls.Add(Me.btnDelete)
        Me.grbThemSanPham.Controls.Add(Me.btnedit)
        Me.grbThemSanPham.Controls.Add(Me.btnadd)
        Me.grbThemSanPham.Controls.Add(Me.txtLoaiMaLoai)
        Me.grbThemSanPham.Controls.Add(Me.txtChitiet)
        Me.grbThemSanPham.Controls.Add(Me.txtDongia)
        Me.grbThemSanPham.Controls.Add(Me.txtTensp)
        Me.grbThemSanPham.Controls.Add(Me.txtMasp)
        Me.grbThemSanPham.Controls.Add(Me.lblMaSP)
        Me.grbThemSanPham.Controls.Add(Me.lblNgayNhap)
        Me.grbThemSanPham.Controls.Add(Me.lblTenSanPham)
        Me.grbThemSanPham.Controls.Add(Me.lblLoaispMaLoai)
        Me.grbThemSanPham.Controls.Add(Me.lblSoLuong)
        Me.grbThemSanPham.Controls.Add(Me.lblChiTiet)
        Me.grbThemSanPham.Controls.Add(Me.lblDongia)
        Me.grbThemSanPham.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grbThemSanPham.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grbThemSanPham.ForeColor = System.Drawing.SystemColors.ControlText
        Me.grbThemSanPham.Location = New System.Drawing.Point(0, 0)
        Me.grbThemSanPham.Name = "grbThemSanPham"
        Me.grbThemSanPham.Size = New System.Drawing.Size(723, 522)
        Me.grbThemSanPham.TabIndex = 7
        Me.grbThemSanPham.TabStop = False
        Me.grbThemSanPham.Text = "Quản Lý Sản Phẩm"
        '
        'nudSoluong
        '
        Me.nudSoluong.Location = New System.Drawing.Point(33, 214)
        Me.nudSoluong.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.nudSoluong.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudSoluong.Name = "nudSoluong"
        Me.nudSoluong.Size = New System.Drawing.Size(79, 27)
        Me.nudSoluong.TabIndex = 4
        Me.nudSoluong.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.nudSoluong.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'dtpNgayNhapHang
        '
        Me.dtpNgayNhapHang.CustomFormat = ""
        Me.dtpNgayNhapHang.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpNgayNhapHang.Location = New System.Drawing.Point(374, 161)
        Me.dtpNgayNhapHang.Name = "dtpNgayNhapHang"
        Me.dtpNgayNhapHang.Size = New System.Drawing.Size(289, 27)
        Me.dtpNgayNhapHang.TabIndex = 7
        Me.dtpNgayNhapHang.Value = New Date(2015, 3, 30, 0, 0, 0, 0)
        '
        'dgvConnect
        '
        Me.dgvConnect.AllowUserToOrderColumns = True
        Me.dgvConnect.BackgroundColor = System.Drawing.Color.DimGray
        Me.dgvConnect.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvConnect.Location = New System.Drawing.Point(3, 256)
        Me.dgvConnect.Name = "dgvConnect"
        Me.dgvConnect.Size = New System.Drawing.Size(717, 263)
        Me.dgvConnect.TabIndex = 12
        '
        'btnclose
        '
        Me.btnclose.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnclose.Location = New System.Drawing.Point(617, 212)
        Me.btnclose.Name = "btnclose"
        Me.btnclose.Size = New System.Drawing.Size(75, 27)
        Me.btnclose.TabIndex = 11
        Me.btnclose.Text = "Đóng"
        Me.btnclose.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(536, 212)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(75, 27)
        Me.btnDelete.TabIndex = 10
        Me.btnDelete.Text = "Xóa"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnedit
        '
        Me.btnedit.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnedit.Location = New System.Drawing.Point(455, 212)
        Me.btnedit.Name = "btnedit"
        Me.btnedit.Size = New System.Drawing.Size(75, 27)
        Me.btnedit.TabIndex = 9
        Me.btnedit.Text = "Sửa"
        Me.btnedit.UseVisualStyleBackColor = True
        '
        'btnadd
        '
        Me.btnadd.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnadd.Location = New System.Drawing.Point(374, 212)
        Me.btnadd.Name = "btnadd"
        Me.btnadd.Size = New System.Drawing.Size(75, 27)
        Me.btnadd.TabIndex = 8
        Me.btnadd.Text = "Thêm"
        Me.btnadd.UseVisualStyleBackColor = True
        '
        'txtLoaiMaLoai
        '
        Me.txtLoaiMaLoai.Location = New System.Drawing.Point(374, 112)
        Me.txtLoaiMaLoai.MaxLength = 20
        Me.txtLoaiMaLoai.Name = "txtLoaiMaLoai"
        Me.txtLoaiMaLoai.Size = New System.Drawing.Size(289, 27)
        Me.txtLoaiMaLoai.TabIndex = 6
        '
        'txtChitiet
        '
        Me.txtChitiet.Location = New System.Drawing.Point(374, 60)
        Me.txtChitiet.MaxLength = 50
        Me.txtChitiet.Name = "txtChitiet"
        Me.txtChitiet.Size = New System.Drawing.Size(289, 27)
        Me.txtChitiet.TabIndex = 5
        '
        'txtDongia
        '
        Me.txtDongia.Location = New System.Drawing.Point(33, 161)
        Me.txtDongia.Name = "txtDongia"
        Me.txtDongia.Size = New System.Drawing.Size(289, 27)
        Me.txtDongia.TabIndex = 3
        '
        'txtTensp
        '
        Me.txtTensp.Location = New System.Drawing.Point(33, 112)
        Me.txtTensp.MaxLength = 30
        Me.txtTensp.Name = "txtTensp"
        Me.txtTensp.Size = New System.Drawing.Size(289, 27)
        Me.txtTensp.TabIndex = 2
        '
        'txtMasp
        '
        Me.txtMasp.Location = New System.Drawing.Point(33, 60)
        Me.txtMasp.MaxLength = 20
        Me.txtMasp.Name = "txtMasp"
        Me.txtMasp.Size = New System.Drawing.Size(289, 27)
        Me.txtMasp.TabIndex = 1
        '
        'frmProduct_Management
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(723, 522)
        Me.Controls.Add(Me.grbThemSanPham)
        Me.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmProduct_Management"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Product Management"
        Me.grbThemSanPham.ResumeLayout(False)
        Me.grbThemSanPham.PerformLayout()
        CType(Me.nudSoluong, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvConnect, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblMaSP As System.Windows.Forms.Label
    Friend WithEvents lblTenSanPham As System.Windows.Forms.Label
    Friend WithEvents lblDongia As System.Windows.Forms.Label
    Friend WithEvents lblSoLuong As System.Windows.Forms.Label
    Friend WithEvents lblChiTiet As System.Windows.Forms.Label
    Friend WithEvents lblLoaispMaLoai As System.Windows.Forms.Label
    Friend WithEvents lblNgayNhap As System.Windows.Forms.Label
    Friend WithEvents grbThemSanPham As System.Windows.Forms.GroupBox
    Friend WithEvents dgvConnect As System.Windows.Forms.DataGridView
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnedit As System.Windows.Forms.Button
    Friend WithEvents btnadd As System.Windows.Forms.Button
    Friend WithEvents txtLoaiMaLoai As System.Windows.Forms.TextBox
    Friend WithEvents txtChitiet As System.Windows.Forms.TextBox
    Friend WithEvents txtDongia As System.Windows.Forms.TextBox
    Friend WithEvents txtTensp As System.Windows.Forms.TextBox
    Friend WithEvents txtMasp As System.Windows.Forms.TextBox
    Friend WithEvents dtpNgayNhapHang As System.Windows.Forms.DateTimePicker
    Friend WithEvents nudSoluong As System.Windows.Forms.NumericUpDown
End Class
