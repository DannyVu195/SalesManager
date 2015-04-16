<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPersonnel_Management
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPersonnel_Management))
        Me.grbQuanLyNhanVien = New System.Windows.Forms.GroupBox()
        Me.btnclose = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnedit = New System.Windows.Forms.Button()
        Me.btnadd = New System.Windows.Forms.Button()
        Me.dgvDataGridView = New System.Windows.Forms.DataGridView()
        Me.cbbGioiTinh = New System.Windows.Forms.ComboBox()
        Me.txtSoDT = New System.Windows.Forms.TextBox()
        Me.txtDiaChi = New System.Windows.Forms.TextBox()
        Me.txtMatKhau = New System.Windows.Forms.TextBox()
        Me.txtTenNV = New System.Windows.Forms.TextBox()
        Me.txtMaNV = New System.Windows.Forms.TextBox()
        Me.lblSoDienThoai = New System.Windows.Forms.Label()
        Me.lblGioiTinh = New System.Windows.Forms.Label()
        Me.lblDiaChi = New System.Windows.Forms.Label()
        Me.LblMatKhau = New System.Windows.Forms.Label()
        Me.lblTenNhanVien = New System.Windows.Forms.Label()
        Me.lblMaNhanVien = New System.Windows.Forms.Label()
        Me.grbQuanLyNhanVien.SuspendLayout()
        CType(Me.dgvDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grbQuanLyNhanVien
        '
        Me.grbQuanLyNhanVien.Controls.Add(Me.btnclose)
        Me.grbQuanLyNhanVien.Controls.Add(Me.btnDelete)
        Me.grbQuanLyNhanVien.Controls.Add(Me.btnedit)
        Me.grbQuanLyNhanVien.Controls.Add(Me.btnadd)
        Me.grbQuanLyNhanVien.Controls.Add(Me.dgvDataGridView)
        Me.grbQuanLyNhanVien.Controls.Add(Me.cbbGioiTinh)
        Me.grbQuanLyNhanVien.Controls.Add(Me.txtSoDT)
        Me.grbQuanLyNhanVien.Controls.Add(Me.txtDiaChi)
        Me.grbQuanLyNhanVien.Controls.Add(Me.txtMatKhau)
        Me.grbQuanLyNhanVien.Controls.Add(Me.txtTenNV)
        Me.grbQuanLyNhanVien.Controls.Add(Me.txtMaNV)
        Me.grbQuanLyNhanVien.Controls.Add(Me.lblSoDienThoai)
        Me.grbQuanLyNhanVien.Controls.Add(Me.lblGioiTinh)
        Me.grbQuanLyNhanVien.Controls.Add(Me.lblDiaChi)
        Me.grbQuanLyNhanVien.Controls.Add(Me.LblMatKhau)
        Me.grbQuanLyNhanVien.Controls.Add(Me.lblTenNhanVien)
        Me.grbQuanLyNhanVien.Controls.Add(Me.lblMaNhanVien)
        Me.grbQuanLyNhanVien.Location = New System.Drawing.Point(0, 0)
        Me.grbQuanLyNhanVien.Name = "grbQuanLyNhanVien"
        Me.grbQuanLyNhanVien.Size = New System.Drawing.Size(684, 502)
        Me.grbQuanLyNhanVien.TabIndex = 0
        Me.grbQuanLyNhanVien.TabStop = False
        Me.grbQuanLyNhanVien.Text = "Quản Lý Nhân Viên"
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
        Me.dgvDataGridView.Size = New System.Drawing.Size(678, 251)
        Me.dgvDataGridView.TabIndex = 12
        '
        'cbbGioiTinh
        '
        Me.cbbGioiTinh.FormattingEnabled = True
        Me.cbbGioiTinh.Items.AddRange(New Object() {"Nam", "Nữ", "3D"})
        Me.cbbGioiTinh.Location = New System.Drawing.Point(361, 168)
        Me.cbbGioiTinh.Name = "cbbGioiTinh"
        Me.cbbGioiTinh.Size = New System.Drawing.Size(75, 27)
        Me.cbbGioiTinh.TabIndex = 6
        '
        'txtSoDT
        '
        Me.txtSoDT.Location = New System.Drawing.Point(361, 117)
        Me.txtSoDT.MaxLength = 15
        Me.txtSoDT.Name = "txtSoDT"
        Me.txtSoDT.Size = New System.Drawing.Size(289, 27)
        Me.txtSoDT.TabIndex = 5
        '
        'txtDiaChi
        '
        Me.txtDiaChi.Location = New System.Drawing.Point(361, 63)
        Me.txtDiaChi.MaxLength = 50
        Me.txtDiaChi.Name = "txtDiaChi"
        Me.txtDiaChi.Size = New System.Drawing.Size(289, 27)
        Me.txtDiaChi.TabIndex = 4
        '
        'txtMatKhau
        '
        Me.txtMatKhau.Location = New System.Drawing.Point(32, 168)
        Me.txtMatKhau.MaxLength = 20
        Me.txtMatKhau.Name = "txtMatKhau"
        Me.txtMatKhau.Size = New System.Drawing.Size(289, 27)
        Me.txtMatKhau.TabIndex = 3
        '
        'txtTenNV
        '
        Me.txtTenNV.Location = New System.Drawing.Point(32, 117)
        Me.txtTenNV.MaxLength = 30
        Me.txtTenNV.Name = "txtTenNV"
        Me.txtTenNV.Size = New System.Drawing.Size(289, 27)
        Me.txtTenNV.TabIndex = 2
        '
        'txtMaNV
        '
        Me.txtMaNV.Location = New System.Drawing.Point(32, 63)
        Me.txtMaNV.MaxLength = 20
        Me.txtMaNV.Name = "txtMaNV"
        Me.txtMaNV.Size = New System.Drawing.Size(289, 27)
        Me.txtMaNV.TabIndex = 1
        '
        'lblSoDienThoai
        '
        Me.lblSoDienThoai.AutoSize = True
        Me.lblSoDienThoai.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSoDienThoai.Location = New System.Drawing.Point(357, 95)
        Me.lblSoDienThoai.Name = "lblSoDienThoai"
        Me.lblSoDienThoai.Size = New System.Drawing.Size(115, 20)
        Me.lblSoDienThoai.TabIndex = 11
        Me.lblSoDienThoai.Text = "Số Điện Thoại:"
        '
        'lblGioiTinh
        '
        Me.lblGioiTinh.AutoSize = True
        Me.lblGioiTinh.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGioiTinh.Location = New System.Drawing.Point(357, 146)
        Me.lblGioiTinh.Name = "lblGioiTinh"
        Me.lblGioiTinh.Size = New System.Drawing.Size(83, 20)
        Me.lblGioiTinh.TabIndex = 10
        Me.lblGioiTinh.Text = "Giới Tính:"
        '
        'lblDiaChi
        '
        Me.lblDiaChi.AutoSize = True
        Me.lblDiaChi.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDiaChi.Location = New System.Drawing.Point(357, 41)
        Me.lblDiaChi.Name = "lblDiaChi"
        Me.lblDiaChi.Size = New System.Drawing.Size(70, 20)
        Me.lblDiaChi.TabIndex = 9
        Me.lblDiaChi.Text = "Địa Chỉ:"
        '
        'LblMatKhau
        '
        Me.LblMatKhau.AutoSize = True
        Me.LblMatKhau.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblMatKhau.Location = New System.Drawing.Point(28, 146)
        Me.LblMatKhau.Name = "LblMatKhau"
        Me.LblMatKhau.Size = New System.Drawing.Size(155, 20)
        Me.LblMatKhau.TabIndex = 8
        Me.LblMatKhau.Text = "Mật Khẩu Truy Cập:"
        '
        'lblTenNhanVien
        '
        Me.lblTenNhanVien.AutoSize = True
        Me.lblTenNhanVien.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTenNhanVien.Location = New System.Drawing.Point(28, 95)
        Me.lblTenNhanVien.Name = "lblTenNhanVien"
        Me.lblTenNhanVien.Size = New System.Drawing.Size(121, 20)
        Me.lblTenNhanVien.TabIndex = 7
        Me.lblTenNhanVien.Text = "Tên Nhân Viên:"
        '
        'lblMaNhanVien
        '
        Me.lblMaNhanVien.AutoSize = True
        Me.lblMaNhanVien.Font = New System.Drawing.Font("Times New Roman", 12.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaNhanVien.Location = New System.Drawing.Point(28, 41)
        Me.lblMaNhanVien.Name = "lblMaNhanVien"
        Me.lblMaNhanVien.Size = New System.Drawing.Size(118, 20)
        Me.lblMaNhanVien.TabIndex = 6
        Me.lblMaNhanVien.Text = "Mã Nhân Viên:"
        '
        'frmPersonnel_Management
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(680, 501)
        Me.Controls.Add(Me.grbQuanLyNhanVien)
        Me.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmPersonnel_Management"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Personnel Management"
        Me.grbQuanLyNhanVien.ResumeLayout(False)
        Me.grbQuanLyNhanVien.PerformLayout()
        CType(Me.dgvDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents grbQuanLyNhanVien As System.Windows.Forms.GroupBox
    Friend WithEvents lblSoDienThoai As System.Windows.Forms.Label
    Friend WithEvents lblGioiTinh As System.Windows.Forms.Label
    Friend WithEvents lblDiaChi As System.Windows.Forms.Label
    Friend WithEvents LblMatKhau As System.Windows.Forms.Label
    Friend WithEvents lblTenNhanVien As System.Windows.Forms.Label
    Friend WithEvents lblMaNhanVien As System.Windows.Forms.Label
    Friend WithEvents txtMaNV As System.Windows.Forms.TextBox
    Friend WithEvents txtDiaChi As System.Windows.Forms.TextBox
    Friend WithEvents txtMatKhau As System.Windows.Forms.TextBox
    Friend WithEvents txtTenNV As System.Windows.Forms.TextBox
    Friend WithEvents txtSoDT As System.Windows.Forms.TextBox
    Friend WithEvents cbbGioiTinh As System.Windows.Forms.ComboBox
    Friend WithEvents dgvDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents btnclose As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnedit As System.Windows.Forms.Button
    Friend WithEvents btnadd As System.Windows.Forms.Button
End Class
