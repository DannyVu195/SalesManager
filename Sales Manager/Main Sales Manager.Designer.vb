<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain_Sales_Manager
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain_Sales_Manager))
        Me.mnsMenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.mniHethong = New System.Windows.Forms.ToolStripMenuItem()
        Me.ĐăngNhậpLạiToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ThoátToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mniQuantri = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuảnLýSảnPhẩmToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuảnLýKháchHàngToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuảnLýSảnPhẩmToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuảnLýHóaĐơnToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuảnLýLoạiSảnPhẩmToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.mniHotro = New System.Windows.Forms.ToolStripMenuItem()
        Me.stsStatusStrip = New System.Windows.Forms.StatusStrip()
        Me.lblCopyright = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblVuduyanh = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblEmail = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblAnhvdps02128 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ptbLogo = New System.Windows.Forms.PictureBox()
        Me.mnsMenuStrip1.SuspendLayout()
        Me.stsStatusStrip.SuspendLayout()
        CType(Me.ptbLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'mnsMenuStrip1
        '
        Me.mnsMenuStrip1.BackColor = System.Drawing.Color.Lavender
        Me.mnsMenuStrip1.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnsMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mniHethong, Me.mniQuantri, Me.mniHotro})
        Me.mnsMenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.mnsMenuStrip1.Name = "mnsMenuStrip1"
        Me.mnsMenuStrip1.Size = New System.Drawing.Size(642, 27)
        Me.mnsMenuStrip1.TabIndex = 0
        Me.mnsMenuStrip1.Text = "MenuStrip1"
        '
        'mniHethong
        '
        Me.mniHethong.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ĐăngNhậpLạiToolStripMenuItem, Me.ThoátToolStripMenuItem})
        Me.mniHethong.Name = "mniHethong"
        Me.mniHethong.Size = New System.Drawing.Size(92, 23)
        Me.mniHethong.Text = "Hệ Thống"
        '
        'ĐăngNhậpLạiToolStripMenuItem
        '
        Me.ĐăngNhậpLạiToolStripMenuItem.Name = "ĐăngNhậpLạiToolStripMenuItem"
        Me.ĐăngNhậpLạiToolStripMenuItem.Size = New System.Drawing.Size(179, 24)
        Me.ĐăngNhậpLạiToolStripMenuItem.Text = "Đăng nhập lại"
        '
        'ThoátToolStripMenuItem
        '
        Me.ThoátToolStripMenuItem.Name = "ThoátToolStripMenuItem"
        Me.ThoátToolStripMenuItem.Size = New System.Drawing.Size(179, 24)
        Me.ThoátToolStripMenuItem.Text = "Thoát"
        '
        'mniQuantri
        '
        Me.mniQuantri.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.QuảnLýSảnPhẩmToolStripMenuItem1, Me.QuảnLýKháchHàngToolStripMenuItem, Me.QuảnLýSảnPhẩmToolStripMenuItem, Me.QuảnLýHóaĐơnToolStripMenuItem, Me.QuảnLýLoạiSảnPhẩmToolStripMenuItem})
        Me.mniQuantri.Name = "mniQuantri"
        Me.mniQuantri.Size = New System.Drawing.Size(87, 23)
        Me.mniQuantri.Text = "Quản Trị"
        '
        'QuảnLýSảnPhẩmToolStripMenuItem1
        '
        Me.QuảnLýSảnPhẩmToolStripMenuItem1.Name = "QuảnLýSảnPhẩmToolStripMenuItem1"
        Me.QuảnLýSảnPhẩmToolStripMenuItem1.Size = New System.Drawing.Size(240, 24)
        Me.QuảnLýSảnPhẩmToolStripMenuItem1.Text = "Quản lý sản phẩm"
        '
        'QuảnLýKháchHàngToolStripMenuItem
        '
        Me.QuảnLýKháchHàngToolStripMenuItem.Name = "QuảnLýKháchHàngToolStripMenuItem"
        Me.QuảnLýKháchHàngToolStripMenuItem.Size = New System.Drawing.Size(240, 24)
        Me.QuảnLýKháchHàngToolStripMenuItem.Text = "Quản lý khách hàng"
        '
        'QuảnLýSảnPhẩmToolStripMenuItem
        '
        Me.QuảnLýSảnPhẩmToolStripMenuItem.Name = "QuảnLýSảnPhẩmToolStripMenuItem"
        Me.QuảnLýSảnPhẩmToolStripMenuItem.Size = New System.Drawing.Size(240, 24)
        Me.QuảnLýSảnPhẩmToolStripMenuItem.Text = "Quản lý nhân viên"
        '
        'QuảnLýHóaĐơnToolStripMenuItem
        '
        Me.QuảnLýHóaĐơnToolStripMenuItem.Name = "QuảnLýHóaĐơnToolStripMenuItem"
        Me.QuảnLýHóaĐơnToolStripMenuItem.Size = New System.Drawing.Size(240, 24)
        Me.QuảnLýHóaĐơnToolStripMenuItem.Text = "Quản lý hóa đơn"
        '
        'QuảnLýLoạiSảnPhẩmToolStripMenuItem
        '
        Me.QuảnLýLoạiSảnPhẩmToolStripMenuItem.Name = "QuảnLýLoạiSảnPhẩmToolStripMenuItem"
        Me.QuảnLýLoạiSảnPhẩmToolStripMenuItem.Size = New System.Drawing.Size(240, 24)
        Me.QuảnLýLoạiSảnPhẩmToolStripMenuItem.Text = "Quản lý loại sản phẩm"
        '
        'mniHotro
        '
        Me.mniHotro.Name = "mniHotro"
        Me.mniHotro.Size = New System.Drawing.Size(73, 23)
        Me.mniHotro.Text = "Hỗ Trợ"
        '
        'stsStatusStrip
        '
        Me.stsStatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblCopyright, Me.lblVuduyanh, Me.lblEmail, Me.lblAnhvdps02128})
        Me.stsStatusStrip.Location = New System.Drawing.Point(0, 362)
        Me.stsStatusStrip.Name = "stsStatusStrip"
        Me.stsStatusStrip.Size = New System.Drawing.Size(642, 22)
        Me.stsStatusStrip.TabIndex = 1
        Me.stsStatusStrip.Text = "StatusStrip1"
        '
        'lblCopyright
        '
        Me.lblCopyright.Font = New System.Drawing.Font("Times New Roman", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCopyright.Name = "lblCopyright"
        Me.lblCopyright.Size = New System.Drawing.Size(64, 17)
        Me.lblCopyright.Text = "Copyright:"
        '
        'lblVuduyanh
        '
        Me.lblVuduyanh.Name = "lblVuduyanh"
        Me.lblVuduyanh.Size = New System.Drawing.Size(70, 17)
        Me.lblVuduyanh.Text = "Vũ Duy Anh"
        '
        'lblEmail
        '
        Me.lblEmail.Font = New System.Drawing.Font("Times New Roman", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(42, 17)
        Me.lblEmail.Text = "Email:"
        '
        'lblAnhvdps02128
        '
        Me.lblAnhvdps02128.Name = "lblAnhvdps02128"
        Me.lblAnhvdps02128.Size = New System.Drawing.Size(149, 17)
        Me.lblAnhvdps02128.Text = "Anhvdps02128@fpt.edu.vn"
        '
        'ptbLogo
        '
        Me.ptbLogo.BackgroundImage = Global.Sales_Manager.My.Resources.Resources.logo
        Me.ptbLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.ptbLogo.ErrorImage = Global.Sales_Manager.My.Resources.Resources.logo
        Me.ptbLogo.Location = New System.Drawing.Point(11, 38)
        Me.ptbLogo.Name = "ptbLogo"
        Me.ptbLogo.Size = New System.Drawing.Size(149, 57)
        Me.ptbLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.ptbLogo.TabIndex = 5
        Me.ptbLogo.TabStop = False
        '
        'frmMain_Sales_Manager
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImage = Global.Sales_Manager.My.Resources.Resources.QTDN1
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(642, 384)
        Me.Controls.Add(Me.ptbLogo)
        Me.Controls.Add(Me.stsStatusStrip)
        Me.Controls.Add(Me.mnsMenuStrip1)
        Me.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Font = New System.Drawing.Font("Times New Roman", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.mnsMenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmMain_Sales_Manager"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Main Sales Manager"
        Me.mnsMenuStrip1.ResumeLayout(False)
        Me.mnsMenuStrip1.PerformLayout()
        Me.stsStatusStrip.ResumeLayout(False)
        Me.stsStatusStrip.PerformLayout()
        CType(Me.ptbLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents mnsMenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents stsStatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents lblCopyright As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblVuduyanh As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents mniHethong As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mniQuantri As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mniHotro As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblEmail As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblAnhvdps02128 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ptbLogo As System.Windows.Forms.PictureBox
    Friend WithEvents ĐăngNhậpLạiToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ThoátToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents QuảnLýSảnPhẩmToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents QuảnLýKháchHàngToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents QuảnLýHóaĐơnToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents QuảnLýSảnPhẩmToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents QuảnLýLoạiSảnPhẩmToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
