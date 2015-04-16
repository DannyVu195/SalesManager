Imports System.Data.DataTable

Imports System.Data.SqlClient

Public Class frmCustomer_Management
    Dim Connection As String = "workstation id=anhvdps02128.mssql.somee.com;packet size=4096;user id=ps02128;pwd=anhvdps02128;data source=anhvdps02128.mssql.somee.com;persist security info=False;initial catalog=anhvdps02128"
    Dim anhvdps02128 As New DataTable

    Private Sub DataUpdateAll()
        Dim KetNoi As SqlConnection = New SqlConnection(Connection)
        Dim TruyVanChuyenDoi As New SqlDataAdapter("Select * from KhachHang", KetNoi)
        Try
            KetNoi.Open()
            TruyVanChuyenDoi.Fill(anhvdps02128)
        Catch ex As Exception

        End Try
        dgvDataGridView.DataSource = anhvdps02128
        KetNoi.Close()
    End Sub

    Private Sub frmCustomer_Management_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataUpdateAll()
        btnDelete.Enabled = False
        btnedit.Enabled = False
    End Sub

    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Dim dialog As DialogResult = MessageBox.Show("You want to End Application?",
                                    "Customer Management!Anhvdps02128", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (dialog = DialogResult.Yes) Then
            frmMain_Sales_Manager.Show()
            Me.Close()
        End If
    End Sub

    Private Sub btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnadd.Click
        Dim KetNoi As SqlConnection = New SqlConnection(Connection)
        KetNoi.Open()
        Dim Add As String = " Insert into KhachHang values (@Ma, @Ten, @DiaChi, @SoDT, @GioiTinh, @TonGiao)"
        Dim SPadd As New SqlCommand(Add, KetNoi)
        Try
            SPadd.Parameters.AddWithValue("@Ma", txtMaKH.Text)
            SPadd.Parameters.AddWithValue("@Ten", txtTenKH.Text)
            SPadd.Parameters.AddWithValue("@DiaChi", txtDiaChi.Text)
            SPadd.Parameters.AddWithValue("@SoDT", txtSoDienThoai.Text)
            SPadd.Parameters.AddWithValue("@GioiTinh", cbbGioiTinh.Text)
            SPadd.Parameters.AddWithValue("@TonGiao", cbbTonGiao.Text)
            SPadd.ExecuteNonQuery()
            KetNoi.Close()
            MessageBox.Show("Add successful data", "Product Management! Anhvdps02128")
        Catch ex As Exception
            MessageBox.Show("Add connection error data", "Product Management! Anhvdps02128")
        End Try
        anhvdps02128.Clear()
        dgvDataGridView.DataSource = anhvdps02128
        dgvDataGridView.DataSource = Nothing
        DataUpdateAll()
        txtMaKH.Clear()
        txtTenKH.Clear()
        txtDiaChi.Clear()
        txtSoDienThoai.Clear()
        cbbGioiTinh.ResetText()
        cbbTonGiao.ResetText()
        btnadd.Enabled = False
    End Sub

    Private Sub btnedit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnedit.Click
        Dim KetNoi As SqlConnection = New SqlConnection(Connection)
        KetNoi.Open()
        Dim Edit As String = " update KhachHang set TenKhachHang=@Ten, DiaChi=@DiaChi, SoDienThoai=@SoDT, GioiTinh=@GioiTinh, TonGiao=@TonGiao where MaKhachHang=@Ma"
        Dim SPedit As New SqlCommand(Edit, KetNoi)
        Try
            SPedit.Parameters.AddWithValue("@Ma", txtMaKH.Text)
            SPedit.Parameters.AddWithValue("@Ten", txtTenKH.Text)
            SPedit.Parameters.AddWithValue("@DiaChi", txtDiaChi.Text)
            SPedit.Parameters.AddWithValue("@SoDT", txtSoDienThoai.Text)
            SPedit.Parameters.AddWithValue("@GioiTinh", cbbGioiTinh.Text)
            SPedit.Parameters.AddWithValue("@TonGiao", cbbTonGiao.Text)
            SPedit.ExecuteNonQuery()
            KetNoi.Close()
            MessageBox.Show("Edit successful data", "Customer Management! Anhvdps02128")
        Catch ex As Exception
            MessageBox.Show("Edit connection error data", "Customer Management! Anhvdps02128")
        End Try
        anhvdps02128.Clear()
        dgvDataGridView.DataSource = anhvdps02128
        dgvDataGridView.DataSource = Nothing
        DataUpdateAll()
        btnedit.Enabled = False
    End Sub

    Private Sub dgvDataGridView_CellClick(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) Handles dgvDataGridView.CellClick
        Dim Focus As Integer = dgvDataGridView.CurrentCell.RowIndex
        txtMaKH.Text = dgvDataGridView.Item(0, Focus).Value.ToString()
        txtTenKH.Text = dgvDataGridView.Item(1, Focus).Value.ToString()
        txtDiaChi.Text = dgvDataGridView.Item(2, Focus).Value.ToString()
        txtSoDienThoai.Text = dgvDataGridView.Item(3, Focus).Value.ToString()
        cbbGioiTinh.Text = dgvDataGridView.Item(4, Focus).Value.ToString()
        cbbTonGiao.Text = dgvDataGridView.Item(5, Focus).Value.ToString()
        btnedit.Enabled = True
        btnDelete.Enabled = True
        btnadd.Enabled = True
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim Ketnoi As New SqlConnection(Connection)
        Ketnoi.Open()
        Dim Delete As String = " Delete from KhachHang where MaKhachHang=@Ma"
        Dim SPdelete As New SqlCommand(Delete, Ketnoi)
        Try
            SPdelete.Parameters.AddWithValue("@Ma", txtMaKH.Text)
            SPdelete.ExecuteNonQuery()
            Ketnoi.Close()
            MessageBox.Show("Delete successful data", "Customer Management! Anhvdps02128")
        Catch ex As Exception
            MessageBox.Show("Delete connection error data", "Customer Management! Anhvdps02128")
        End Try
        anhvdps02128.Clear()
        dgvDataGridView.DataSource = anhvdps02128
        dgvDataGridView.DataSource = Nothing
        DataUpdateAll()
        btnDelete.Enabled = False
    End Sub
End Class