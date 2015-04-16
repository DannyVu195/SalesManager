Imports System.Data.DataTable

Imports System.Data.SqlClient

Public Class frmPersonnel_Management
    Dim Connection As String = "workstation id=anhvdps02128.mssql.somee.com;packet size=4096;user id=ps02128;pwd=anhvdps02128;data source=anhvdps02128.mssql.somee.com;persist security info=False;initial catalog=anhvdps02128"
    Dim anhvdps02128 As New DataTable

    Private Sub DataUpdateAll()
        Dim KetNoi As SqlConnection = New SqlConnection(Connection)
        Dim TruyVanChuyenDoi As New SqlDataAdapter("Select * from NhanVien", KetNoi)
        Try
            KetNoi.Open()
            TruyVanChuyenDoi.Fill(anhvdps02128)
        Catch ex As Exception

        End Try
        dgvDataGridView.DataSource = anhvdps02128
        KetNoi.Close()
    End Sub

    Private Sub frmPersonnel_Management_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataUpdateAll()
        btnDelete.Enabled = False
        btnedit.Enabled = False
End Sub

    Private Sub btnclose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnclose.Click
        Dim dialog As DialogResult = MessageBox.Show("You want to end application?",
                                    "Sales Manager!-anhvdps02128", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (dialog = DialogResult.Yes) Then
            frmMain_Sales_Manager.Show()
            Me.Close()
        End If
    End Sub

    Private Sub dgvDataGridView_CellClick(ByVal sender As System.Object, ByVal e As DataGridViewCellEventArgs) Handles dgvDataGridView.CellClick
        Dim Focus As Integer = dgvDataGridView.CurrentCell.RowIndex
        txtMaNV.Text = dgvDataGridView.Item(0, Focus).Value.ToString()
        txtTenNV.Text = dgvDataGridView.Item(1, Focus).Value.ToString()
        txtMatKhau.Text = dgvDataGridView.Item(2, Focus).Value.ToString()
        txtDiaChi.Text = dgvDataGridView.Item(3, Focus).Value.ToString()
        txtSoDT.Text = dgvDataGridView.Item(4, Focus).Value.ToString()
        cbbGioiTinh.Text = dgvDataGridView.Item(5, Focus).Value.ToString()
        btnadd.Enabled = True
        btnDelete.Enabled = True
        btnedit.Enabled = True
    End Sub

    Private Sub btnedit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnedit.Click
        Dim KetNoi As SqlConnection = New SqlConnection(Connection)
        KetNoi.Open()
        Dim Edit As String = " Update NhanVien set TenNhanVien= @Ten, MatKhauTruyCap= @MatKhau, DiaChi= @DiaChi, SoDienThoai= @SoDT, GioiTinh= @GioiTinh where MaNhanVien=@Ma"
        Dim SPedit As New SqlCommand(Edit, KetNoi)
        Try
            SPedit.Parameters.AddWithValue("@Ma", txtMaNV.Text)
            SPedit.Parameters.AddWithValue("@Ten", txtTenNV.Text)
            SPedit.Parameters.AddWithValue("@MaKhau", txtMatKhau.Text)
            SPedit.Parameters.AddWithValue("@DiaChi", txtDiaChi.Text)
            SPedit.Parameters.AddWithValue("@SoDT", txtSoDT.Text)
            SPedit.Parameters.AddWithValue("@GioiTinh", cbbGioiTinh.Text)
            SPedit.ExecuteNonQuery()
            KetNoi.Close()
            MessageBox.Show("Edit Successful Data", "Personnel Management! Anhvdps02128")
        Catch ex As Exception
            MessageBox.Show("Edit Connection Error Data", "Personnel Management! Anhvdps02128")
        End Try
        anhvdps02128.Clear()
        dgvDataGridView.DataSource = anhvdps02128
        dgvDataGridView.DataSource = Nothing
        DataUpdateAll()
        btnedit.Enabled = False
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim Ketnoi As New SqlConnection(Connection)
        Ketnoi.Open()
        Dim Delete As String = " Delete from NhanVien where MaNhanVien=@Ma"
        Dim SPdelete As New SqlCommand(Delete, Ketnoi)
        Try
            SPdelete.Parameters.AddWithValue("@Ma", txtMaNV.Text)
            SPdelete.ExecuteNonQuery()
            Ketnoi.Close()
            MessageBox.Show("Delete Successful Data", "Personnel Management! Anhvdps02128")
        Catch ex As Exception
            MessageBox.Show("Delete Connection Error Data", "Personnel Management! Anhvdps02128")
        End Try
        anhvdps02128.Clear()
        dgvDataGridView.DataSource = anhvdps02128
        dgvDataGridView.DataSource = Nothing
        DataUpdateAll()
        btnDelete.Enabled = False
    End Sub

    Private Sub btnadd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnadd.Click
        Dim KetNoi As SqlConnection = New SqlConnection(Connection)
        KetNoi.Open()
        Dim Add As String = " Insert into NhanVien values (@Ma, @Ten, @MatKhau, @DiaChi, @SoDT, @GioiTinh)"
        Dim SPadd As New SqlCommand(Add, KetNoi)
        Try
            SPadd.Parameters.AddWithValue("@Ma", txtMaNV.Text)
            SPadd.Parameters.AddWithValue("@Ten", txtTenNV.Text)
            SPadd.Parameters.AddWithValue("@MatKhau", txtMatKhau.Text)
            SPadd.Parameters.AddWithValue("@DiaChi", txtDiaChi.Text)
            SPadd.Parameters.AddWithValue("@SoDT", txtSoDT.Text)
            SPadd.Parameters.AddWithValue("@GioiTinh", cbbGioiTinh.Text)
            SPadd.ExecuteNonQuery()
            KetNoi.Close()
            MessageBox.Show("Add Successful Data", "Product Management! Anhvdps02128")
        Catch ex As Exception
            MessageBox.Show("Add Connection Error Data", "Product Management! Anhvdps02128")
        End Try
        anhvdps02128.Clear()
        dgvDataGridView.DataSource = anhvdps02128
        dgvDataGridView.DataSource = Nothing
        DataUpdateAll()
        btnadd.Enabled = False
        txtMaNV.Clear()
        txtTenNV.Clear()
        txtMatKhau.Clear()
        txtDiaChi.Clear()
        txtSoDT.Clear()
        cbbGioiTinh.ResetText()
    End Sub
End Class