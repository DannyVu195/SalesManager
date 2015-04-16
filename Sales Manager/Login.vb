Imports System.Data.SqlClient

Public Class frmLogin

    Private Sub txtUsername_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtUsername.KeyDown
        If e.KeyCode = Keys.Enter Then txtPass.Focus()
    End Sub

    Private Sub txtPass_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPass.KeyDown
        If e.KeyCode = Keys.Enter Then btnSingin.PerformClick()
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Dim dialog As DialogResult = MessageBox.Show("You want to end application?",
                                    "Sales Manager!-anhvdps02128", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (dialog = DialogResult.Yes) Then
            Me.Close()
            Application.Exit()
        End If
    End Sub

    Private Sub btnSingin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSingin.Click
        Dim Connection As String = "workstation id=anhvdps02128.mssql.somee.com;packet size=4096;user id=ps02128;pwd=anhvdps02128;data source=anhvdps02128.mssql.somee.com;persist security info=False;initial catalog=anhvdps02128"
        Dim KetNoi As SqlConnection = New SqlConnection(Connection)
        Dim TruyVanChuyenDoi As New SqlDataAdapter("Select * from NhanVien where MaNhanVien= '" & txtUsername.Text & "' and MatKhauTruyCap = '" & txtPass.Text & "' ", KetNoi)
        Dim anhvdps02128 As New DataTable
        Try
            KetNoi.Open()
            TruyVanChuyenDoi.Fill(anhvdps02128)
            If anhvdps02128.Rows.Count > 0 Then
                MessageBox.Show("Wellcome to Sales Manader! Congratulations you have successfully logged into my application!", "Sales Manager!-anhvdps02128", MessageBoxButtons.OK, MessageBoxIcon.Information)
                frmMain_Sales_Manager.Show()
                Me.Hide()
            Else
                MessageBox.Show("Wrong login into my Application! Please check the information", "Sales Manager!-anhvdps02128", MessageBoxButtons.OK, MessageBoxIcon.Error)
                txtPass.Clear()
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class
