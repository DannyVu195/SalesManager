Public Class frmMain_Sales_Manager


    Private Sub mniHotro_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mniHotro.Click
        MessageBox.Show("Wellcome to support Sales Manader! Mọi thông tin chi tiết vui lòng liên hệ tới Anhvdps02128.@fpt.edu.vn! or sdt: 01667109966", "Sales Manager!-anhvdps02128", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ThoátToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ThoátToolStripMenuItem.Click
        Dim dialog As DialogResult = MessageBox.Show("You want to end application?",
                                    "Sales Manager!-anhvdps02128", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If (dialog = DialogResult.Yes) Then
            Me.Close()
            Application.Exit()
        End If
    End Sub

    Private Sub ĐăngNhậpLạiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ĐăngNhậpLạiToolStripMenuItem.Click
        Me.Close()
        frmLogin.Show()
        frmLogin.txtUsername.Clear()
        frmLogin.txtPass.Clear()
    End Sub

    Private Sub QuảnLýSảnPhẩmToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuảnLýSảnPhẩmToolStripMenuItem1.Click
        frmProduct_Management.Show()
    End Sub

    Private Sub QuảnLýSảnPhẩmToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuảnLýSảnPhẩmToolStripMenuItem.Click
        frmPersonnel_Management.Show()
    End Sub

    Private Sub QuảnLýKháchHàngToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuảnLýKháchHàngToolStripMenuItem.Click
        frmCustomer_Management.Show()
    End Sub

    Private Sub QuảnLýHóaĐơnToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuảnLýHóaĐơnToolStripMenuItem.Click
        frmProduct_Management.Show()
    End Sub

    Private Sub QuảnLýLoạiSảnPhẩmToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles QuảnLýLoạiSảnPhẩmToolStripMenuItem.Click
        frmProduct_Management.Show()
    End Sub
End Class