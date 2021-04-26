Public Class FormMenuUtama
    Sub Terkunci()
        LoginToolStripMenuItem.Enabled = True
        LogoutToolStripMenuItem.Enabled = False
        MasterToolStripMenuItem.Enabled = False
        TransaksiToolStripMenuItem.Enabled = False
        UtilityToolStripMenuItem.Enabled = False
        STLabel2.Text = ""
        STLabel4.Text = ""
        STLabel6.Text = ""
    End Sub

    Private Sub FormMenuUtama_Load(ByVal sender As Object, e As EventArgs) Handles MyBase.Load
        Call Terkunci()
        STLabel10.Text = Today
    End Sub
    Private Sub LoginToolStripMenu_Item_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoginToolStripMenuItem.Click
        FormLogin.ShowDialog()
    End Sub
    Private Sub LogoutToolStripMenu_Item_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        Call Terkunci()
    End Sub
    Private Sub KeluarToolStripMenu_Item_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles KeluarToolStripMenuItem.Click
        End
    End Sub
    Private Sub AdminToolStripMenu_Item_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AdminToolStripMenuItem.Click
        FormMasterAdmin.ShowDialog()
    End Sub
    Private Sub PelangganToolStripMenu_Item_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PelangganToolStripMenuItem.Click
        FormMasterPelanggan.ShowDialog()
    End Sub
    Private Sub BarangToolStripMenu_Item_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BarangToolStripMenuItem.Click
        FormMasterBarang.ShowDialog()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        STLabel8.Text = TimeOfDay
    End Sub
    Private Sub PenjualanToolStripMenu_Item_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PenjualanToolStripMenuItem.Click
        FormTransaksiPenjualan.ShowDialog()
    End Sub
    Private Sub GantiPasswordToolStripMenu_Item_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GantiPasswordToolStripMenuItem.Click
        FormGantiPassword.ShowDialog()
    End Sub
End Class
