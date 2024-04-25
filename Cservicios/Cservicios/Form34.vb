Public Class Frm34SubMenuCxC

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Frm33Clientes.Show()

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Frm60Facturas.Show()

    End Sub

    Private Sub ClienteToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ClienteToolStripMenuItem.Click
        Frm33Clientes.Show()
    End Sub

    Private Sub FacturasToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles FacturasToolStripMenuItem.Click
        Frm60Facturas.Show()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Frm61Cobranza.Show()

    End Sub

    Private Sub CobranzaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CobranzaToolStripMenuItem.Click
        Button3_Click(sender, e)
    End Sub
End Class