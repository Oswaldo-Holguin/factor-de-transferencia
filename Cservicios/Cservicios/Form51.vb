Public Class Frm51Procedimientos

    Private Sub ToolStripLabel1_Click(sender As System.Object, e As System.EventArgs)


    End Sub

    Private Sub ProcedimientoFactorDeTransferenciaToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProcedimientoFactorDeTransferenciaToolStripMenuItem.Click
        ' Abre el Archivo de Word
        Process.Start("Z:\Formatos Factor\Procedimiento Factor de Transferencia.docx")
    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()
    End Sub
End Class