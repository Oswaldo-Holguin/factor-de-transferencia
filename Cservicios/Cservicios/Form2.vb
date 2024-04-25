Public Class Frm2Menu

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        LoginForm1.Show()
        LoginForm1.Close()
        Me.Close()

    End Sub

    Private Sub Button36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button36.Click

        Frm1Centros.Show()

    End Sub

    Private Sub Frm2Menu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If CInt(LoginForm1.TextBox5.Text) = 4 Then
            Panel4.Enabled = True
        Else
            If CInt(LoginForm1.TextBox5.Text) < 5 Then
                Panel1.Enabled = False
                Panel4.Enabled = False
                Panel6.Enabled = False
            End If
        End If
 
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Frm3Pacientes.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Frm6TablaMateriales.Show()

    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Frm3Pacientes.Show()
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Frm3Pacientes.Show()

    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        Frm3Pacientes.Show()

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Frm6TablaMateriales.Show()

    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        'Frm6TablaMateriales.Show()
        Frm49RecibosProvisionales.Show()

    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Frm3Pacientes.Show()

    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        Frm6TablaMateriales.Show()

    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        Frm11Tablas.Show()

    End Sub

    Private Sub Button35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button35.Click
        Frm8Empleados.Show()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Frm10Formulas.Show()

    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Frm10Formulas.Show()

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Frm13ElectrosParaHoy.Show()

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        ' Muestra todas las citas

    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Frm18SubMenuConsultas.Show()

    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Frm18SubMenuConsultas.Show()

    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        Frm18SubMenuConsultas.Show()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Frm19ConsultadeCitas.Show()

    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Frm19ConsultadeCitas.Show()

    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        Frm19ConsultadeCitas.Show()

    End Sub

    Private Sub Button33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button33.Click
        Frm24consultasMedicas.Show()

    End Sub

    Private Sub Button32_Click(sender As System.Object, e As System.EventArgs) Handles Button32.Click
        'Frm28VentaFT.Show()
        Frm63VentaFactor.Show()

    End Sub

    Private Sub Button22_Click(sender As System.Object, e As System.EventArgs) Handles Button22.Click
        Frm29FT.Show()

    End Sub

    Private Sub Button39_Click(sender As System.Object, e As System.EventArgs) Handles Button39.Click
        Frm34SubMenuCxC.Show()
    End Sub

    Private Sub Button19_Click(sender As System.Object, e As System.EventArgs) Handles Button19.Click
        Frm37InformeMensual.Show()

    End Sub

    Private Sub Button37_Click(sender As System.Object, e As System.EventArgs) Handles Button37.Click
        Frm24consultasMedicas.Show()

    End Sub

    Private Sub Button21_Click(sender As System.Object, e As System.EventArgs) Handles Button21.Click
        Frm38Envios.Show()

    End Sub

    Private Sub Button29_Click(sender As System.Object, e As System.EventArgs) Handles Button29.Click

    End Sub

    Private Sub Button38_Click(sender As System.Object, e As System.EventArgs) Handles Button38.Click
        Frm24consultasMedicas.Show()

    End Sub

    Private Sub Button40_Click(sender As System.Object, e As System.EventArgs) Handles Button40.Click
        Frm45Reservaciones.Show()

    End Sub

    Private Sub Button46_Click(sender As System.Object, e As System.EventArgs) Handles Button46.Click
        Frm42Recursos.Show()

    End Sub

    Private Sub Button43_Click(sender As System.Object, e As System.EventArgs) Handles Button43.Click
        Frm43Equipos.Show()

    End Sub

    Private Sub Button45_Click(sender As System.Object, e As System.EventArgs) Handles Button45.Click
        Frm45Reservaciones.Show()

    End Sub

    Private Sub Button41_Click(sender As System.Object, e As System.EventArgs) Handles Button41.Click
        Frm46ReservacionesConsecutivas.Show()

    End Sub

    Private Sub Button48_Click(sender As System.Object, e As System.EventArgs) Handles Button48.Click
        Frm50Solicitudes.Show()

    End Sub

    Private Sub Button44_Click(sender As System.Object, e As System.EventArgs) Handles Button44.Click
        'Frm43Equipos.Show()
        Frm6TablaMateriales.Show()
    End Sub

    Private Sub Button42_Click(sender As System.Object, e As System.EventArgs) Handles Button42.Click
        Frm45Reservaciones.Show()

    End Sub

    Private Sub Button23_Click(sender As System.Object, e As System.EventArgs) Handles Button23.Click
        Frm51Procedimientos.Show()

    End Sub

    Private Sub Button20_Click(sender As System.Object, e As System.EventArgs) Handles Button20.Click
        Frm52Bitacoras.Show()

    End Sub

    Private Sub Button54_Click(sender As System.Object, e As System.EventArgs) Handles Button54.Click
        Frm54Requisiciones.Show()

    End Sub

    Private Sub Button55_Click(sender As System.Object, e As System.EventArgs) Handles Button55.Click
        Frm55Recepcion.Show()
    End Sub

    Private Sub Button56_Click(sender As System.Object, e As System.EventArgs) Handles Button56.Click
        Frm56Inventarios.Show()
    End Sub

    Private Sub Button58_Click(sender As System.Object, e As System.EventArgs) Handles Button58.Click
        Frm19ConsultadeCitas.Show()

    End Sub

    Private Sub Button57_Click(sender As System.Object, e As System.EventArgs) Handles Button57.Click
        Frm3Pacientes.Show()

    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        Frm27TablaProductos.Show()

    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        Frm58Envios.Show()
    End Sub

    Private Sub Button60_Click(sender As System.Object, e As System.EventArgs) Handles Button60.Click
        Frm54Requisiciones.Show()

    End Sub

    Private Sub Button61_Click(sender As System.Object, e As System.EventArgs) Handles Button61.Click
        Frm55Recepcion.Show()

    End Sub

    Private Sub Button62_Click(sender As System.Object, e As System.EventArgs) Handles Button62.Click
        Frm56Inventarios.Show()

    End Sub

    Private Sub Button63_Click(sender As System.Object, e As System.EventArgs) Handles Button63.Click
        Frm59HistorialRemisiones.Show()

    End Sub

    Private Sub Button52_Click(sender As System.Object, e As System.EventArgs) Handles Button52.Click
        Frm50Solicitudes.Show()
    End Sub

    Private Sub Button59_Click(sender As System.Object, e As System.EventArgs) Handles Button59.Click
        Frm62HistorialElectros.Show()

    End Sub

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        Frm34SubMenuCxC.Show()

    End Sub

    Private Sub Button64_Click(sender As Object, e As EventArgs) Handles Button64.Click
        Frm63VentaFactor.Show()

    End Sub

    Private Sub Button47_Click(sender As Object, e As EventArgs) Handles Button47.Click
        Frm54Requisiciones.Show()
    End Sub

    Private Sub Button65_Click(sender As Object, e As EventArgs) Handles Button65.Click
        Frm6TablaMateriales.Show()
    End Sub
End Class