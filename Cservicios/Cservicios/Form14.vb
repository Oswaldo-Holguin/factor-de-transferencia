Public Class Form14

    Private Sub CitasBindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub Form14_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet10.Empleados' Puede moverla o quitarla según sea necesario.
        Me.EmpleadosTableAdapter.Fill(Me.CServiciosDataSet10.Empleados)
        'TODO: This line of code loads data into the 'CServiciosDataSet9.Citas' table. You can move, or remove it, as needed.
        ' Me.CitasTableAdapter.Fill(Me.CServiciosDataSet9.Citas)

    End Sub
End Class