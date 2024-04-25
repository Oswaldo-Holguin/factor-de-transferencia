Public Class LoginForm1

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If Trim(TextBox4.Text) <> Trim(TextBox3.Text) Then
            MsgBox("El Password es Incorrecto. Verifique")
            TextBox3.Text = ""
            TextBox3.Focus()
            Exit Sub
        End If
        If ComboBox1.Text = "" Then
            MsgBox("Debe seleccionar un Usuario")
            Exit Sub
        End If


        Frm2Menu.Show()
        Frm2Menu.ToolStripTextBox1.Text = ComboBox1.Text
        Frm2Menu.ToolStripTextBox2.Text = TextBox2.Text
        Frm2Menu.ToolStripTextBox3.Text = TextBox5.Text
        Me.Hide()
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub LoginForm1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '  MsgBox("Hola")
        'TODO: This line of code loads data into the 'CServiciosDataSet.Centros' table. You can move, or remove it, as needed.
        Me.CentrosTableAdapter.Fill(Me.CServiciosDataSet.Centros)
        ' MsgBox("Paso1")
        'TODO: This line of code loads data into the 'CServiciosDataSet1.Usuarios' table. You can move, or remove it, as needed.
        Me.UsuariosTableAdapter.Fill(Me.CServiciosDataSet1.Usuarios)
        Dim i As Integer
        For i = 0 To CServiciosDataSet1.Usuarios.Rows.Count - 1
            ComboBox1.Items.Add(CServiciosDataSet1.Usuarios.Rows(i).Item("Nombre"))
        Next

    End Sub

    Private Sub ComboBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.Click
     
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim i As Integer
        Dim Busca As DataRow

        For i = 0 To CServiciosDataSet1.Usuarios.Rows.Count - 1
            If CServiciosDataSet1.Usuarios.Rows(i).Item("Nombre") = ComboBox1.Text Then
                TextBox1.Text = CServiciosDataSet1.Usuarios.Rows(i).Item("Centro")
                Busca = CServiciosDataSet.Centros.FindByCentro(TextBox1.Text)
                TextBox2.Text = Busca.Item("Descripcion")
                TextBox4.Text = CServiciosDataSet1.Usuarios.Rows(i).Item("Password")
                TextBox5.Text = CServiciosDataSet1.Usuarios.Rows(i).Item("Nivel")
                TextBox6.Text = CServiciosDataSet1.Usuarios.Rows(i).Item("Usuario")
            End If
        Next
        TextBox3.Focus()
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub
End Class
