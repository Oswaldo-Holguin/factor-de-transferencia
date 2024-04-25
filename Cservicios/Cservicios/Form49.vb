Public Class Frm49RecibosProvisionales
    Public I As Integer
    Public _Renglon As DataGridViewRow
    Public Msg, _Recibo, _Paciente As String
    Public Style As MsgBoxStyle
    Public _Checacontrol As Control
    Public _Busca As DataRow

    Private Sub Form49_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Button2_Click(sender, e)

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet32.VentaFactor' Puede moverla o quitarla según sea necesario.
        Me.VentaFactorTableAdapter.Fill(Me.CServiciosDataSet32.VentaFactor)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet40.Recibos' Puede moverla o quitarla según sea necesario.
        Me.RecibosTableAdapter.Fill(Me.CServiciosDataSet40.Recibos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet30.Clientes' Puede moverla o quitarla según sea necesario.
        Me.ClientesTableAdapter.Fill(Me.CServiciosDataSet30.Clientes)

        DataGridView1.Rows.Clear()
        For Me.I = 0 To CServiciosDataSet40.Recibos.Rows.Count - 1
            _Recibo = CServiciosDataSet40.Recibos.Rows(I).Item("Recibo")
            _Paciente = CServiciosDataSet40.Recibos.Rows(I).Item("Paciente")

            DataGridView1.Rows.Add(_Recibo, _Paciente)
        Next
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = False
                _Checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""

        ToolStripButton4.Enabled = False

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        _Renglon = DataGridView1.CurrentRow
        If _Renglon.Cells(0).Value = "" Then
            Exit Sub
        End If
        _Renglon.DefaultCellStyle.BackColor = Color.Gold

        Dim _Cliente, _NombreCliente, _Folio As String
        Dim _Rows_Vf() As DataRow
        Dim _Lote As String

        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""
        CheckBox1.Checked = False


        TextBox1.Text = _Renglon.Cells(0).Value
        TextBox1.Enabled = False
        _Recibo = TextBox1.Text

        _Busca = CServiciosDataSet40.Recibos.FindByRecibo(_Recibo)
        _Cliente = _Busca.Item("Cliente")
        TextBox2.Text = _Busca.Item("Fecha")
        TextBox3.Text = _Busca.Item("NumeroFrascos")
        TextBox4.Text = _Busca.Item("Comprador")
        TextBox5.Text = _Busca.Item("PrecioUnitario")
        TextBox7.Text = _Busca.Item("Importe")
        TextBox12.Text = _Busca.Item("Paciente")
        TextBox11.Text = _Busca.Item("TelefonoRecetas")
        If _Busca.Item("Prorroga") = True Then
            TextBox16.Text = "PRÓRROGA"
        End If
        If _Busca.Item("Donacion") = True Then
            TextBox16.Text = "DONACIÓN"
        End If

        If _Busca.Item("Recetas") = True Then
            TextBox16.Text = "RECETAS"
        End If
        RichTextBox1.Text = _Busca.Item("Comentarios")
        TextBox15.Text = _Busca.Item("Folio")

        TextBox8.Text = _Busca.Item("Referencia")
        TextBox10.Text = _Busca.Item("Cedula")
        TextBox6.Text = _Busca.Item("NumReceta")
        If DBNull.Value.Equals(_Busca.Item("FechaReceta")) Then
            TextBox14.Text = ""
        Else
            TextBox14.Text = _Busca.Item("FechaReceta")

        End If
        If _Busca.Item("Comprobado") = True Then
            TextBox21.Text = _Busca.Item("Documento")
        End If

        CheckBox1.Checked = _Busca.Item("Comprobado")
        TextBox13.Text = _Busca.Item("Edad")

        _Folio = TextBox15.Text

        _Rows_Vf = CServiciosDataSet32.VentaFactor.Select("Folio = " & "'" & _Folio & "'")

        _Lote = ""
        For Me.I = 0 To _Rows_Vf.GetUpperBound(0)
            TextBox9.Text = _Rows_Vf(I).Item("Medico")
            _Lote = _Rows_Vf(I).Item("Referencia")
        Next

        '   _Busca = CServiciosDataSet32.VentaFactor.FindByFolio(_Folio)
        '   TextBox9.Text = _Busca.Item("Medico")

        _Busca = CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
        _NombreCliente = _Busca.Item("NombreCliente")

        Label8.Text = _NombreCliente
        TextBox17.Text = _Cliente

        ' Busca el Lote.
        TextBox18.Text = _Lote

        Panel1.Visible = False
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        _Renglon = DataGridView1.CurrentRow
        _Renglon.DefaultCellStyle.BackColor = Color.White

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub ToolStripButton3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton3.Click
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = True
            End If
        Next
        ToolStripButton4.Enabled = True
        TextBox1.Enabled = False
        TextBox15.Enabled = False

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click
        ' Guarda los Datos

        Dim _Motivo As String = TextBox16.Text

        _Recibo = TextBox1.Text
        _Busca = CServiciosDataSet40.Recibos.FindByRecibo(_Recibo)


        _Busca.Item("Fecha") = CDate(TextBox2.Text)
        _Busca.Item("NumeroFrascos") = TextBox3.Text
        _Busca.Item("PrecioUnitario") = TextBox5.Text
        _Busca.Item("Importe") = TextBox7.Text
        _Busca.Item("Comprador") = TextBox4.Text
        _Busca.Item("Paciente") = TextBox12.Text
        _Busca.Item("Cliente") = TextBox17.Text
        _Busca.Item("TelefonoRecetas") = TextBox11.Text
        _Busca.Item("Prorroga") = False
        _Busca.Item("Donacion") = False
        _Busca.Item("Recetas") = False

        If TextBox16.Text = "PRÓRROGA" Then
            _Busca.Item("Prorroga") = True
        End If
        If TextBox16.Text = "DONACIÓN" Then
            _Busca.Item("Donacion") = True
        End If
        If TextBox16.Text = "RECETAS" Then
            _Busca.Item("Recetas") = True
        End If

        _Busca.Item("Donacion") = RadioButton2.Checked
        _Busca.Item("Recetas") = RadioButton3.Checked
        _Busca.Item("Comentarios") = UCase(RichTextBox1.Text)
        _Busca.Item("Referencia") = TextBox8.Text
        _Busca.Item("FechaRegistro") = Today
        _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
        _Busca.Item("Medico") = TextBox9.Text
        _Busca.Item("Cedula") = TextBox10.Text
        _Busca.Item("NumReceta") = TextBox6.Text
        If TextBox14.Text = "N/A" Then
            _Busca.Item("FechaReceta") = DBNull.Value
        Else
            If IsDate(TextBox14.Text) Then
                _Busca.Item("FechaReceta") = TextBox14.Text
            Else
                _Busca.Item("FechaReceta") = DBNull.Value
            End If
        End If

        _Busca.Item("Documento") = ""
        _Busca.Item("Comprobado") = False
        _Busca.Item("Edad") = TextBox13.Text

        Me.Validate()
        RecibosBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet40)

        Msg = "Registro guardado Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

    End Sub

    Private Sub ToolStripButton6_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton6.Click
        ' Imprime el Recibo
        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Recibo Provisional.xlsx"
        Dim _Recibo As String = TextBox1.Text
        Dim _Concepto, _Referencia, _Comprador As String
        Dim _Fecha As Date = CDate(TextBox2.Text)
        Dim _Cantidad, _PrecioUnitario, _Importe, _Mide, _Renglon As Integer
        Dim _Comentarios As String = RichTextBox1.Text
        Dim _Linea As String



        _Concepto = TextBox16.Text
        _Cantidad = CInt(TextBox3.Text)
        _PrecioUnitario = CInt(TextBox5.Text)
        _Importe = CInt(TextBox7.Text)
        _Referencia = TextBox8.Text
        _Comprador = TextBox4.Text

        m_Excel = CreateObject("Excel.Application")
        m_Excel.Workbooks.Open(strRutaExcel)
        m_Excel.Visible = False

        m_Excel.Worksheets("RECIBO").cells(7, 3).value = _Fecha
        m_Excel.Worksheets("RECIBO").cells(8, 3).value = _Recibo
        m_Excel.Worksheets("RECIBO").cells(10, 9).value = _Cantidad
        m_Excel.Worksheets("RECIBO").cells(12, 8).value = _Concepto


        m_Excel.Worksheets("RECIBO").cells(8, 3).value = _Recibo
        m_Excel.Worksheets("RECIBO").cells(17, 4).value = _Cantidad
        m_Excel.Worksheets("RECIBO").cells(17, 5).value = _PrecioUnitario
        m_Excel.Worksheets("RECIBO").cells(17, 6).value = _Importe


        m_Excel.Worksheets("RECIBO").cells(18, 2).value = Label8.Text

        m_Excel.Worksheets("RECIBO").cells(21, 3).value = _Referencia
        m_Excel.Worksheets("RECIBO").cells(21, 8).value = _Comprador
        m_Excel.Worksheets("RECIBO").cells(23, 3).value = _Fecha
        m_Excel.Worksheets("RECIBO").cells(23, 8).value = _Fecha

        ' imprime los comentarios
        If _Comentarios <> "" Then
            _Linea = _Comentarios
            _Mide = Len(_Comentarios)
            If _Mide > 100 Then
                _Renglon = 30
                For Me.I = 1 To _Mide Step 100
                    _Linea = Mid(_Comentarios, I, 100)
                    m_Excel.Worksheets("RECIBO").cells(_Renglon, 2).value = _Linea
                    _Renglon = _Renglon + 1
                Next

            Else
                m_Excel.Worksheets("RECIBO").cells(30, 2).value = _Linea
            End If


        End If


        m_Excel.Visible = True
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox1.GotFocus

    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        ' Busqueda
        Label18.Text = "Intituto"
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Text = ""
            End If
        Next
        DataGridView1.Rows.Clear()
        RichTextBox1.Text = ""

        Dim _Folio, _Institucion, _Medico, _Texto, _Partenombre As String
        Dim _Rows_Rp() As DataRow
        Dim _Mensaje, _Titulo, _Cliente As String
        Dim _Fecha As Date
        Dim _Mide As Integer
        Dim _Parte, _NombreCliente As String
        Dim X As Integer

        _NombreCliente = ""
        _Cliente = ""
        _Institucion = ""
        _Medico = ""
        _Titulo = "Búsqueda de Recibos Provisionales"
        _Parte = ""
        '-------------------------------------------------------------------------------------------
        If RadioButton1.Checked = True Then
            _Mensaje = "Teclee al Recibo a Buscar"
            _Recibo = InputBox(_Mensaje, _Titulo)
            _Recibo = UCase(_Recibo)

            _Rows_Rp = CServiciosDataSet40.Recibos.Select("Recibo = " & "'" & _Recibo & "'")
        End If
        ' -------------------------------------------------------------------------------------------
        If RadioButton2.Checked = True Then
            _Mensaje = "Teclee la Fecha a Buscar"
            _Fecha = InputBox(_Mensaje, _Titulo)

            _Rows_Rp = CServiciosDataSet40.Recibos.Select("Fecha = " & "'" & _Fecha & "'")
        End If
        ' -------------------------------------------------------------------------------------------
        If RadioButton6.Checked = True Then
            _Rows_Rp = CServiciosDataSet40.Recibos.Select("Recibo <> " & "'" & "X" & "'")
        End If
        ' -------------------------------------------------------------------------------------------
        If RadioButton7.Checked = True Then
            _Mensaje = "Teclee la Folio de Venta a Buscar"
            _Folio = InputBox(_Mensaje, _Titulo)

            _Rows_Rp = CServiciosDataSet40.Recibos.Select("Folio = " & "'" & _Folio & "'")
        End If

        ' -------------------------------------------------------------------------------------------
        If RadioButton8.Checked = True Then
            _Rows_Rp = CServiciosDataSet40.Recibos.Select("Comprobado = " & "'" & True & "'")
        End If
        ' -------------------------------------------------------------------------------------------
        If RadioButton9.Checked = True Then
            _Rows_Rp = CServiciosDataSet40.Recibos.Select("Comprobado = " & "'" & False & "'")
        End If


        ' -------------------------------------------------------------------------------------------
        If RadioButton1.Checked = True Or RadioButton2.Checked = True Or RadioButton6.Checked = True Or RadioButton7.Checked = True Or RadioButton8.Checked = True Or RadioButton9.Checked = True Then

            For Me.I = 0 To _Rows_Rp.GetUpperBound(0)
                _Recibo = _Rows_Rp(I).Item("Recibo")
                _Paciente = _Rows_Rp(I).Item("Paciente")

                DataGridView1.Rows.Add(_Recibo, _Paciente)
            Next

        End If
        ' -------------------------------------------------------------------------------------------
        If RadioButton3.Checked = True Then
            _Mensaje = "Teclee el nombre del Paciente a Buscar"
            _Parte = UCase(InputBox(_Mensaje, _Titulo))

            If _Parte = "" Then
                Msg = "Debe teclear un Nombre de Paciente o una parte de Nombre de Paciente"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If
        End If
        ' -------------------------------------------------------------------------------------------
        If RadioButton4.Checked = True Then
            _Mensaje = "Teclee el nombre de la Institucion a Buscar"
            _Parte = UCase(InputBox(_Mensaje, _Titulo))

            _Mide = Len(_Parte)
            For Me.I = 0 To CServiciosDataSet40.Recibos.Rows.Count - 1
                _Cliente = CServiciosDataSet40.Recibos.Rows(I).Item("Cliente")
                _Recibo = CServiciosDataSet40.Recibos.Rows(I).Item("Recibo")
                _Paciente = CServiciosDataSet40.Recibos.Rows(I).Item("Paciente")

                _Busca = CServiciosDataSet30.Clientes.FindByCliente(_Cliente)
                _NombreCliente = _Busca.Item("NombreCliente")
                _Texto = _NombreCliente

                For X = 1 To Len(_Texto)
                    _Partenombre = Mid(_Texto, X, _Mide)

                    If _Partenombre = _Parte Then
                        DataGridView1.Rows.Add(_Recibo, _Paciente)
                    End If
                Next

            Next

        End If
        ' -------------------------------------------------------------------------------------------
        If RadioButton5.Checked = True Then
            _Mensaje = "Teclee el nombre del Médico a Buscar"
            _Parte = UCase(InputBox(_Mensaje, _Titulo))

            If _Parte = "" Then
                Msg = "Debe teclear un Nombre de Médico a Buscar"
                Style = MsgBoxStyle.Information
                MsgBox(Msg, Style)
                Exit Sub
            End If
        End If

        If RadioButton3.Checked = True = True Or RadioButton5.Checked = True Then
            _Texto = ""
            _Mide = Len(_Parte)
            For Me.I = 0 To CServiciosDataSet40.Recibos.Rows.Count - 1
                If RadioButton3.Checked = True Then
                    _Texto = CServiciosDataSet40.Recibos.Rows(I).Item("Paciente")
                End If
                If RadioButton5.Checked = True Then
                    _Texto = CServiciosDataSet40.Recibos.Rows(I).Item("Medico")
                End If
                _Recibo = CServiciosDataSet40.Recibos.Rows(I).Item("Recibo")
                _Paciente = CServiciosDataSet40.Recibos.Rows(I).Item("Paciente")

                For X = 1 To Len(_Texto)
                    _Partenombre = Mid(_Texto, X, _Mide)

                    If _Partenombre = _Parte Then
                        DataGridView1.Rows.Add(_Recibo, _Paciente)
                    End If
                Next

            Next

        End If




    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton6.CheckedChanged

    End Sub

    Private Sub ToolStripButton5_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton5.Click
        ' Comprobación del Recibo
        For Each Me._Checacontrol In Me.Controls
            If TypeOf _Checacontrol Is TextBox Then
                _Checacontrol.Enabled = False
            End If
        Next
        Panel1.Visible = True
        Panel1.Enabled = True
        TextBox19.Text = Date.Today
        TextBox20.Focus()


    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        ' Comprobar
        If TextBox19.Text = "" Then
            Msg = "Debe seleccionar una fecha de Comprobación"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        If TextBox20.Text = "" Then
            Msg = "Debe proporcionar un Folio de Venta para poder Comprobar este Recibo"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If
        Dim _FechaComprobacion As Date = CDate(TextBox19.Text)
        If Not IsDate(_FechaComprobacion) Then
            Msg = "Debe proporcionar una fecha de Comprobación Válida"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        ' Valida que no exista ese folio

        Dim _Folio As String = TextBox20.Text
        Dim _Rows_vf() As DataRow = CServiciosDataSet32.VentaFactor.Select("Folio = " & "'" & _Folio & "'")
        Dim _Find As Boolean
        Dim _FechaFolio As Date

        _Find = False

        For Me.I = 0 To _Rows_vf.GetUpperBound(0)
            _Find = True
            _FechaFolio = _Rows_vf(I).Item("FechaFolio")
            _Paciente = _Rows_vf(I).Item("NombrePaciente")
        Next

        If _Find Then
            Msg = "Ese folio ya está registrado con fecha " & CStr(_FechaFolio) & " para el paciente " & _Paciente
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        _Recibo = TextBox1.Text
        _Busca = CServiciosDataSet40.Recibos.FindByRecibo(_Recibo)
        _Busca.Item("Documento") = TextBox20.Text & " " & "Fecha : " & TextBox19.Text
        _Busca.Item("Comprobado") = True

        Me.Validate()
        RecibosBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(CServiciosDataSet40)
        Me.RecibosTableAdapter.Fill(Me.CServiciosDataSet40.Recibos)

        Msg = "Recibo Comprobado Correctamente"
        Style = MsgBoxStyle.Information
        MsgBox(Msg, Style)

        Panel1.Visible = False
        Button2_Click(sender, e)

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Button2_Click(sender, e)
        Panel1.Visible = False
    End Sub
End Class