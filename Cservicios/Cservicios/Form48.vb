Public Class Frm48ReciboProvisional
    Public I As Integer
    Public _Busca As DataRow
    Public Msg As String
    Public Style As MsgBoxStyle



    Private Sub Frm48ReciboProvisional_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet40.Recibos' Puede moverla o quitarla según sea necesario.
        Me.RecibosTableAdapter1.Fill(Me.CServiciosDataSet40.Recibos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet39.Recibos' Puede moverla o quitarla según sea necesario.
        '  Me.RecibosTableAdapter.Fill(Me.CServiciosDataSet39.Recibos)
        'TODO: esta línea de código carga datos en la tabla 'CServiciosDataSet3.Documentos' Puede moverla o quitarla según sea necesario.
        Me.DocumentosTableAdapter.Fill(Me.CServiciosDataSet3.Documentos)

        Button1_Click(sender, e)

    End Sub

    Private Sub ToolStripButton1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton1.Click
        Me.Close()

    End Sub

    Private Sub ToolStripButton2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton2.Click
        Button1_Click(sender, e)


    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim _Consecutivo As Integer
        For Me.I = 0 To CServiciosDataSet3.Documentos.Rows.Count - 1
            If CServiciosDataSet3.Documentos.Rows(I).Item("Documento") = "REP" Then
                _Consecutivo = CServiciosDataSet3.Documentos.Rows(I).Item("Consecutivo")
            End If
        Next
        Dim Checacontrol As Control
        For Each Checacontrol In Me.Controls
            If TypeOf Checacontrol Is TextBox Then
                Checacontrol.Text = ""
            End If
        Next
        RichTextBox1.Text = ""

        Dim _XConsecutivo As String = "000" + CStr(_Consecutivo)
        Dim _Mide As Integer = Len(_XConsecutivo)
        Dim _x As Integer = _Mide - 3
        Dim _PrecioUnitario, _Importe, _NumeroFrascos As Integer

        _XConsecutivo = "RP" + Mid(_XConsecutivo, _x, 4)
        TextBox1.Text = _XConsecutivo
        TextBox2.Text = Today

        _NumeroFrascos = Frm28VentaFT.TextBox5.Text
        _PrecioUnitario = Frm28VentaFT.TextBox4.Text
        _Importe = _NumeroFrascos * _PrecioUnitario

        TextBox3.Text = _NumeroFrascos
        '  TextBox5.Text = Format(_PrecioUnitario, "$##,##0.00")
        '  TextBox7.Text = Format(_Importe, "$##,##0.00")
        TextBox5.Text = _PrecioUnitario
        TextBox7.Text = _Importe
        TextBox4.Text = Frm28VentaFT.TextBox15.Text
        Label8.Text = Frm28VentaFT.ComboBox1.Text
        TextBox9.Text = Frm28VentaFT.TextBox13.Text
        TextBox10.Text = Frm28VentaFT.TextBox14.Text
        TextBox12.Text = Frm28VentaFT.TextBox3.Text
        TextBox13.Text = Frm28VentaFT.TextBox12.Text
        CheckBox1.Checked = True
        ToolStripTextBox1.Text = Frm28VentaFT.TextBox1.Text
        ToolStripTextBox1.Enabled = False
        ToolStripTextBox2.Text = Frm28VentaFT.TextBox19.Text

        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox5.Enabled = False
        TextBox7.Enabled = False
        TextBox9.Enabled = False
        TextBox10.Enabled = False
        TextBox12.Enabled = False
        TextBox13.Enabled = False

    End Sub

    Private Sub ToolStripButton4_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripButton4.Click

        ' Guardar el Recibo
        Dim Checacontrol As Control
        Dim _Find As Boolean
        Dim title As String

        Dim response As MsgBoxResult
        _Find = False
        For Each Checacontrol In Me.Controls
            If TypeOf Checacontrol Is TextBox Then
                If Checacontrol.Text = "" Then
                    _Find = True
                End If
            End If
        Next

        If _Find Then
            Msg = "Hay casillas en blanco. Debe proporcionar todos los datos"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

        If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False Then
            Msg = "Debe seleccionar una opcion entre PRÓRROGA, DONACIÓN O RECETAS"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)
            Exit Sub
        End If

 
    
        Msg = "Está seguro de Guardar este Recibo Provisional?"
        style = MsgBoxStyle.DefaultButton2 Or _
           MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Entraga de Factor de TRansferencia"

        response = MsgBox(msg, style, title)
        If response = MsgBoxResult.Yes Then
            ' Guarda la Información

            _Busca = CServiciosDataSet40.Recibos.NewRow
            _Busca.Item("Recibo") = TextBox1.Text
            _Busca.Item("Fecha") = CDate(TextBox2.Text)
            _Busca.Item("NumeroFrascos") = TextBox3.Text
            _Busca.Item("PrecioUnitario") = TextBox5.Text
            _Busca.Item("Importe") = TextBox7.Text
            _Busca.Item("Comprador") = TextBox4.Text
            _Busca.Item("Paciente") = TextBox12.Text
            _Busca.Item("Cliente") = ToolStripTextBox2.Text
            _Busca.Item("TelefonoRecetas") = TextBox11.Text
            _Busca.Item("Prorroga") = RadioButton1.Checked
            _Busca.Item("Donacion") = RadioButton2.Checked
            _Busca.Item("Recetas") = RadioButton3.Checked
            _Busca.Item("Comentarios") = UCase(RichTextBox1.Text)
            _Busca.Item("Referencia") = TextBox8.Text
            _Busca.Item("FechaRegistro") = Today
            _Busca.Item("HoraRegistro") = Mid(TimeOfDay, 1, 5)
            _Busca.Item("Folio") = ToolStripTextBox1.Text
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

            CServiciosDataSet40.Recibos.Rows.Add(_Busca)

            Me.Validate()
            RecibosBindingSource1.EndEdit()
            Me.TableAdapterManager2.UpdateAll(CServiciosDataSet40)

            ' Incrementa el Consecutivo
            _Busca = CServiciosDataSet3.Documentos.FindByDocumento("REP")
            _Busca.Item("Consecutivo") = _Busca.Item("Consecutivo") + 1

            Me.Validate()
            DocumentosBindingSource.EndEdit()
            Me.TableAdapterManager.UpdateAll(CServiciosDataSet3)


            Msg = "Registro guardado Correctamente"
            Style = MsgBoxStyle.Information
            MsgBox(Msg, Style)

            If CheckBox1.Checked = True Then
                Button2_Click(sender, e)
                Msg = "Repote enviado a la Impresora"
                MsgBox(Msg, Style)
            End If
            Me.Close()

        Else
            Exit Sub
        End If

    End Sub

    Private Sub TextBox6_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox6.GotFocus
        TextBox6.Text = "N/A"
    End Sub

    Private Sub TextBox6_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox6.TextChanged

    End Sub

    Private Sub RichTextBox1_GotFocus(sender As Object, e As System.EventArgs) Handles RichTextBox1.GotFocus

    End Sub

    Private Sub RichTextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub TextBox8_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox8.GotFocus

    End Sub

    Private Sub TextBox8_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub TextBox11_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox11.GotFocus
        TextBox11.Text = "N/A"
    End Sub

    Private Sub TextBox11_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox11.TextChanged

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        ' Imprime el Recibo

        Dim m_Excel As Microsoft.Office.Interop.Excel.Application
        Dim strRutaExcel As String = "Z:\Recibo Provisional.xlsx"
        Dim _Recibo As String = TextBox1.Text
        Dim _Concepto, _Referencia, _Comprador As String
        Dim _Fecha As Date = CDate(TextBox2.Text)
        Dim _Cantidad, _PrecioUnitario, _Importe, _Mide, _Renglon As Integer
        Dim _Comentarios As String = RichTextBox1.Text
        Dim _Linea As String

        _Concepto = ""
        If RadioButton1.Checked = True Then
            _Concepto = "PRÓRROGA"
        End If
        If RadioButton2.Checked = True Then
            _Concepto = "DONACIÓN"
        End If
        If RadioButton3.Checked = True Then
            _Concepto = "RECETAS"
        End If

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
        m_Excel.Worksheets("RECIBO").cells(17, 7).value = Frm28VentaFT.ToolStripComboBox2.Text

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



    Private Sub TextBox14_GotFocus(sender As Object, e As System.EventArgs) Handles TextBox14.GotFocus
        TextBox14.Text = "N/A"
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged

    End Sub
End Class