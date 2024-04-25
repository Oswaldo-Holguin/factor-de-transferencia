<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm40diagnósticos
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm40diagnósticos))
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton4 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton5 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton6 = New System.Windows.Forms.ToolStripButton()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.CServiciosDataSet2 = New Cservicios.CServiciosDataSet2()
        Me.DiagnosticosBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DiagnosticosTableAdapter = New Cservicios.CServiciosDataSet2TableAdapters.DiagnosticosTableAdapter()
        Me.TableAdapterManager = New Cservicios.CServiciosDataSet2TableAdapters.TableAdapterManager()
        Me.CServiciosDataSet = New Cservicios.CServiciosDataSet()
        Me.CentrosBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.CentrosTableAdapter = New Cservicios.CServiciosDataSetTableAdapters.CentrosTableAdapter()
        Me.TableAdapterManager1 = New Cservicios.CServiciosDataSetTableAdapters.TableAdapterManager()
        Me.CServiciosDataSet3 = New Cservicios.CServiciosDataSet3()
        Me.DocumentosBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DocumentosTableAdapter = New Cservicios.CServiciosDataSet3TableAdapters.DocumentosTableAdapter()
        Me.TableAdapterManager2 = New Cservicios.CServiciosDataSet3TableAdapters.TableAdapterManager()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.CServiciosDataSet2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DiagnosticosBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CServiciosDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CentrosBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CServiciosDataSet3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DocumentosBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Nothing
        Me.BindingNavigator1.BackColor = System.Drawing.Color.DodgerBlue
        Me.BindingNavigator1.CountItem = Nothing
        Me.BindingNavigator1.DeleteItem = Nothing
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2, Me.ToolStripButton3, Me.ToolStripButton4, Me.ToolStripButton5, Me.ToolStripButton6})
        Me.BindingNavigator1.Location = New System.Drawing.Point(0, 0)
        Me.BindingNavigator1.MoveFirstItem = Nothing
        Me.BindingNavigator1.MoveLastItem = Nothing
        Me.BindingNavigator1.MoveNextItem = Nothing
        Me.BindingNavigator1.MovePreviousItem = Nothing
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Nothing
        Me.BindingNavigator1.Size = New System.Drawing.Size(909, 25)
        Me.BindingNavigator1.TabIndex = 1
        Me.BindingNavigator1.Text = "BindingNavigator1"
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "ToolStripButton1"
        Me.ToolStripButton1.ToolTipText = "Salir" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton2.Text = "ToolStripButton2"
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton3.Image = CType(resources.GetObject("ToolStripButton3.Image"), System.Drawing.Image)
        Me.ToolStripButton3.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton3.Text = "ToolStripButton3"
        '
        'ToolStripButton4
        '
        Me.ToolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton4.Image = CType(resources.GetObject("ToolStripButton4.Image"), System.Drawing.Image)
        Me.ToolStripButton4.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton4.Name = "ToolStripButton4"
        Me.ToolStripButton4.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton4.Text = "ToolStripButton4"
        Me.ToolStripButton4.ToolTipText = "refrescar Información" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'ToolStripButton5
        '
        Me.ToolStripButton5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton5.Image = CType(resources.GetObject("ToolStripButton5.Image"), System.Drawing.Image)
        Me.ToolStripButton5.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton5.Name = "ToolStripButton5"
        Me.ToolStripButton5.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton5.Text = "ToolStripButton5"
        Me.ToolStripButton5.ToolTipText = "Modificar Datos de Personal Médico" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'ToolStripButton6
        '
        Me.ToolStripButton6.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton6.Image = Global.Cservicios.My.Resources.Resources.arrow_undo
        Me.ToolStripButton6.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton6.Name = "ToolStripButton6"
        Me.ToolStripButton6.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton6.Text = "ToolStripButton6"
        Me.ToolStripButton6.ToolTipText = "Deshacer" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Location = New System.Drawing.Point(462, 44)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(427, 193)
        Me.TabControl1.TabIndex = 3
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(419, 167)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Tabla de Diagnósticos"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(46, 59)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(103, 35)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Form16 Load"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 3)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(413, 161)
        Me.DataGridView1.TabIndex = 0
        '
        'Column1
        '
        Me.Column1.HeaderText = "Diagnóstico"
        Me.Column1.Name = "Column1"
        Me.Column1.Width = 60
        '
        'Column2
        '
        Me.Column2.HeaderText = "Descripción"
        Me.Column2.Name = "Column2"
        Me.Column2.Width = 300
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.TextBox5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.TextBox4)
        Me.Panel1.Controls.Add(Me.ComboBox1)
        Me.Panel1.Controls.Add(Me.TextBox3)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.TextBox2)
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Location = New System.Drawing.Point(12, 44)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(444, 193)
        Me.Panel1.TabIndex = 4
        '
        'TextBox5
        '
        Me.TextBox5.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox5.Location = New System.Drawing.Point(354, 126)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(49, 20)
        Me.TextBox5.TabIndex = 23
        Me.TextBox5.TabStop = False
        Me.TextBox5.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(10, 103)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(38, 13)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "Centro"
        '
        'TextBox4
        '
        Me.TextBox4.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox4.Enabled = False
        Me.TextBox4.Location = New System.Drawing.Point(354, 100)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(49, 20)
        Me.TextBox4.TabIndex = 21
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(81, 100)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(233, 21)
        Me.ComboBox1.TabIndex = 3
        '
        'TextBox3
        '
        Me.TextBox3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox3.Location = New System.Drawing.Point(81, 74)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(322, 20)
        Me.TextBox3.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 13)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "Referecia"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Descripción"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 13)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Diagnóstico"
        '
        'TextBox2
        '
        Me.TextBox2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox2.Location = New System.Drawing.Point(81, 48)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(322, 20)
        Me.TextBox2.TabIndex = 1
        '
        'TextBox1
        '
        Me.TextBox1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox1.Location = New System.Drawing.Point(81, 21)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(49, 20)
        Me.TextBox1.TabIndex = 0
        '
        'CServiciosDataSet2
        '
        Me.CServiciosDataSet2.DataSetName = "CServiciosDataSet2"
        Me.CServiciosDataSet2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'DiagnosticosBindingSource
        '
        Me.DiagnosticosBindingSource.DataMember = "Diagnosticos"
        Me.DiagnosticosBindingSource.DataSource = Me.CServiciosDataSet2
        '
        'DiagnosticosTableAdapter
        '
        Me.DiagnosticosTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.ConsultasTableAdapter = Nothing
        Me.TableAdapterManager.DetalleRecetaTableAdapter = Nothing
        Me.TableAdapterManager.DiagnosticosTableAdapter = Me.DiagnosticosTableAdapter
        Me.TableAdapterManager.EvetosPacienteTableAdapter = Nothing
        Me.TableAdapterManager.PacientesTableAdapter = Nothing
        Me.TableAdapterManager.RecetasTableAdapter = Nothing
        Me.TableAdapterManager.UpdateOrder = Cservicios.CServiciosDataSet2TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'CServiciosDataSet
        '
        Me.CServiciosDataSet.DataSetName = "CServiciosDataSet"
        Me.CServiciosDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'CentrosBindingSource
        '
        Me.CentrosBindingSource.DataMember = "Centros"
        Me.CentrosBindingSource.DataSource = Me.CServiciosDataSet
        '
        'CentrosTableAdapter
        '
        Me.CentrosTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager1
        '
        Me.TableAdapterManager1.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager1.CentrosTableAdapter = Me.CentrosTableAdapter
        Me.TableAdapterManager1.UpdateOrder = Cservicios.CServiciosDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'CServiciosDataSet3
        '
        Me.CServiciosDataSet3.DataSetName = "CServiciosDataSet3"
        Me.CServiciosDataSet3.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'DocumentosBindingSource
        '
        Me.DocumentosBindingSource.DataMember = "Documentos"
        Me.DocumentosBindingSource.DataSource = Me.CServiciosDataSet3
        '
        'DocumentosTableAdapter
        '
        Me.DocumentosTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager2
        '
        Me.TableAdapterManager2.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager2.DocumentosTableAdapter = Me.DocumentosTableAdapter
        Me.TableAdapterManager2.UpdateOrder = Cservicios.CServiciosDataSet3TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'Frm40diagnósticos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(909, 278)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.BindingNavigator1)
        Me.Name = "Frm40diagnósticos"
        Me.Text = "Mantenimiento a Tabla de Diagnósticos"
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.CServiciosDataSet2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DiagnosticosBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CServiciosDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CentrosBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CServiciosDataSet3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DocumentosBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BindingNavigator1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton3 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton4 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton5 As System.Windows.Forms.ToolStripButton
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Column1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Column2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents CServiciosDataSet2 As Cservicios.CServiciosDataSet2
    Friend WithEvents DiagnosticosBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DiagnosticosTableAdapter As Cservicios.CServiciosDataSet2TableAdapters.DiagnosticosTableAdapter
    Friend WithEvents TableAdapterManager As Cservicios.CServiciosDataSet2TableAdapters.TableAdapterManager
    Friend WithEvents CServiciosDataSet As Cservicios.CServiciosDataSet
    Friend WithEvents CentrosBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CentrosTableAdapter As Cservicios.CServiciosDataSetTableAdapters.CentrosTableAdapter
    Friend WithEvents TableAdapterManager1 As Cservicios.CServiciosDataSetTableAdapters.TableAdapterManager
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents CServiciosDataSet3 As Cservicios.CServiciosDataSet3
    Friend WithEvents DocumentosBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DocumentosTableAdapter As Cservicios.CServiciosDataSet3TableAdapters.DocumentosTableAdapter
    Friend WithEvents TableAdapterManager2 As Cservicios.CServiciosDataSet3TableAdapters.TableAdapterManager
    Friend WithEvents ToolStripButton6 As System.Windows.Forms.ToolStripButton
End Class
