<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm57NuevosPacientes
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm57NuevosPacientes))
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ComboBox4 = New System.Windows.Forms.ComboBox()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.TextBox30 = New System.Windows.Forms.TextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.TextBox24 = New System.Windows.Forms.TextBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox11 = New System.Windows.Forms.TextBox()
        Me.TextBox9 = New System.Windows.Forms.TextBox()
        Me.TextBox8 = New System.Windows.Forms.TextBox()
        Me.TextBox7 = New System.Windows.Forms.TextBox()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.CServiciosDataSet3 = New Cservicios.CServiciosDataSet3()
        Me.DocumentosBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DocumentosTableAdapter = New Cservicios.CServiciosDataSet3TableAdapters.DocumentosTableAdapter()
        Me.TableAdapterManager = New Cservicios.CServiciosDataSet3TableAdapters.TableAdapterManager()
        Me.CServiciosDataSet2 = New Cservicios.CServiciosDataSet2()
        Me.PacientesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.PacientesTableAdapter = New Cservicios.CServiciosDataSet2TableAdapters.PacientesTableAdapter()
        Me.TableAdapterManager1 = New Cservicios.CServiciosDataSet2TableAdapters.TableAdapterManager()
        Me.DiagnosticosTableAdapter = New Cservicios.CServiciosDataSet2TableAdapters.DiagnosticosTableAdapter()
        Me.CServiciosDataSet = New Cservicios.CServiciosDataSet()
        Me.CentrosBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.CentrosTableAdapter = New Cservicios.CServiciosDataSetTableAdapters.CentrosTableAdapter()
        Me.TableAdapterManager2 = New Cservicios.CServiciosDataSetTableAdapters.TableAdapterManager()
        Me.DiagnosticosBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        CType(Me.CServiciosDataSet3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DocumentosBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CServiciosDataSet2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PacientesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CServiciosDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CentrosBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DiagnosticosBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Nothing
        Me.BindingNavigator1.BackColor = System.Drawing.Color.DodgerBlue
        Me.BindingNavigator1.CountItem = Nothing
        Me.BindingNavigator1.DeleteItem = Nothing
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2})
        Me.BindingNavigator1.Location = New System.Drawing.Point(0, 0)
        Me.BindingNavigator1.MoveFirstItem = Nothing
        Me.BindingNavigator1.MoveLastItem = Nothing
        Me.BindingNavigator1.MoveNextItem = Nothing
        Me.BindingNavigator1.MovePreviousItem = Nothing
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Nothing
        Me.BindingNavigator1.Size = New System.Drawing.Size(526, 25)
        Me.BindingNavigator1.TabIndex = 53
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
        Me.ToolStripButton2.ToolTipText = "Guardar Datos" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.ComboBox4)
        Me.Panel1.Controls.Add(Me.Label32)
        Me.Panel1.Controls.Add(Me.Label31)
        Me.Panel1.Controls.Add(Me.TextBox30)
        Me.Panel1.Controls.Add(Me.Label23)
        Me.Panel1.Controls.Add(Me.TextBox24)
        Me.Panel1.Controls.Add(Me.ComboBox2)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.ComboBox1)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.TextBox11)
        Me.Panel1.Controls.Add(Me.TextBox9)
        Me.Panel1.Controls.Add(Me.TextBox8)
        Me.Panel1.Controls.Add(Me.TextBox7)
        Me.Panel1.Controls.Add(Me.TextBox6)
        Me.Panel1.Controls.Add(Me.TextBox5)
        Me.Panel1.Controls.Add(Me.TextBox4)
        Me.Panel1.Controls.Add(Me.TextBox3)
        Me.Panel1.Controls.Add(Me.TextBox2)
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 25)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(526, 290)
        Me.Panel1.TabIndex = 54
        '
        'ComboBox4
        '
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Location = New System.Drawing.Point(336, 10)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(123, 21)
        Me.ComboBox4.TabIndex = 1
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(253, 223)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(32, 13)
        Me.Label32.TabIndex = 28
        Me.Label32.Text = "Edad"
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(33, 223)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(60, 13)
        Me.Label31.TabIndex = 27
        Me.Label31.Text = "Fecha Nac"
        '
        'TextBox30
        '
        Me.TextBox30.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox30.Location = New System.Drawing.Point(117, 220)
        Me.TextBox30.Name = "TextBox30"
        Me.TextBox30.Size = New System.Drawing.Size(102, 20)
        Me.TextBox30.TabIndex = 10
        '
        'Label23
        '
        Me.Label23.AutoSize = True
        Me.Label23.Location = New System.Drawing.Point(303, 13)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(31, 13)
        Me.Label23.TabIndex = 25
        Me.Label23.Text = "Sexo"
        '
        'TextBox24
        '
        Me.TextBox24.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox24.Location = New System.Drawing.Point(465, 11)
        Me.TextBox24.Name = "TextBox24"
        Me.TextBox24.Size = New System.Drawing.Size(25, 20)
        Me.TextBox24.TabIndex = 1
        Me.TextBox24.TabStop = False
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(117, 193)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(265, 21)
        Me.ComboBox2.TabIndex = 9
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(33, 193)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(63, 13)
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "Diagnóstico"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(295, 168)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(73, 13)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "Fecha de Alta"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(33, 168)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(75, 13)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Fecha Ingreso"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(33, 142)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(59, 13)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "Referencia"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(33, 115)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(38, 13)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Centro"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(117, 112)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(265, 21)
        Me.ComboBox1.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(33, 89)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 13)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Nombres"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(33, 63)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 13)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Materno"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(33, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Paterno"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(33, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(49, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Paciente"
        '
        'TextBox11
        '
        Me.TextBox11.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox11.Enabled = False
        Me.TextBox11.Location = New System.Drawing.Point(388, 194)
        Me.TextBox11.Name = "TextBox11"
        Me.TextBox11.Size = New System.Drawing.Size(102, 20)
        Me.TextBox11.TabIndex = 10
        '
        'TextBox9
        '
        Me.TextBox9.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox9.Location = New System.Drawing.Point(388, 165)
        Me.TextBox9.Name = "TextBox9"
        Me.TextBox9.Size = New System.Drawing.Size(102, 20)
        Me.TextBox9.TabIndex = 8
        '
        'TextBox8
        '
        Me.TextBox8.Location = New System.Drawing.Point(238, 10)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(59, 20)
        Me.TextBox8.TabIndex = 7
        Me.TextBox8.Visible = False
        '
        'TextBox7
        '
        Me.TextBox7.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox7.Location = New System.Drawing.Point(117, 165)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(102, 20)
        Me.TextBox7.TabIndex = 7
        '
        'TextBox6
        '
        Me.TextBox6.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox6.Location = New System.Drawing.Point(117, 139)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(373, 20)
        Me.TextBox6.TabIndex = 6
        '
        'TextBox5
        '
        Me.TextBox5.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox5.Enabled = False
        Me.TextBox5.Location = New System.Drawing.Point(388, 112)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(102, 20)
        Me.TextBox5.TabIndex = 4
        Me.TextBox5.TabStop = False
        '
        'TextBox4
        '
        Me.TextBox4.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox4.Location = New System.Drawing.Point(117, 86)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(373, 20)
        Me.TextBox4.TabIndex = 4
        '
        'TextBox3
        '
        Me.TextBox3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox3.Location = New System.Drawing.Point(117, 60)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(373, 20)
        Me.TextBox3.TabIndex = 3
        '
        'TextBox2
        '
        Me.TextBox2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox2.Location = New System.Drawing.Point(117, 34)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(373, 20)
        Me.TextBox2.TabIndex = 2
        '
        'TextBox1
        '
        Me.TextBox1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(117, 8)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(102, 22)
        Me.TextBox1.TabIndex = 0
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
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.DocumentosTableAdapter = Me.DocumentosTableAdapter
        Me.TableAdapterManager.UpdateOrder = Cservicios.CServiciosDataSet3TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'CServiciosDataSet2
        '
        Me.CServiciosDataSet2.DataSetName = "CServiciosDataSet2"
        Me.CServiciosDataSet2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'PacientesBindingSource
        '
        Me.PacientesBindingSource.DataMember = "Pacientes"
        Me.PacientesBindingSource.DataSource = Me.CServiciosDataSet2
        '
        'PacientesTableAdapter
        '
        Me.PacientesTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager1
        '
        Me.TableAdapterManager1.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager1.ConsultasTableAdapter = Nothing
        Me.TableAdapterManager1.DetalleRecetaTableAdapter = Nothing
        Me.TableAdapterManager1.DiagnosticosTableAdapter = Me.DiagnosticosTableAdapter
        Me.TableAdapterManager1.EvetosPacienteTableAdapter = Nothing
        Me.TableAdapterManager1.PacientesTableAdapter = Me.PacientesTableAdapter
        Me.TableAdapterManager1.RecetasTableAdapter = Nothing
        Me.TableAdapterManager1.UpdateOrder = Cservicios.CServiciosDataSet2TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'DiagnosticosTableAdapter
        '
        Me.DiagnosticosTableAdapter.ClearBeforeFill = True
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
        'TableAdapterManager2
        '
        Me.TableAdapterManager2.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager2.CentrosTableAdapter = Me.CentrosTableAdapter
        Me.TableAdapterManager2.UpdateOrder = Cservicios.CServiciosDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'DiagnosticosBindingSource
        '
        Me.DiagnosticosBindingSource.DataMember = "Diagnosticos"
        Me.DiagnosticosBindingSource.DataSource = Me.CServiciosDataSet2
        '
        'Frm57NuevosPacientes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(526, 315)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.BindingNavigator1)
        Me.Name = "Frm57NuevosPacientes"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Registro de Nuevos Pacientes"
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.CServiciosDataSet3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DocumentosBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CServiciosDataSet2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PacientesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CServiciosDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CentrosBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DiagnosticosBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BindingNavigator1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ComboBox4 As System.Windows.Forms.ComboBox
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents TextBox30 As System.Windows.Forms.TextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents TextBox24 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox11 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents CServiciosDataSet3 As Cservicios.CServiciosDataSet3
    Friend WithEvents DocumentosBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DocumentosTableAdapter As Cservicios.CServiciosDataSet3TableAdapters.DocumentosTableAdapter
    Friend WithEvents TableAdapterManager As Cservicios.CServiciosDataSet3TableAdapters.TableAdapterManager
    Friend WithEvents CServiciosDataSet2 As Cservicios.CServiciosDataSet2
    Friend WithEvents PacientesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents PacientesTableAdapter As Cservicios.CServiciosDataSet2TableAdapters.PacientesTableAdapter
    Friend WithEvents TableAdapterManager1 As Cservicios.CServiciosDataSet2TableAdapters.TableAdapterManager
    Friend WithEvents CServiciosDataSet As Cservicios.CServiciosDataSet
    Friend WithEvents CentrosBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents CentrosTableAdapter As Cservicios.CServiciosDataSetTableAdapters.CentrosTableAdapter
    Friend WithEvents TableAdapterManager2 As Cservicios.CServiciosDataSetTableAdapters.TableAdapterManager
    Friend WithEvents DiagnosticosTableAdapter As Cservicios.CServiciosDataSet2TableAdapters.DiagnosticosTableAdapter
    Friend WithEvents DiagnosticosBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
