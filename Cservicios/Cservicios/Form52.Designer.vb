<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm52Bitacoras
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm52Bitacoras))
        Me.BindingNavigator1 = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.BitácorasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UsoDeBañoMaríaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeProducciónToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeAlmacénToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeEntregaDeProducciónToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeUsoDeProcesaodrUltrasónicoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RegistroDeTemperaturasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RequisicionesDeMaterialToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MonthCalendar1 = New System.Windows.Forms.MonthCalendar()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.CServiciosDataSet32 = New Cservicios.CServiciosDataSet32()
        Me.VentaFactorBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.VentaFactorTableAdapter = New Cservicios.CServiciosDataSet32TableAdapters.VentaFactorTableAdapter()
        Me.TableAdapterManager = New Cservicios.CServiciosDataSet32TableAdapters.TableAdapterManager()
        Me.CServiciosDataSet26 = New Cservicios.CServiciosDataSet26()
        Me.ProduccionBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.ProduccionTableAdapter = New Cservicios.CServiciosDataSet26TableAdapters.ProduccionTableAdapter()
        Me.TableAdapterManager1 = New Cservicios.CServiciosDataSet26TableAdapters.TableAdapterManager()
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BindingNavigator1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.CServiciosDataSet32, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.VentaFactorBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CServiciosDataSet26, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ProduccionBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BindingNavigator1
        '
        Me.BindingNavigator1.AddNewItem = Nothing
        Me.BindingNavigator1.BackColor = System.Drawing.Color.DodgerBlue
        Me.BindingNavigator1.CountItem = Nothing
        Me.BindingNavigator1.DeleteItem = Nothing
        Me.BindingNavigator1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton1, Me.ToolStripButton2})
        Me.BindingNavigator1.Location = New System.Drawing.Point(0, 24)
        Me.BindingNavigator1.MoveFirstItem = Nothing
        Me.BindingNavigator1.MoveLastItem = Nothing
        Me.BindingNavigator1.MoveNextItem = Nothing
        Me.BindingNavigator1.MovePreviousItem = Nothing
        Me.BindingNavigator1.Name = "BindingNavigator1"
        Me.BindingNavigator1.PositionItem = Nothing
        Me.BindingNavigator1.Size = New System.Drawing.Size(799, 25)
        Me.BindingNavigator1.TabIndex = 1
        Me.BindingNavigator1.Text = "BindingNavigator1"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(100, 89)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(144, 33)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Uso del Baño María"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(250, 89)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(144, 33)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Registro de Temperaturas"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(400, 89)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(144, 33)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "Limpieza Áreas Críticas"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(550, 89)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(144, 33)
        Me.Button4.TabIndex = 5
        Me.Button4.Text = "Registro de Producción"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(550, 128)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(144, 33)
        Me.Button5.TabIndex = 6
        Me.Button5.Text = "Registro de Almacén"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(100, 128)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(144, 33)
        Me.Button6.TabIndex = 7
        Me.Button6.Text = "Registro de Entrega"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(250, 130)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(144, 33)
        Me.Button7.TabIndex = 8
        Me.Button7.Text = "Uso del Sonificador"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(400, 130)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(144, 33)
        Me.Button8.TabIndex = 9
        Me.Button8.Text = "Pedidos de Material"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BitácorasToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(799, 24)
        Me.MenuStrip1.TabIndex = 10
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'BitácorasToolStripMenuItem
        '
        Me.BitácorasToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UsoDeBañoMaríaToolStripMenuItem, Me.RegistroDeProducciónToolStripMenuItem, Me.RegistroDeAlmacénToolStripMenuItem, Me.RegistroDeEntregaDeProducciónToolStripMenuItem, Me.RegistroDeUsoDeProcesaodrUltrasónicoToolStripMenuItem, Me.RegistroDeTemperaturasToolStripMenuItem, Me.RequisicionesDeMaterialToolStripMenuItem})
        Me.BitácorasToolStripMenuItem.Name = "BitácorasToolStripMenuItem"
        Me.BitácorasToolStripMenuItem.Size = New System.Drawing.Size(145, 20)
        Me.BitácorasToolStripMenuItem.Text = "Impresión de Bitáccoras"
        '
        'UsoDeBañoMaríaToolStripMenuItem
        '
        Me.UsoDeBañoMaríaToolStripMenuItem.Name = "UsoDeBañoMaríaToolStripMenuItem"
        Me.UsoDeBañoMaríaToolStripMenuItem.Size = New System.Drawing.Size(297, 22)
        Me.UsoDeBañoMaríaToolStripMenuItem.Text = "Uso de Baño María"
        '
        'RegistroDeProducciónToolStripMenuItem
        '
        Me.RegistroDeProducciónToolStripMenuItem.Name = "RegistroDeProducciónToolStripMenuItem"
        Me.RegistroDeProducciónToolStripMenuItem.Size = New System.Drawing.Size(297, 22)
        Me.RegistroDeProducciónToolStripMenuItem.Text = "Registro de Producción"
        '
        'RegistroDeAlmacénToolStripMenuItem
        '
        Me.RegistroDeAlmacénToolStripMenuItem.Name = "RegistroDeAlmacénToolStripMenuItem"
        Me.RegistroDeAlmacénToolStripMenuItem.Size = New System.Drawing.Size(297, 22)
        Me.RegistroDeAlmacénToolStripMenuItem.Text = "Registro de Almacén"
        '
        'RegistroDeEntregaDeProducciónToolStripMenuItem
        '
        Me.RegistroDeEntregaDeProducciónToolStripMenuItem.Name = "RegistroDeEntregaDeProducciónToolStripMenuItem"
        Me.RegistroDeEntregaDeProducciónToolStripMenuItem.Size = New System.Drawing.Size(297, 22)
        Me.RegistroDeEntregaDeProducciónToolStripMenuItem.Text = "Registro de Entrega de Producción"
        '
        'RegistroDeUsoDeProcesaodrUltrasónicoToolStripMenuItem
        '
        Me.RegistroDeUsoDeProcesaodrUltrasónicoToolStripMenuItem.Name = "RegistroDeUsoDeProcesaodrUltrasónicoToolStripMenuItem"
        Me.RegistroDeUsoDeProcesaodrUltrasónicoToolStripMenuItem.Size = New System.Drawing.Size(297, 22)
        Me.RegistroDeUsoDeProcesaodrUltrasónicoToolStripMenuItem.Text = "Registro de Uso de Procesaodr Ultrasónico"
        '
        'RegistroDeTemperaturasToolStripMenuItem
        '
        Me.RegistroDeTemperaturasToolStripMenuItem.Name = "RegistroDeTemperaturasToolStripMenuItem"
        Me.RegistroDeTemperaturasToolStripMenuItem.Size = New System.Drawing.Size(297, 22)
        Me.RegistroDeTemperaturasToolStripMenuItem.Text = "Registro de Temperaturas"
        '
        'RequisicionesDeMaterialToolStripMenuItem
        '
        Me.RequisicionesDeMaterialToolStripMenuItem.Name = "RequisicionesDeMaterialToolStripMenuItem"
        Me.RequisicionesDeMaterialToolStripMenuItem.Size = New System.Drawing.Size(297, 22)
        Me.RequisicionesDeMaterialToolStripMenuItem.Text = "Requisiciones de Material"
        '
        'MonthCalendar1
        '
        Me.MonthCalendar1.Location = New System.Drawing.Point(499, 62)
        Me.MonthCalendar1.Name = "MonthCalendar1"
        Me.MonthCalendar1.TabIndex = 11
        Me.MonthCalendar1.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(187, 59)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(89, 20)
        Me.TextBox1.TabIndex = 12
        Me.TextBox1.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(105, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(67, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Fecha Inicial"
        Me.Label1.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(304, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Fecha Final"
        Me.Label2.Visible = False
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(386, 59)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(89, 20)
        Me.TextBox2.TabIndex = 14
        Me.TextBox2.Visible = False
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(250, 169)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(89, 20)
        Me.TextBox3.TabIndex = 16
        Me.TextBox3.Visible = False
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(10, 62)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(48, 20)
        Me.TextBox4.TabIndex = 17
        Me.TextBox4.Visible = False
        '
        'CServiciosDataSet32
        '
        Me.CServiciosDataSet32.DataSetName = "CServiciosDataSet32"
        Me.CServiciosDataSet32.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'VentaFactorBindingSource
        '
        Me.VentaFactorBindingSource.DataMember = "VentaFactor"
        Me.VentaFactorBindingSource.DataSource = Me.CServiciosDataSet32
        '
        'VentaFactorTableAdapter
        '
        Me.VentaFactorTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.UpdateOrder = Cservicios.CServiciosDataSet32TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        Me.TableAdapterManager.VentaFactorTableAdapter = Me.VentaFactorTableAdapter
        '
        'CServiciosDataSet26
        '
        Me.CServiciosDataSet26.DataSetName = "CServiciosDataSet26"
        Me.CServiciosDataSet26.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'ProduccionBindingSource
        '
        Me.ProduccionBindingSource.DataMember = "Produccion"
        Me.ProduccionBindingSource.DataSource = Me.CServiciosDataSet26
        '
        'ProduccionTableAdapter
        '
        Me.ProduccionTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager1
        '
        Me.TableAdapterManager1.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager1.ProduccionTableAdapter = Me.ProduccionTableAdapter
        Me.TableAdapterManager1.UpdateOrder = Cservicios.CServiciosDataSet26TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "ToolStripButton1"
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
        'Frm52Bitacoras
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(799, 241)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.MonthCalendar1)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.BindingNavigator1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Frm52Bitacoras"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Generación de Bitácoras "
        CType(Me.BindingNavigator1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BindingNavigator1.ResumeLayout(False)
        Me.BindingNavigator1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.CServiciosDataSet32, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.VentaFactorBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CServiciosDataSet26, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ProduccionBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BindingNavigator1 As System.Windows.Forms.BindingNavigator
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents BitácorasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UsoDeBañoMaríaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeProducciónToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeAlmacénToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeEntregaDeProducciónToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeUsoDeProcesaodrUltrasónicoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RegistroDeTemperaturasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RequisicionesDeMaterialToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MonthCalendar1 As System.Windows.Forms.MonthCalendar
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents CServiciosDataSet32 As Cservicios.CServiciosDataSet32
    Friend WithEvents VentaFactorBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents VentaFactorTableAdapter As Cservicios.CServiciosDataSet32TableAdapters.VentaFactorTableAdapter
    Friend WithEvents TableAdapterManager As Cservicios.CServiciosDataSet32TableAdapters.TableAdapterManager
    Friend WithEvents CServiciosDataSet26 As Cservicios.CServiciosDataSet26
    Friend WithEvents ProduccionBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents ProduccionTableAdapter As Cservicios.CServiciosDataSet26TableAdapters.ProduccionTableAdapter
    Friend WithEvents TableAdapterManager1 As Cservicios.CServiciosDataSet26TableAdapters.TableAdapterManager
End Class
