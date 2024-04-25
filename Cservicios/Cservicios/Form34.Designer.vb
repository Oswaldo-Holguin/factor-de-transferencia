<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm34SubMenuCxC
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ClienteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FacturasToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CobranzaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContrarecibosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(9, 25)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 31)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Clientes"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(154, 25)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 31)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Facturas"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(295, 25)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(112, 31)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "Cobranza"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(434, 25)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(112, 31)
        Me.Button4.TabIndex = 3
        Me.Button4.Text = "Contrarecibos"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClienteToolStripMenuItem, Me.FacturasToolStripMenuItem, Me.CobranzaToolStripMenuItem, Me.ContrarecibosToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(760, 24)
        Me.MenuStrip1.TabIndex = 4
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ClienteToolStripMenuItem
        '
        Me.ClienteToolStripMenuItem.Name = "ClienteToolStripMenuItem"
        Me.ClienteToolStripMenuItem.Size = New System.Drawing.Size(56, 20)
        Me.ClienteToolStripMenuItem.Text = "Cliente"
        '
        'FacturasToolStripMenuItem
        '
        Me.FacturasToolStripMenuItem.Name = "FacturasToolStripMenuItem"
        Me.FacturasToolStripMenuItem.Size = New System.Drawing.Size(63, 20)
        Me.FacturasToolStripMenuItem.Text = "Facturas"
        '
        'CobranzaToolStripMenuItem
        '
        Me.CobranzaToolStripMenuItem.Name = "CobranzaToolStripMenuItem"
        Me.CobranzaToolStripMenuItem.Size = New System.Drawing.Size(69, 20)
        Me.CobranzaToolStripMenuItem.Text = "Cobranza"
        '
        'ContrarecibosToolStripMenuItem
        '
        Me.ContrarecibosToolStripMenuItem.Name = "ContrarecibosToolStripMenuItem"
        Me.ContrarecibosToolStripMenuItem.Size = New System.Drawing.Size(93, 20)
        Me.ContrarecibosToolStripMenuItem.Text = "Contrarecibos"
        '
        'Frm34SubMenuCxC
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(760, 68)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Frm34SubMenuCxC"
        Me.Text = "Menú Cuentras por Cobrar"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ClienteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FacturasToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CobranzaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ContrarecibosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
