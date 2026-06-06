<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCalendario
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(disposing As Boolean)
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCalendario))
        Me.mncOrionPlus = New System.Windows.Forms.MonthCalendar()
        Me.SuspendLayout()
        '
        'mncOrionPlus
        '
        Me.mncOrionPlus.CalendarDimensions = New System.Drawing.Size(2, 2)
        Me.mncOrionPlus.Dock = System.Windows.Forms.DockStyle.Fill
        Me.mncOrionPlus.Location = New System.Drawing.Point(0, 0)
        Me.mncOrionPlus.Margin = New System.Windows.Forms.Padding(14)
        Me.mncOrionPlus.Name = "mncOrionPlus"
        Me.mncOrionPlus.ShowWeekNumbers = True
        Me.mncOrionPlus.TabIndex = 0
        '
        'frmCalendario
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(144.0!, 144.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(750, 482)
        Me.Controls.Add(Me.mncOrionPlus)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmCalendario"
        Me.ShowIcon = False
        Me.Text = My.Resources.Calendario
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents MncOrionPlus As System.Windows.Forms.MonthCalendar
End Class
