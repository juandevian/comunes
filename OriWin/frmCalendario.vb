Friend Class FrmCalendario
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub BttCancelClick(sender As System.Object, e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub FrmCalendario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Width = MncOrionPlus.Width + 30
        Me.Height = MncOrionPlus.Height + 60
    End Sub
End Class