Public Class LoadingScreen
    Private Sub TimerLoading_Tick(sender As Object, e As EventArgs) Handles TimerLoading.Tick
        If (ProgressBar1.Value = ProgressBar1.Maximum) Then
            TimerLoading.Stop()
            Me.Hide()
            LoginPage.Show()
        Else
            ProgressBar1.PerformStep()
            Label1.Text = ProgressBar1.Value & ("%")
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TimerLoading.Start()
        Me.CenterToScreen()
    End Sub
End Class
