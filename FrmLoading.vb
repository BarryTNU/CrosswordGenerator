Public Class FrmLoading
    Inherits Form
    Public Sub UpdateProgress(percent As Integer)

        If ProgressBar1.InvokeRequired Then
            ProgressBar1.Invoke(Sub() ProgressBar1.Value = percent)
        Else
            ProgressBar1.Value = percent
        End If

    End Sub


End Class