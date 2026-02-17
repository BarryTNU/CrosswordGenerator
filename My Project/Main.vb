Imports System.Threading

Module Main

    Private appMutex As Mutex

    <STAThread>
    Sub Main()

        Dim createdNew As Boolean

        appMutex = New Mutex(True, "NewCodeWord_SingleInstance", createdNew)

        If Not createdNew Then
            MessageBox.Show(
                "NewCodeWord is already running.",
                "Already running",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            )
            Return
        End If

        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New Form1())

        appMutex.ReleaseMutex()
    End Sub

End Module
