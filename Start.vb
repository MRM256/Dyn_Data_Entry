'Nice Start up routine from Active Query Builder
Imports System.Threading
Imports System.Windows.Forms

Friend NotInheritable Class Start
    Public strAppPath = Application.StartupPath

    Private Sub New()

    End Sub
    ''' <summary>
    ''' The main entry point for the application.
    ''' </summary>
    <STAThread> _
    Friend Shared Sub Main()
        ' Catch unhandled exceptions for debugging purposes
        AddHandler AppDomain.CurrentDomain.UnhandledException, _
            AddressOf CurrentDomain_UnhandledException
        AddHandler Application.ThreadException, _
            AddressOf Thread_UnhandledException

        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        'The following is were we start the application
        Application.Run(New frmMain())
    End Sub
    Private Shared Sub CurrentDomain_UnhandledException(sender As Object, _
                                                        e As UnhandledExceptionEventArgs)
        Dim exception As Exception = TryCast(e.ExceptionObject, Exception)
        If exception IsNot Nothing Then
            Dim exceptionDialog As New ThreadExceptionDialog(exception)
            If exceptionDialog.ShowDialog() = DialogResult.Abort Then
                Application.[Exit]()
            End If
        End If
    End Sub
    Private Shared Sub Thread_UnhandledException(sender As Object, _
                                                 e As ThreadExceptionEventArgs)
        If e.Exception IsNot Nothing Then
            Dim exceptionDialog As New ThreadExceptionDialog(e.Exception)
            If exceptionDialog.ShowDialog() = DialogResult.Abort Then
                Application.[Exit]()
            End If
        End If
    End Sub
End Class
