Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.
    Dim c As Tray
    'Dim r As Reminder
    Protected Overrides Sub OnStartup(ByVal e As System.Windows.StartupEventArgs)
        MyBase.OnStartup(e)
        'MDebug.Init()
        'r = New Reminder()
        ' r.Show()
        c = New Tray()

    End Sub
    Protected Overrides Sub OnExit(ByVal e As System.Windows.ExitEventArgs)
        MyBase.OnExit(e)
        c.Dispose()
        ' MDebug.Dispose()

    End Sub
End Class
