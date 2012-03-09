Public Class Tray
    Public Sub ShowCmdWindow()
        mw.Show()
        mw.Activate()
        buffer.Clear()
    End Sub
    Public Sub RunCommand()
        mw.Run()
    End Sub
    Public Sub OnKeyDown(ByVal keyCode As Key)
        Dim needFresh As Boolean = True


        Select Case keyCode
            Case 65 To 90, 49 To 57
                buffer.Append(ChrW(keyCode))
                ' MDebug.Print("append " + c + " cur " + buffer.ToString)
            Case 8
                If buffer.Length > 0 Then
                    buffer.Remove(buffer.Length - 1, 1)
                End If
            Case Else
                needFresh = False
        End Select
        If needFresh Then
            mw.RefreshCmd(buffer.ToString.ToLower)
        End If
    End Sub
    Public Sub HideCmdWindow()
        mw.Clear()
        mw.Hide()

    End Sub

    Private Sub m_exit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles m_exit.Click, NotifyIcon1.DoubleClick
#If Not Debug Then
        keyHook.Stops()
#End If
        System.Windows.Application.Current.Shutdown()
    End Sub

  
End Class
