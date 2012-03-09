Public Class Reminder

    Private Sub StackPanel_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs)
        If e.ClickCount >= 2 Then
            Me.Close()
        End If
    End Sub
    Private Shared FFamily As String = "微软雅黑"
    Private Shared FSize As Double = 35
    Private Shared FColor As Brush = Brushes.Red

    Public Sub ShowRemind(ByVal rlist As List(Of String()))
        Dim height As Double = 100
        Dim max_width As Double = 300
        For Each rs As String() In rlist
            Dim TxtMsg As TextBlock = New TextBlock
            TxtMsg.Text = rs(0)
            TxtMsg.FontFamily = New FontFamily("微软雅黑")
            TxtMsg.FontSize = FSize
            TxtMsg.TextAlignment = TextAlignment.Center
            TxtMsg.Foreground = FColor

            Dim TxtRemind As TextBlock = New TextBlock
            TxtRemind.Text = rs(1)
            TxtRemind.FontFamily = New FontFamily("微软雅黑")
            TxtRemind.FontSize = FSize
            TxtRemind.TextAlignment = TextAlignment.Center
            TxtRemind.Foreground = FColor

            MainStackPanel.Children.Add(TxtMsg)
            MainStackPanel.Children.Add(TxtRemind)

            Dim ft As FormattedText = New FormattedText(rs(0), System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, _
                New Typeface(FFamily), FSize, FColor)
            If max_width < ft.Width + 20 Then
                max_width = ft.Width + 20
            End If
            ft = New FormattedText(rs(1), System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, _
                New Typeface(FFamily), FSize, FColor)
            If max_width < ft.Width + 20 Then
                max_width = ft.Width + 20
            End If
            height += 80
        Next
        Me.Height = height
        Me.Width = max_width
        Me.Show()
        Me.Activate()
        Me.Focus()
    End Sub
End Class
