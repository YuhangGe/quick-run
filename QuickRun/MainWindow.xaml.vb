Imports System.Globalization

Class MainWindow

    Private CmdTable As Dictionary(Of String, CmdStruc)
    Public RemindList As List(Of String()) = New List(Of String())()

    Public Structure CmdStruc
        Public Path As String
        Public NeedConfirm As Boolean
        Public Params As String
        Public Sub New(ByVal path As String, ByVal need As Boolean, ByVal param As String)
            Me.Path = path
            Me.NeedConfirm = need
            Me.Params = param
        End Sub
    End Structure
    Public Sub RefreshCmd(ByVal pre As String)
        'MDebug.Print("refresh " & pre)
        TxtInput.Text = pre
        If Not String.IsNullOrWhiteSpace(pre) Then
            AddToCanvas(GetSimilarities(pre))
        Else
            MainCanvas.Children.Clear()
        End If

    End Sub
    Public Sub Clear()
        TxtInput.Text = ""
        MainCanvas.Children.Clear()
        TxtInput.InvalidateVisual()
        MainCanvas.InvalidateVisual()
    End Sub
    Private Sub AddToCanvas(ByRef commands As List(Of String))
        MainCanvas.Children.Clear()
        Dim i_end As Integer = commands.Count - 1
        For i As Integer = 0 To i_end
            Dim cmd As String = commands(i)
            Dim t As New TextBlock
            t.Text = cmd
            t.FontSize = 40
            If i = 0 Then
                t.Foreground = Brushes.Orange
            Else
                t.Foreground = Brushes.BlueViolet
            End If
            t.Margin = New Thickness(10, i * 50, 0, 0)
            MainCanvas.Children.Add(t)
        Next

    End Sub

    Private Function GetSimilarities(ByVal pre As String) As List(Of String)
        Dim rtn As New List(Of String)
        For Each cmd In CmdTable.Keys
            If cmd.ToLower.StartsWith(pre) Then
                rtn.Add(cmd)
            End If
        Next
        Return rtn
    End Function
    Public Sub Run()
        If MainCanvas.Children.Count = 0 Then
            Exit Sub
        End If
        Dim cur As TextBlock = MainCanvas.Children(0)
        Dim cmd As CmdStruc = CmdTable(cur.Text)
        If (cmd.NeedConfirm) Then
            If MessageBox.Show("Sure to run " + cur.Text + "?", "QuickRun", MessageBoxButton.YesNo, MessageBoxImage.Question) <> MessageBoxResult.Yes Then
                Exit Sub
            End If
        End If
        If cmd.Params IsNot Nothing Then
            Process.Start(cmd.Path, cmd.Params)
        Else
            Process.Start(cmd.Path)
        End If
        TxtInput.Text = 0
        MainCanvas.Children.Clear()
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        CmdTable = New Dictionary(Of String, CmdStruc)
        Dim xdoc As XDocument = Nothing
        Try
            xdoc = XDocument.Load(System.Windows.Forms.Application.StartupPath + "\data.xml")
        Catch ex As Exception
            MsgBox("读取数据文件出错！", MsgBoxStyle.Information)
            System.Windows.Application.Current.Shutdown()
        End Try

        Dim elems As IEnumerable(Of XElement) = xdoc...<cmd>

        For Each cmd As XElement In elems
            Dim path As String = cmd.@path
            Dim name As String = cmd.@name
            If (String.IsNullOrWhiteSpace(name) OrElse String.IsNullOrWhiteSpace(path)) Then
                Continue For
            End If
            Dim need As Boolean = False
            Try
                need = CBool(cmd.@needConfirm)
            Catch ex As Exception
                need = False
            End Try
            Dim param As String = cmd.@param
            If String.IsNullOrWhiteSpace(param) Then
                param = Nothing
            End If
            CmdTable.Add(name, New CmdStruc(path, need, param))
        Next


        Dim items As IEnumerable(Of XElement) = xdoc...<rem>
        For Each r As XElement In items
            Dim _msg As String = r.@msg
            Dim _date As String = r.@date
            Dim _days As String = r.@days
            Dim _remind As String = r.@remind
            Dim _nongli As String = r.@nongli
            Dim days As Integer
            Dim lunar As Boolean
            If String.IsNullOrWhiteSpace(_msg) OrElse String.IsNullOrWhiteSpace(_date) Then
                Continue For
            End If
            If String.IsNullOrWhiteSpace(_days) Then
                days = 3
            Else
                Try
                    days = Integer.Parse(_days)
                Catch ex As Exception
                    days = 3
                End Try
            End If
            If String.IsNullOrWhiteSpace(_nongli) Then
                lunar = False
            Else
                Try
                    lunar = Boolean.Parse(_nongli)
                Catch ex As Exception
                    lunar = False
                End Try
            End If
            If String.IsNullOrWhiteSpace(_remind) Then
                _remind = "可不要忘记了!"
            End If
            Dim cr As String = CreateRemind(_date, days, lunar)
            If cr <> Nothing Then
                Dim tmp(1) As String
                tmp(0) = cr & _msg
                tmp(1) = _remind
                RemindList.Add(tmp)
            End If
        Next

        Me.Focus()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Function CreateRemind(ByVal _date As String, ByVal days As Integer, ByVal lunar As Boolean) As String
        Dim month As Integer = Integer.Parse(_date.Split("-")(0))
        Dim day As Integer = Integer.Parse(_date.Split("-")(1))

        Dim gap As Integer = -1
        If lunar = False Then '如果是公历生日
            gap = GetSoliarGap(month, day, days)
        Else '如果是农历生日
            gap = GetLunarGap(month, day, days)
        End If
        If gap <> -1 Then
            Select Case gap
                Case 0
                    Return "今天是"
                Case 1
                    Return "明天是"
                Case 2
                    Return "后天是"
                Case Else
                    Return gap.ToString() + "天后是"
            End Select
        End If
        Return Nothing
    End Function
 
    Private Function GetSoliarGap(ByVal month As Integer, ByVal day As Integer, ByVal days As Integer) As Integer
        Dim d As DateTime = DateTime.Now
        For i As Integer = 0 To days - 1
            d = DateTime.Now.AddDays(i)
            Dim _m As Integer = d.Month
            Dim _d As Integer = d.Day
            If _m = month AndAlso _d = day Then
                Return i
            End If
        Next
        Return -1
    End Function
    Private Function GetLunarGap(ByVal month As Integer, ByVal day As Integer, ByVal days As Integer) As Integer

        Dim clc As ChineseLunisolarCalendar = New ChineseLunisolarCalendar
        Dim d As DateTime = DateTime.Now
        For i As Integer = 0 To days - 1
            d = DateTime.Now.AddDays(i)
            Dim lyear As Integer = clc.GetYear(d)
            Dim lmonth As Integer = clc.GetMonth(d)
            Dim lday As Integer = clc.GetDayOfMonth(d)
            Dim lr As Integer = clc.GetLeapMonth(lyear)
            If lr <> 0 AndAlso lmonth > lr Then
                lmonth -= 1
            End If
            If lmonth = month AndAlso lday = day Then
                Return i
            End If
        Next
        Return -1
    End Function



    Private Sub textBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs)
        If e.Key = Key.Enter Then
            If MainCanvas.Children.Count = 0 Then
                Exit Sub
            End If
            Dim cur As TextBlock = MainCanvas.Children(0)
            Dim cmd As CmdStruc = CmdTable(cur.Text)
            If (cmd.NeedConfirm) Then
                If MessageBox.Show("Sure to run " + cur.Text + "?", "QuickRun", MessageBoxButton.YesNo, MessageBoxImage.Question) <> MessageBoxResult.Yes Then
                    Exit Sub
                End If

            End If
            If cmd.Params IsNot Nothing Then
                Process.Start(cmd.Path, cmd.Params)
            Else
                Process.Start(cmd.Path)
            End If

        End If
    End Sub


End Class
