Imports System.Globalization
Module Module1

    Sub Main()

        Console.WriteLine(CreateRemind(5, 1, 20, True))
        Console.ReadKey()
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="month">月份</param>
    ''' <param name="day">日期</param>
    ''' <param name="lunar">是否是农历</param>
    ''' <param name="days">提前多少天提醒</param>
    ''' <returns>返回提醒的内容</returns>
    ''' <remarks></remarks>
    Function CreateRemind(ByVal month As Integer, ByVal day As Integer, Optional ByVal days As Integer = 3, Optional ByVal lunar As Boolean = False) As String
        Dim gap As Integer = -1
        If lunar = False Then '如果是公历生日
            gap = GetSoliarGap(month, day, days)
        Else '如果是农历生日
            gap = GetLunarGap(month, day, days)
        End If
        If gap <> -1 Then
            Select Case gap
                Case 0
                    Return "今天"
                Case 1
                    Return "明天"
                Case 2
                    Return "后天"
                Case Else
                    Return gap.ToString() + "天后"
            End Select
        End If
        Return Nothing
    End Function
    Function GetSoliarGap(ByVal month As Integer, ByVal day As Integer, ByVal days As Integer) As Integer
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
    Function GetLunarGap(ByVal month As Integer, ByVal day As Integer, ByVal days As Integer) As Integer

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
End Module
