Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class HookCore
    Public Declare Auto Function GetKeyState Lib "user32" Alias "GetKeyState" (ByVal nVirtKey As Integer) As Short
    '虚拟键值
    Public Const VK_CONTROL = &H11
    Public Const VK_SHIFT = &H10
    Public Const VK_MENU = &H12
    Protected Overrides Sub Finalize()
        Stops()
    End Sub
    Private parent As Tray
    Public Sub New(ByVal parent As Tray)
        Me.parent = parent
    End Sub

    ''' <summary>
    ''' Hook Function
    ''' </summary>
    ''' <param name="nCode">HC_ACTION或HC_NOREMOVE</param>
    ''' <param name="wParam">表按键Virtual Key</param>
    ''' <param name="lParam">与WM_KEYDOWN同</param>
    ''' <returns>若讯息要被处理传0反之传1</returns>
    Public Delegate Function HookProc(ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Shared hMouseHook As Integer = 0
    Shared hKeyboardHook As Integer = 0
    'Shared 关键字指示一个或多个被声明的编程元素将被共享。
    '共享元素不关联于某类或结构的特定实例。
    '可以通过使用类名或结构名称或者类或结构的特定实例的变量名称限定共享元素来访问它们。
    Public Const WH_MOUSE_LL As Integer = 14
    Public Const WH_KEYBOARD_LL As Integer = 13
    Private MouseHookProcedure As HookProc
    Private KeyboardHookProcedure As HookProc
    '<StructLayout(LayoutKind.Sequential)> _
    'Public Class POINT
    '    Public x As Integer
    '    Public y As Integer
    'End Class
    <StructLayout(LayoutKind.Sequential)> _
    Public Class MouseHookStruct
        Public pt As Point
        Public hwnd As Integer
        Public wHitTestCode As Integer
        Public dwExtraInfo As Integer
    End Class
    <StructLayout(LayoutKind.Sequential)> _
    Public Class KeyboardHookStruct
        Public vkCode As Integer '// 指定一个virtual-key码，必须是从1到254的一个值。 
        Public scanCode As Integer '// 为KEY指定一个硬件扫描码。 
        Public flags As Integer '// 指定extended-key标志，事件注入标志, 上下文代码，和状态转换标志。
        Public time As Integer '// 为这个消息指定时间标志。
        Public dwExtraInfo As Integer '// 指定额外的和这个消息有关联的信息。 
    End Class
    ''' <summary>
    ''' 本函数是用于启动Hook设置。
    ''' </summary>
    ''' <param name="idHook">Hook的类型，即处理的消息类型。</param>
    ''' <param name="Lpfn">Hook子程（函数或过程）的地址指针。如果dwThreadId参数为0或是一个由别的进程创建的线程的标识，lpfn必须指向DLL中的Hook子程。 除此以外，lpfn可以指向当前进程的一段Hook子程代码</param>
    ''' <param name="hMod">是应用程序实例的句柄，标识包含lpfn所指的子程的DLL。如果dwThreadId 标识当前进程创建的一个线程，而且子程代码位于当前进程，hMod必须为NULL。</param>
    ''' <param name="dwThreadId">与安装Hook子程相关联的线程的标识符，如果为0，Hook子程与所有的线程关联。 </param>
    ''' <returns>函数成功则返回Hook子程的句柄，失败返回NULL。</returns>
    Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As HookType, ByVal lpfn As HookProc, ByVal hmod As Integer, ByVal dwThreadId As Integer) As Integer
    ''' <summary>
    ''' 本函数是解除Hook之用。
    ''' </summary>
    ''' <param name="hHook">Hook函数的句柄。</param>
    ''' <returns>函数成功则返回TRUE，失败返回FALSE。</returns>
    Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Integer) As Integer
    ''' <summary>
    ''' 本函数的作用是将当前Hook链中的Hook信息传递给下一个Hook。
    ''' </summary>
    ''' <param name="hHook">当前Hook的句柄，一个应用程序接收这个句柄，作为先前调用SetWindowsHookEx函数的结果。</param>
    ''' <param name="nCode">传递到当前Hook过程的Hook代码，下一个Hook过程使用这段代码去决定如何处理Hook信息。</param>
    ''' <param name="wParam">传递给当前Hook过程的wParam值，它的具体含义是由当前Hook链中的相关Hook的类型决定的。</param>
    ''' <param name="lParam">传递给当前Hook过程的lParam值，它的具体含义是由当前Hook链中的相关Hook的类型决定的。</param>
    Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Integer, ByVal ncode As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    ''' <summary>
    ''' 本函数是用于获取当前线程一个唯一的线程标识符。
    ''' </summary>
    Declare Function GetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Integer
    Public Enum HookType
        WH_JOURNALRECORD = 0
        WH_JOURNALPLAYBACK = 1
        WH_KEYBOARD = 2
        WH_GETMESSAGE = 3
        WH_CALLWNDPROC = 4
        WH_CBT = 5
        WH_SYSMSGFILTER = 6
        WH_MOUSE = 7
        WH_HARDWARE = 8
        WH_DEBUG = 9
        WH_SHELL = 10
        WH_FOREGROUNDIDLE = 11
        WH_CALLWNDPROCRET = 12
        WH_KEYBOARD_LL = 13
        WH_MOUSE_LL = 14
    End Enum
    Public Sub Start()

        If hKeyboardHook = 0 Then
            KeyboardHookProcedure = New HookProc(AddressOf KeyboardHookProc)
            hKeyboardHook = SetWindowsHookEx(HookType.WH_KEYBOARD_LL, KeyboardHookProcedure, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
            If hKeyboardHook = 0 Then
                Stop
                Throw New Exception("SetWindowsHookEx WH_KEYBOARD_LL failed.")
            End If
        End If
    End Sub

    Public Sub Stops()

        Dim retKeyboard As Boolean = True
        If Not (hKeyboardHook = 0) Then
            retKeyboard = UnhookWindowsHookEx(hKeyboardHook)
            hKeyboardHook = 0
        End If
        If Not retKeyboard Then
            'AndAlso 运算符,对两个表达式执行简化逻辑合取.
            '如果 expression1 为 并且 expression2 为 result 的值为 
            'True True True 
            'True False False 
            'False （不计算） False 
            Throw New Exception("UnhookWindowsHookEx failed.")
        End If
    End Sub

    Declare Function ToAscii Lib "user32" Alias "ToAscii" (ByVal uVirtKey As Integer, ByVal uScanCode As Integer, ByVal lpbKeyState As Byte(), ByVal lpwTransKey As Byte(), ByVal fuState As Integer) As Integer
    Declare Function GetKeyboardState Lib "user32" Alias "GetKeyboardState" (ByVal pbKeyState As Byte()) As Integer
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_KEYUP As Integer = &H101
    Private Const WM_SYSKEYDOWN As Integer = &H104
    Private Const WM_SYSKEYUP As Integer = &H105


    Public Property IsShow As Boolean = False

    Private Function KeyboardHookProc(ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
        If nCode >= 0 Then
            Dim ptrlParam As IntPtr = New IntPtr(lParam)
            Dim MyKeyboardHookStruct As KeyboardHookStruct = CType(Marshal.PtrToStructure(ptrlParam, GetType(KeyboardHookStruct)), KeyboardHookStruct)
            Dim _t As Boolean = True
            If wParam = WM_KEYDOWN OrElse wParam = WM_SYSKEYDOWN Then
                Dim keyData As Keys = CType(MyKeyboardHookStruct.vkCode, Keys)
                Dim e As System.Windows.Forms.KeyEventArgs = New System.Windows.Forms.KeyEventArgs(keyData)

                If IsShow = True Then
                    _t = False

                    If e.KeyCode = Keys.Enter Then
                        parent.RunCommand()
                        If IsShow = True Then
                            IsShow = False
                            parent.HideCmdWindow()
                        End If
                    ElseIf e.KeyCode = Keys.LMenu Then
                        _t = True
                    Else
                        parent.OnKeyDown(e.KeyCode)
                    End If

                End If
                'MDebug.Print(GetKeyState(VK_MENU).ToString)
                If GetKeyState(VK_MENU) < 0 AndAlso e.KeyCode = Keys.Space Then
                    _t = False
                    If IsShow = False Then
                        IsShow = True
                        parent.ShowCmdWindow()
                    Else
                        parent.HideCmdWindow()
                        IsShow = False
                    End If
                End If
           
            End If
            If (_t = True) Then
                Return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam)
            Else
                Return -1
            End If
        End If
        Return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam)
    End Function

End Class
