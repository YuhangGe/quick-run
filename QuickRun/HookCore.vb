Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Class HookCore
    Public Declare Auto Function GetKeyState Lib "user32" Alias "GetKeyState" (ByVal nVirtKey As Integer) As Short
    '�����ֵ
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
    ''' <param name="nCode">HC_ACTION��HC_NOREMOVE</param>
    ''' <param name="wParam">����Virtual Key</param>
    ''' <param name="lParam">��WM_KEYDOWNͬ</param>
    ''' <returns>��ѶϢҪ������0��֮��1</returns>
    Public Delegate Function HookProc(ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Shared hMouseHook As Integer = 0
    Shared hKeyboardHook As Integer = 0
    'Shared �ؼ���ָʾһ�������������ı��Ԫ�ؽ�������
    '����Ԫ�ز�������ĳ���ṹ���ض�ʵ����
    '����ͨ��ʹ��������ṹ���ƻ������ṹ���ض�ʵ���ı��������޶�����Ԫ�����������ǡ�
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
        Public vkCode As Integer '// ָ��һ��virtual-key�룬�����Ǵ�1��254��һ��ֵ�� 
        Public scanCode As Integer '// ΪKEYָ��һ��Ӳ��ɨ���롣 
        Public flags As Integer '// ָ��extended-key��־���¼�ע���־, �����Ĵ��룬��״̬ת����־��
        Public time As Integer '// Ϊ�����Ϣָ��ʱ���־��
        Public dwExtraInfo As Integer '// ָ������ĺ������Ϣ�й�������Ϣ�� 
    End Class
    ''' <summary>
    ''' ����������������Hook���á�
    ''' </summary>
    ''' <param name="idHook">Hook�����ͣ����������Ϣ���͡�</param>
    ''' <param name="Lpfn">Hook�ӳ̣���������̣��ĵ�ַָ�롣���dwThreadId����Ϊ0����һ���ɱ�Ľ��̴������̵߳ı�ʶ��lpfn����ָ��DLL�е�Hook�ӳ̡� �������⣬lpfn����ָ��ǰ���̵�һ��Hook�ӳ̴���</param>
    ''' <param name="hMod">��Ӧ�ó���ʵ���ľ������ʶ����lpfn��ָ���ӳ̵�DLL�����dwThreadId ��ʶ��ǰ���̴�����һ���̣߳������ӳ̴���λ�ڵ�ǰ���̣�hMod����ΪNULL��</param>
    ''' <param name="dwThreadId">�밲װHook�ӳ���������̵߳ı�ʶ�������Ϊ0��Hook�ӳ������е��̹߳����� </param>
    ''' <returns>�����ɹ��򷵻�Hook�ӳ̵ľ����ʧ�ܷ���NULL��</returns>
    Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As HookType, ByVal lpfn As HookProc, ByVal hmod As Integer, ByVal dwThreadId As Integer) As Integer
    ''' <summary>
    ''' �������ǽ��Hook֮�á�
    ''' </summary>
    ''' <param name="hHook">Hook�����ľ����</param>
    ''' <returns>�����ɹ��򷵻�TRUE��ʧ�ܷ���FALSE��</returns>
    Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Integer) As Integer
    ''' <summary>
    ''' �������������ǽ���ǰHook���е�Hook��Ϣ���ݸ���һ��Hook��
    ''' </summary>
    ''' <param name="hHook">��ǰHook�ľ����һ��Ӧ�ó����������������Ϊ��ǰ����SetWindowsHookEx�����Ľ����</param>
    ''' <param name="nCode">���ݵ���ǰHook���̵�Hook���룬��һ��Hook����ʹ����δ���ȥ������δ���Hook��Ϣ��</param>
    ''' <param name="wParam">���ݸ���ǰHook���̵�wParamֵ�����ľ��庬�����ɵ�ǰHook���е����Hook�����;����ġ�</param>
    ''' <param name="lParam">���ݸ���ǰHook���̵�lParamֵ�����ľ��庬�����ɵ�ǰHook���е����Hook�����;����ġ�</param>
    Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Integer, ByVal ncode As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    ''' <summary>
    ''' �����������ڻ�ȡ��ǰ�߳�һ��Ψһ���̱߳�ʶ����
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
            'AndAlso �����,���������ʽִ�м��߼���ȡ.
            '��� expression1 Ϊ ���� expression2 Ϊ result ��ֵΪ 
            'True True True 
            'True False False 
            'False �������㣩 False 
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
