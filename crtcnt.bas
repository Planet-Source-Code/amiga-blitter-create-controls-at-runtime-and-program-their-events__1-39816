Attribute VB_Name = "Module1"
'Programmed by Amiga Blitter
'Some Constants declaration
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const GWL_WNDPROC = (-4)
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_IME_KEYUP = &H291
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYLAST = &H108
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const WM_MOUSEMOVE = &H200
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
'The object carrier
Public strnew As Variant
'Handle for the procedures
Global oldWndProc As Long
Global oldWndProc1 As Long
'Api
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nindex As Long, ByVal dwnewlong As Long) As Long
'Global Obj As Variant

Public Function newWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Long) As Long
    If uMsg = WM_LBUTTONDOWN Then
    ' do this when MouseButtonDown on the object
        Form1.Text1.Text = "ok " & CStr(hWnd)
        Beep
        Form1.Text1.Text = "You hear a beep???"
    ElseIf uMsg = WM_MOUSEMOVE Then
    'And this when MouseMove on the Object
        Form1.Text1.Text = "MouseMove " & CStr(hWnd)
    End If
        newWndProc = CallWindowProc(oldWndProc, hWnd, uMsg, wParam, lParam)
End Function

Public Function newWndProc1(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Long) As Long
    If uMsg = WM_LBUTTONDOWN Then
    'Beep when Mousedown events occur
        Beep
        Form1.Text1.Text = Form1.Text1.Text + "beep"
    ElseIf uMsg = WM_MOUSEMOVE Then
    'print this when Mousedown events occur
        Form1.Text1.Text = "MouseMove " & CStr(hWnd)
    End If
    
        newWndProc1 = CallWindowProc(oldWndProc1, hWnd, uMsg, wParam, lParam)
End Function

Public Sub CreateObj(ByVal obj As Variant, ObjName As String, ObjType As String)
'Another little example
Set obj = Form1.Controls.Add("VB.CommandButton", ObjName)
'Debug.Print strnew.hWnd
obj.Top = 100
obj.Left = 2000
obj.Caption = "cico2"
obj.Visible = True

End Sub
