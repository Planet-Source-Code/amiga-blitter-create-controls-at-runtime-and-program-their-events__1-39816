VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Create Controls at Runtime"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Create Obj Func"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton CreateObjectCB 
      Caption         =   "Create Object"
      Height          =   555
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton ChangeObject 
      Caption         =   "ChangeObj"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton End 
      Caption         =   "End"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   4500
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strnew As Object ' the new object
Dim NameNew As String ' The new name of the object
Dim obj As Object


Private Sub ChangeObject_Click()

If NameNew <> "" Then
    'remove the object created before
    Form1.Controls.Remove NameNew
    'dereference
    Set strnew = Nothing
    'Set a new name for the object
    NameNew = "BLU"
Else
    Exit Sub
End If
'create the new object
Set strnew = Form1.Controls.Add("VB.Picturebox", NameNew)
'just print the object handle
Debug.Print strnew.hWnd

'play with some object property
strnew.Top = 100
strnew.Left = 1500
'Warning change the path of the bitmap
strnew.Picture = LoadPicture("d:\windows\winnt256.bmp")
strnew.AutoSize = True
strnew.Visible = True
'assign to it a new procedure (newWndProc1) with my more than one event
oldWndProc1 = SetWindowLong(strnew.hWnd, GWL_WNDPROC, AddressOf newWndProc1)

End Sub

Private Sub Command1_Click()
'Set strnew = Form1.Controls.Item("billo")
'strnew.Top = 1000
Call CreateObj("hallo", "CICO", "COMMANDBUTTON")
End Sub



Private Sub CreateObjectCB_Click()
'Check for the existence of a named object
If NameNew <> "" Then
    'remove if present and set hes reference to nothing
    Form1.Controls.Remove NameNew
    Set strnew = Nothing
Else
    Set strnew = Nothing
End If

NameNew = InputBox("Type the name of the control you want to create" + Chr(13) + "This will be the real name of the control", "Control Name")
'create a new picturebox with the name i choose before
If NameNew = "" Then Exit Sub
Set strnew = Form1.Controls.Add("VB.Picturebox", NameNew)

Debug.Print strnew.hWnd
'setup the picturebox object
strnew.Left = 2000
strnew.Top = 2000
strnew.Height = 1375
strnew.Width = 1135
strnew.Visible = True
'strnew.Caption = "Test"

'assign to it a procedure (sub newWndProc) and programs events
oldWndProc = SetWindowLong(strnew.hWnd, GWL_WNDPROC, AddressOf newWndProc)
Text1.Text = "Mousemove and click on the new object created to test the programmed events"
End Sub

Private Sub End_Click()
End
End Sub

Public Sub Form_Load()
'Programmed by Amiga Blitter
'Comment and suggestion are welcome
'Please vote for me so i put some other code on this site
End Sub


Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Text = "Mousemove and click on the new object created to test the programmed events"
End Sub
