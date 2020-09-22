VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim Ret As Integer

Select Case NewType
    Case Is = 1
        Ret = CreateCommandButton
    Case Is = 2
        Ret = CreateLabel
    Case Is = 3
        Ret = CreatePictureBox
    Case Is = 4
    Case Is = 5
End Select

'Reset The Object Type
    NewObject = 0
    
End Sub
Function CreateCommandButton()
'Determine Object Index
        ObjIndex = Me.CommandButton.ubound + 1
        Load Me.CommandButton(ObjIndex)
            'Set The Object Default Value and position
                Me.CommandButton(ObjIndex).Top = CurX
                Me.CommandButton(ObjIndex).Left = CurY
                Me.CommandButton(ObjIndex).Caption = "Button " + CStr(ObjIndex)
                Me.CommandButton(ObjIndex).Width = DefaultCommandButtonWidth
                Me.CommandButton(ObjIndex).Height = DefaultCommandButtonEight
                Me.CommandButton(ObjIndex).Visible = True
End Function

Function CreateLabel()
Dim ObjIndex As Long
'Determine Object Index
        ObjIndex = Me.Label.ubound + 1
        Load Me.Label1(ObjIndex)
            'Set The Object Default Value and position
                Me.Label(ObjIndex).Top = CurX
                Me.Label(ObjIndex).Left = CurY
                Me.Label(ObjIndex).Caption = "Label " + CStr(ObjIndex)
                Me.Label(ObjIndex).Width = DefaultLabelWidth
                Me.Label(ObjIndex).Height = DefaultLabelEight
                Me.Label(ObjIndex).Visible = True
End Function
Function CreatePictureBox()
Dim ObjIndex As Long
'Determine Object Index
        ObjIndex = Me.PictureBox.ubound + 1
        Load Me.PictureBox(ObjIndex)
            'Set The Object Default Value and position
                Me.Label1(PictureBox).Top = CurX
                Me.Label1(PictureBox).Left = CurY
                Me.Label1(PictureBox).Caption = "Label " + CStr(ObjIndex)
                Me.Label1(PictureBox).Width = DefaultPictureWidth
                Me.Label1(PictureBox).Height = DefaultPictureEight
                Me.Label1(PictureBox).Visible = True
End Function

Sub DeleteObject(Objname As Control)
'Delete Object
If MsgBox("Are you sure you want to delete the " + CStr(Objname) + "object", vbYesNo, "Deleting Object" = vbYes) Then
    object.Unload
Else
    MsgBox ("Operation Cancelled"), vbInformation, "Object Operation"
End If

End Sub

Sub FormSize(Frm As Form)

'Define form size
Select Case FormXSize
    Case Is = 640
        Frm.Height = 7320
        Frm.Width = 9720
        Frm.Top = -15
        Frm.Left = 1710
    Case Is = 800
        Frm.Height = 9120
        Frm.Width = 12120
        Frm.Top = -15
        Frm.Left = 1710
    Case Is = 1024
        Frm.Height = 11640
        Frm.Width = 15480
        Frm.Top = -15
        Frm.Left = 1710
End Select


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Load Form Attributes from database
'Startup form info - BackGround Picture or Slides, Midi etc

'Determine if RunMode or DesignMode
Dim FormQuery As Recordset
    Set FormQuery = dbase.openRecordset("Select * from FormStatrtUp Where
End Sub
