VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Twisted Text Example"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":000C
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By Randy Porosky
'Twisted Text is Kinda like Scrambled Text
Function TwistText(Text As String)
Dim CurPos As Integer
Dim endstr As String
CurPos = 1
Start:
endstr$ = endstr$ & Mid$(Text$, CurPos + 1, 1) & Mid$(Text$, CurPos, 1)
CurPos = CurPos + 2
DoEvents
If DoCancel = True Then Exit Function
If Len(Text$) > CurPos Then
GoTo Start
ElseIf Len(Text$) = CurPos Then
endstr$ = endstr$ & Mid$(Text$, Len(Text$), 1)
End If
TwistText = endstr$
End Function
Private Sub Command1_Click()
Text1.Text = TwistText(Text1.Text)
End Sub
