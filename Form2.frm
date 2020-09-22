VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11430
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   11430
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "X"
      Height          =   255
      Left            =   14880
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Down"
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton shoot 
      Caption         =   "UP"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Right"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Left"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox ship 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      Picture         =   "Form2.frx":1EA0BA
      ScaleHeight     =   2055
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   9120
      Width           =   2895
   End
   Begin MSComctlLib.ProgressBar mov 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   12300
   End
   Begin MSComctlLib.ProgressBar sho 
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
      Max             =   10000
   End
   Begin VB.PictureBox blu 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      Picture         =   "Form2.frx":1FD4C4
      ScaleHeight     =   2055
      ScaleWidth      =   3015
      TabIndex        =   4
      Top             =   6960
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Height          =   3375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
mov.Value = mov.Value + 100
End Sub



Private Sub Command2_Click()
On Error Resume Next
mov.Value = mov.Value - 100
End Sub



Private Sub Command3_Click()
On Error Resume Next
sho.Value = sho.Value + 100
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
sho.Value = 7920
End Sub

Private Sub shoot_Click()
On Error Resume Next
sho.Value = sho.Value - 100
End Sub

Private Sub Timer1_Timer()
ship.Left = mov.Value
blu.Left = mov.Value
blu.Top = sho.Value
End Sub
