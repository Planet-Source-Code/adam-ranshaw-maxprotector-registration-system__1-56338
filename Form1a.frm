VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00F48A2E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Racer"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   10695
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer7 
      Interval        =   10
      Left            =   1440
      Top             =   0
   End
   Begin Progetto1.XpBs command2 
      Height          =   855
      Left            =   6600
      TabIndex        =   18
      Top             =   5880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      Caption         =   "&Exit"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin Progetto1.XpBs command1 
      Height          =   855
      Left            =   8520
      TabIndex        =   17
      Top             =   5880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      Caption         =   "&Try Game"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
   End
   Begin VB.Timer Timer6 
      Interval        =   10
      Left            =   960
      Top             =   480
   End
   Begin MSComctlLib.ProgressBar time 
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   5160
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   84
   End
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   480
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   15
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1a.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   840
      TabIndex        =   13
      Top             =   4800
      Width           =   4935
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "With this 30 minute trial you can use all of the features in this game with no restrictions of play free for 30 minutes."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   840
      TabIndex        =   12
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the 30 minute trial of Super Racer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   6015
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADRANIX"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "* Instant Game Activation"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "* No CD requied to play"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "* Play any where any time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Own the full version for unlimited gameplay."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   6600
      TabIndex        =   6
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "...Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8520
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Demo time reaming:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   4
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E86400&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "of gameplay reaming"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "...Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6600
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Have"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X Minutes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0064E600&
      BackStyle       =   1  'Opaque
      Height          =   6735
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Function TrialTime(TheForm As Form, TrialOverMSG As String, TrialOverMSGTitle As String, TrialOverMSGType As String, TrialCount As Integer, Work As Boolean)

    If Not Work Then SaveSetting TheForm.Name, "Trial", "TimesOpen", "."
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting TheForm.Name, "Trial", "TimesOpen", Val(GetSetting(TheForm.Name, "Trial", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened

    If GetSetting(TheForm.Name, "Trial", "TimesOpen") > TrialCount Then SaveSetting TheForm.Name, "Trial", "TimesOpen", TrialCount: MsgBox TrialOverMSG, TrialOverMSGType, TrialOverMSGTitle: Timer1.Enabled = False
'If the amount of times open is > then the TrialCount..
'Reset it to the number in TrialCount specified
'Display a message and terminate the program
End Function


Private Sub command1_Click()
Timer7.Enabled = False
frmMain.Visible = True
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub form_load()
Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
End Sub


Private Sub Label1_Click()
    SaveSetting Me.Name, "Trial", "TimesOpen", 0
'Resets the trial
    Label1.Caption = 0
'Resets the Label
End Sub





Private Sub Label9_Click()
SaveSetting Me.Name, "Trial", "TimesOpen", 0
Label1.Caption = 0
End
End Sub

Private Sub Timer1_Timer()
TrialTime Me, "This software" & " has expired. Please register this product to get the full version.", "Trial Expired", vbCritical, 1800, True
Label6.Caption = Label3.Caption
Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
End Sub

Private Sub Timer2_Timer()
If Label1.Caption = 60 Then
Timer1.Interval = 1000
Label3.Caption = "29 minutes"
Else
If Label1.Caption = 120 Then
Timer1.Interval = 1000
Label3.Caption = "28 minutes"
Else
If Label1.Caption = 180 Then
Timer1.Interval = 1000
Label3.Caption = "27 minutes"
Else
If Label1.Caption = 240 Then
Timer1.Interval = 1000
Label3.Caption = "26 minutes"
Else
If Label1.Caption = 300 Then
Timer1.Interval = 1000
Label3.Caption = "25 minutes"
Else
If Label1.Caption = 360 Then
Timer1.Interval = 1000
Label3.Caption = "24 minutes"
Else
If Label1.Caption = 420 Then
Timer1.Interval = 1000
Label3.Caption = "23 minutes"
Else
If Label1.Caption = 480 Then
Timer1.Interval = 1000
Label3.Caption = "22 minutes"
Else
If Label1.Caption = 540 Then
Timer1.Interval = 1000
Label3.Caption = "21 minutes"
Else
If Label1.Caption = 600 Then
Timer1.Interval = 1000
Label3.Caption = "20 minutes"
Else
If Label1.Caption = 660 Then
Timer1.Interval = 1000
Label3.Caption = "19 minutes"
Else
If Label1.Caption = 720 Then
Timer1.Interval = 1000
Label3.Caption = "18 minutes"
Else
If Label1.Caption = 780 Then
Timer1.Interval = 1000
Label3.Caption = "17 minutes"
Else
If Label1.Caption = 840 Then
Timer1.Interval = 1000
Label3.Caption = "16 minutes"
Else
If Label1.Caption = 900 Then
Timer1.Interval = 1000
Label3.Caption = "15 minutes"
Else
If Label1.Caption = 960 Then
Timer1.Interval = 1000
Label3.Caption = "14 minutes"
Else
If Label1.Caption = 1020 Then
Timer1.Interval = 1000
Label3.Caption = "13 minutes"
Else
If Label1.Caption = 1080 Then
Timer1.Interval = 1000
Label3.Caption = "12 minutes"
Else
If Label1.Caption = 1140 Then
Timer1.Interval = 1000
Label3.Caption = "11 minutes"
Else
If Label1.Caption = 1200 Then
Timer1.Interval = 1000
Label3.Caption = "10 minutes"
Else
If Label1.Caption = 1260 Then
Timer1.Interval = 1000
Label3.Caption = "9 minutes"
Else
If Label1.Caption = 1320 Then
Timer1.Interval = 1000
Label3.Caption = "8 minutes"
Else
If Label1.Caption = 1380 Then
Timer1.Interval = 1000
Label3.Caption = "7 minutes"
Else
If Label1.Caption = 1440 Then
Timer1.Interval = 1000
Label3.Caption = "6 minutes"
Else
If Label1.Caption = 1500 Then
Timer1.Interval = 1000
Label3.Caption = "5 minutes"
Else
If Label1.Caption = 1560 Then
Timer1.Interval = 1000
Label3.Caption = "4 minutes"
Else
If Label1.Caption = 1620 Then
Timer1.Interval = 1000
Label3.Caption = "3 minutes"
Else
If Label1.Caption = 1680 Then
Timer1.Interval = 1000
Label3.Caption = "2 minutes"
Else
If Label1.Caption = 1740 Then
Timer1.Interval = 1000
Label3.Caption = "1 minute"
Else
If Label1.Caption > 1800 Then
Timer1.Interval = 1000
Label3.Caption = "<1 minute"
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer3_Timer()
If Timer1.Interval = 1000 Then
command1.Enabled = True
End If
End Sub

Private Sub Timer4_Timer()
If Label1.Caption = 60 Then
time.Value = 1
Else
If Label1.Caption = 120 Then
time.Value = 3
Else
If Label1.Caption = 180 Then
time.Value = 6
Else
If Label1.Caption = 240 Then
time.Value = 9
Else
If Label1.Caption = 300 Then
time.Value = 12
Else
If Label1.Caption = 360 Then
time.Value = 15
Else
If Label1.Caption = 420 Then
time.Value = 18
Else
If Label1.Caption = 480 Then
time.Value = 21
Else
If Label1.Caption = 540 Then
time.Value = 24
Else
If Label1.Caption = 600 Then
time.Value = 27
Else
If Label1.Caption = 660 Then
time.Value = 30
Else
If Label1.Caption = 720 Then
time.Value = 33
Else
If Label1.Caption = 780 Then
time.Value = 34
Else
If Label1.Caption = 840 Then
time.Value = 37
Else
If Label1.Caption = 900 Then
time.Value = 40
Else
If Label1.Caption = 960 Then
time.Value = 43
Else
If Label1.Caption = 1020 Then
time.Value = 46
Else
If Label1.Caption = 1080 Then
time.Value = 49
Else
If Label1.Caption = 1140 Then
time.Value = 52
Else
If Label1.Caption = 1200 Then
time.Value = 55
Else
If Label1.Caption = 1260 Then
time.Value = 58
Else
If Label1.Caption = 1320 Then
time.Value = 61
Else
If Label1.Caption = 1380 Then
time.Value = 64
Else
If Label1.Caption = 1440 Then
time.Value = 67
Else
If Label1.Caption = 1500 Then
time.Value = 70
Else
If Label1.Caption = 1560 Then
time.Value = 73
Else
If Label1.Caption = 1620 Then
time.Value = 76
Else
If Label1.Caption = 1680 Then
time.Value = 79
Else
If Label1.Caption = 1740 Then
time.Value = 81
Else
If Label1.Caption > 1800 Then
time.Value = 84
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer5_Timer()
If Label1.Caption > 1799 Then
command1.Enabled = False
command1.Visible = False
command2.Width = 2700
command2.Left = 7100
Label3.Caption = "E X P I R E D"
time.Value = 84
End If
End Sub

Private Sub Timer6_Timer()
If frmMain.Visible = True Then
If Label1.Caption > 1799 Then
Form1.Visible = True
frmMain.Visible = False
End If
End If
End Sub

Private Sub Timer7_Timer()
frmMain.Hide
End Sub
