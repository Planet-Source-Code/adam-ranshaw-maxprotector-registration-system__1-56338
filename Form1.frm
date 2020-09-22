VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H001449FB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Software - MaxProtector"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   10695
   ControlBox      =   0   'False
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Buy Game"
      Height          =   855
      Left            =   6600
      TabIndex        =   26
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Try Game"
      Default         =   -1  'True
      Height          =   855
      Left            =   8520
      TabIndex        =   25
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton exit1 
      Caption         =   "Exit"
      Height          =   855
      Left            =   8520
      TabIndex        =   24
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Timer Timer11 
      Interval        =   10
      Left            =   2400
      Top             =   0
   End
   Begin VB.Timer Timer10 
      Interval        =   1
      Left            =   1920
      Top             =   480
   End
   Begin VB.Timer Timer9 
      Interval        =   1
      Left            =   1920
      Top             =   0
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   480
   End
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
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
      Appearance      =   0
      Max             =   30
      Scrolling       =   1
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
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   0
      TabIndex        =   23
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Label22 
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
      Height          =   1215
      Left            =   360
      TabIndex        =   22
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait while loading remaining time..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   3960
      Width           =   5895
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2520
      TabIndex        =   20
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   4
      Height          =   1695
      Left            =   2280
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADRANIX"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5640
      TabIndex        =   10
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Your free time is being counted down even now so click the 'Try Game' button to begin the game so that your time is not wasted."
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
      Height          =   1335
      Left            =   840
      TabIndex        =   18
      Top             =   5400
      Width           =   4935
   End
   Begin VB.Label Label17 
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
      TabIndex        =   17
      Top             =   5280
      Width           =   735
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
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   15
      Top             =   3600
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
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the 30 minute trial of Software"
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
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "of gameplay remaining"
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
   Begin VB.Label Label19 
      BackColor       =   &H00009EEA&
      Height          =   1695
      Left            =   2280
      TabIndex        =   19
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0442
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
      Height          =   1335
      Left            =   840
      TabIndex        =   13
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "With this 30 minute trial you can use all of the features in this game with no restrictions of play free for 30 minutes."
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
      Height          =   1095
      Left            =   840
      TabIndex        =   12
      Top             =   2400
      Width           =   5175
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
   Begin VB.Shape Shape2 
      BackColor       =   &H00F48A2E&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0014EFFB&
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


End Function





Private Sub Command1_Click()
Unload Form1
form3.Label1.Visible = True
about.Label6.Caption = "NO"

form3.unreg.Visible = False
form3.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Unload reg
reg.Visible = True
End Sub


Private Sub exit1_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
reg.sec.LoadFile "c:\windows\system32\adranixsec000.rtf"
reg.name1.LoadFile "c:\windows\system32\adranixname000.rtf"
reg.code.LoadFile "c:\windows\system32\adranixcode000.rtf"
reg.Text1.Text = reg.name1.Text
reg.Text2.Text = reg.code.Text
If reg.sec.Text = "111000" Then
Command1.Visible = False
'Registration  Code format
Dim i
Dim zip
Dim final
Dim code1 As Single
If reg.Text1.Text = "" Or reg.Text2.Text = "" Or reg.Text5.Text = "" Or reg.Text6.Text = "" Then
MsgBox "Security files have been changed. Please contact tech support at: help@aranix.co.uk", vbCritical
Command1.Visible = False
Exit Sub
End If


If Len(reg.Text1.Text) < 4 Then
MsgBox "Security files have been changed. Please contact tech support at: help@aranix.co.uk", vbCritical
Command1.Visible = False
    Exit Sub
End If

If reg.Text5.Text = ("8546854") And reg.Text6.Text = "64381" Then


Else
MsgBox "Security files have been changed. Please contact tech support at: help@aranix.co.uk", vbCritical
Command1.Visible = False
Exit Sub
End If


For i = 1 To Len(reg.Text1.Text) - 1
    code1 = Format(Asc(Right(reg.Text1.Text, Len(reg.Text1.Text) - i)) * 2 + (39 / i) + (i + 3 / 7), "#.#")
    zip = zip & code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    code1 = Format(Asc(Right(zip, Len(zip) - i)) * 0.5 + (1 / i) + (i + 1 / 7), "#00")
    final = final & code1
Next i
final = Right(final, Len(final) - 4)
final = final & Asc(reg.Text1)
'If reg code is correct
If reg.Text2.Text = final Then
'Enable License file Frame
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
reg.name1.Text = reg.Text1.Text
reg.code.Text = reg.Text2.Text
form3.regnow.Visible = False
Form1.Visible = False
form3.Visible = True
Else
MsgBox "Security files have been changed. Please contact tech support at: help@aranix.co.uk", vbCritical
Command1.Visible = False
End If
End If
End Sub



Private Sub Label1_Click()
    SaveSetting Me.Name, "Trial", "TimesOpen", 0
'Resets the trial
    Label1.Caption = 0
'Resets the Label
End Sub



Private Sub Label10_Click()
If Label23.Visible = False Then
SaveSetting Me.Name, "Trial", "TimesOpen", 1799
Label1.Caption = 1799
End
End If
End Sub











Private Sub Label15_Click()
End
End Sub





Private Sub Label23_Click()
Label23.Visible = False
End Sub

Private Sub Label8_Click()
If Label23.Visible = False Then
SaveSetting Me.Name, "Trial", "TimesOpen", 1739
Label1.Caption = 1739
End
End If
End Sub

Private Sub Label9_Click()
If Label23.Visible = False Then
SaveSetting Me.Name, "Trial", "TimesOpen", 0
Label1.Caption = 0
End
End If
End Sub

Private Sub Timer1_Timer()
TrialTime Me, "Your 30 minute trial is Expired.  To continue using software please register.", "Trial Expired", vbCritical, 1800, True
Label6.Caption = Label3.Caption
Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
End Sub

Private Sub Timer10_Timer()
If Label3.Caption = "" Then
'Hold the user
Else
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label11.Visible = False
Shape4.Visible = False
Form1.Enabled = True
Timer9.Enabled = False
Timer10.Enabled = False
End If
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
Label3.Caption = "<1 minute"
Else
If Label1.Caption > 1800 Then
Timer1.Interval = 1000
Label3.Caption = "E X P I R E D"
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
Command1.Enabled = True
End If
End Sub

Private Sub Timer4_Timer()
If Label1.Caption = 60 Then
time.Value = 1
Else
If Label1.Caption = 120 Then
time.Value = 2
Else
If Label1.Caption = 180 Then
time.Value = 3
Else
If Label1.Caption = 240 Then
time.Value = 4
Else
If Label1.Caption = 300 Then
time.Value = 5
Else
If Label1.Caption = 360 Then
time.Value = 6
Else
If Label1.Caption = 420 Then
time.Value = 7
Else
If Label1.Caption = 480 Then
time.Value = 8
Else
If Label1.Caption = 540 Then
time.Value = 9
Else
If Label1.Caption = 600 Then
time.Value = 10
Else
If Label1.Caption = 660 Then
time.Value = 11
Else
If Label1.Caption = 720 Then
time.Value = 12
Else
If Label1.Caption = 780 Then
time.Value = 13
Else
If Label1.Caption = 840 Then
time.Value = 14
Else
If Label1.Caption = 900 Then
time.Value = 15
Else
If Label1.Caption = 960 Then
time.Value = 16
Else
If Label1.Caption = 1020 Then
time.Value = 17
Else
If Label1.Caption = 1080 Then
time.Value = 18
Else
If Label1.Caption = 1140 Then
time.Value = 19
Else
If Label1.Caption = 1200 Then
time.Value = 20
Else
If Label1.Caption = 1260 Then
time.Value = 21
Else
If Label1.Caption = 1320 Then
time.Value = 22
Else
If Label1.Caption = 1380 Then
time.Value = 23
Else
If Label1.Caption = 1440 Then
time.Value = 24
Else
If Label1.Caption = 1500 Then
time.Value = 25
Else
If Label1.Caption = 1560 Then
time.Value = 26
Else
If Label1.Caption = 1620 Then
time.Value = 27
Else
If Label1.Caption = 1680 Then
time.Value = 28
Else
If Label1.Caption = 1740 Then
time.Value = 29
Else
If Label1.Caption > 1800 Then
time.Value = 30
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
Form10.Visible = True
Form1.Visible = False
reg.Visible = False
Command1.Enabled = False
Command1.Visible = False
exit1.Visible = True
Label3.Caption = "E X P I R E D"
time.Value = 30
Timer5.Enabled = False
End If
End Sub

Private Sub Timer6_Timer()
time2.Value = time.Value
End Sub




Private Sub Timer8_Timer()
Timer1.Enabled = True
End Sub

Private Sub Timer9_Timer()
If exit1.Visible = True Then
'Hold the user
Else
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label11.Visible = False
Shape4.Visible = False
Form1.Enabled = True
Timer9.Enabled = False
Timer10.Enabled = False
End If
End Sub
