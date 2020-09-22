VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form reg 
   AutoRedraw      =   -1  'True
   BackColor       =   &H001449FB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registraion"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   10695
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin RichTextLib.RichTextBox sec 
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"reg.frx":0000
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Text            =   "64381"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Text            =   "8546854"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H001449FB&
      Caption         =   "How to Buy"
      ForeColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"reg.frx":0082
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
   End
   Begin RichTextLib.RichTextBox name1 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"reg.frx":0308
   End
   Begin RichTextLib.RichTextBox code 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"reg.frx":038A
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H001449FB&
      Caption         =   "Registration Infomation"
      ForeColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   4440
      TabIndex        =   3
      Top             =   1320
      Width           =   6135
      Begin VB.CommandButton admin 
         Caption         =   "Admin"
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   4800
         Width           =   1815
      End
      Begin VB.CommandButton exit 
         Caption         =   "Exit"
         Height          =   615
         Left            =   2040
         TabIndex        =   20
         Top             =   4800
         Width           =   1935
      End
      Begin VB.CommandButton ok 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   615
         Left            =   4080
         TabIndex        =   19
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H001449FB&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H001449FB&
         Caption         =   "      Please change computer             name to over 3 characters"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   2535
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H001449FB&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3960
         Width           =   5775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2520
         Width           =   5775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Code:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Name:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"reg.frx":040C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Label labelpuk 
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   6000
      TabIndex        =   7
      Top             =   360
      Width           =   4575
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   1095
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4575
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
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Function TrialTime(TheForm As Form, TrialOverMSG As String, TrialOverMSGTitle As String, TrialOverMSGType As String, TrialCount As Integer, Work As Boolean)

    If Not Work Then SaveSetting TheForm.Name, "puk", "TimesOpen", "."
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting TheForm.Name, "puk", "TimesOpen", Val(GetSetting(TheForm.Name, "puk", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened

    If GetSetting(TheForm.Name, "puk", "TimesOpen") > TrialCount Then SaveSetting TheForm.Name, "puk", "TimesOpen", TrialCount: MsgBox TrialOverMSG, TrialOverMSGType, TrialOverMSGTitle: Timer1.Enabled = False
'If the amount of times open is > then the TrialCount..
'Reset it to the number in TrialCount specified
'Display a message and terminate the program
End Function


Private Sub admin_Click()
If Text1.Text = "ADAMS-PC" Then
Form4.Visible = False
Form4.Visible = True
Else
Form5.Visible = False
Form5.Visible = True
End If
End Sub

Private Sub exit_Click()
Unload reg
End Sub






Public Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
End Function




Private Sub Form_Load()
reg.labelpuk.Caption = GetSetting(Me.Name, "PUK", "TimesOpen")
Dim PCName As String
Dim P As Long
 P = NameOfPC(PCName)
 Text1.Text = PCName
If Len(reg.Text1.Text) < 4 Then
Check1.Visible = True
 End If
 End Sub

Private Sub ok_Click()
'Registration  Code format
Dim i
Dim zip
Dim final
Dim code1 As Single
If Text1.Text = "" Or Text2.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
    reg.Enabled = False
    Form9.Visible = True
Text2.Text = ""
Exit Sub
End If


If Len(Text1.Text) < 4 Then
   MsgBox "Please change your computer name to somthing over 3 letter/numbers.", vbExclamation
Text2.Text = ""
    Exit Sub
End If

If Text5.Text = ("8546854") And Text6.Text = "64381" Then


Else
    reg.Enabled = False
    Form9.Visible = True
Text2.Text = ""
Exit Sub
End If


For i = 1 To Len(Text1.Text) - 1
    code1 = Format(Asc(Right(Text1.Text, Len(Text1.Text) - i)) * 2 + (39 / i) + (i + 3 / 7), "#.#")
    zip = zip & code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    code1 = Format(Asc(Right(zip, Len(zip) - i)) * 0.5 + (1 / i) + (i + 1 / 7), "#00")
    final = final & code1
Next i
final = Right(final, Len(final) - 4)
final = final & Asc(Text1)
'If reg code is correct
If Text2.Text = final Then
'Enable License file Frame
name1.Text = Text1.Text
code.Text = Text2.Text
sec.Text = "111000"
sec.SaveFile "c:\windows\system32\adranixsec000.rtf"
name1.SaveFile "c:\windows\system32\adranixname000.rtf"
code.SaveFile "c:\windows\system32\adranixcode000.rtf"
MsgBox "Thank you for registering this product with ADRANIX. Click OK to close software then re-launch software to continue.", vbInformation + vbOKOnly, "Registered"
End
Else
TrialTime Me, "A PUK Code is needed to continue.", "PUK Code Needed", vbCritical, 5, True
labelpuk.Caption = GetSetting(Me.Name, "PUK", "TimesOpen")
    reg.Enabled = False
    Form9.Visible = True
Text2.Text = ""
End If

End Sub



Private Sub Text1_Click()
MsgBox "You can not change the text in this box.", vbCritical
End Sub

Private Sub Timer1_Timer()
If Label5.Caption = "E X P I R E D left" Then
Timer1.Enabled = False
Label5.Caption = "E X P I R E D"
Else
Label5.Caption = Form1.Label3.Caption + " left"
End If
End Sub

Private Sub Timer2_Timer()
Text1.Text = PCName
End Sub

Private Sub Timer3_Timer()
If labelpuk.Caption = "5" Then
Form10.Visible = True
Form1.Command1.Visible = False
Form10.Option1.Enabled = False
Form10.Option2.Enabled = False
Form10.Option4.Value = True
Form10.Option4.Enabled = True
Form10.Label1.Caption = "Software is Frozen"
Form10.Label2.Caption = "A PUK code is needed, please exit this software or enter the PUK."
Form1.Visible = False
reg.Visible = False
Form1.Timer5.Enabled = False
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
SaveSetting Me.Name, "puk", "TimesOpen", 0
reg.labelpuk.Caption = 0
End
End Sub

Private Sub Timer5_Timer()
If labelpuk.Caption = "5" Then
Form1.Visible = False
reg.Visible = False
End If
End Sub
