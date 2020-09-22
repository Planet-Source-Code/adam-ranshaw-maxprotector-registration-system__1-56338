VERSION 5.00
Begin VB.Form form3 
   BackColor       =   &H00FF80FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5820
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   11040
   ControlBox      =   0   'False
   DrawWidth       =   80
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   5820
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox coll 
      Height          =   375
      Left            =   5760
      ScaleHeight     =   315
      ScaleWidth      =   2355
      TabIndex        =   36
      Top             =   4080
      Width           =   2415
   End
   Begin MaxProtector.XpBs XpBs1 
      Height          =   735
      Left            =   5760
      TabIndex        =   35
      Top             =   4560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      Caption         =   "Adranix Website"
      ButtonStyle     =   3
      OriginalPicSizeW=   0
      OriginalPicSizeH=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      URL             =   "www.adranix.co.uk"
   End
   Begin VB.PictureBox start 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      ScaleHeight     =   915
      ScaleWidth      =   1875
      TabIndex        =   26
      Top             =   4680
      Width           =   1935
   End
   Begin VB.PictureBox nudge3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      ScaleHeight     =   555
      ScaleWidth      =   2355
      TabIndex        =   21
      Top             =   3000
      Width           =   2415
   End
   Begin VB.PictureBox jack3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   6120
      Picture         =   "Form3.frx":144042
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   11
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox jack2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   3240
      Picture         =   "Form3.frx":15F3A6
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   10
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox jack1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   360
      Picture         =   "Form3.frx":17A70A
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   9
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox grape3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   6120
      Picture         =   "Form3.frx":195A6E
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox grape2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   3240
      Picture         =   "Form3.frx":197CB4
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox grape1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   360
      Picture         =   "Form3.frx":199EFA
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   6120
      Picture         =   "Form3.frx":19C140
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   5
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox apple2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   3240
      Picture         =   "Form3.frx":19D2DC
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox apple1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   360
      Picture         =   "Form3.frx":19E478
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox apple3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   6120
      Picture         =   "Form3.frx":19F614
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox cherry2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   3240
      Picture         =   "Form3.frx":1A3C44
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox cherry1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   360
      Picture         =   "Form3.frx":1A8274
      ScaleHeight     =   2415
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.PictureBox nudge2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      ScaleHeight     =   555
      ScaleWidth      =   2355
      TabIndex        =   22
      Top             =   3000
      Width           =   2415
   End
   Begin VB.PictureBox nudge1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      ScaleHeight     =   555
      ScaleWidth      =   2355
      TabIndex        =   23
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Shape Shape18 
      Height          =   1455
      Left            =   5640
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4680
      TabIndex        =   34
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Nudges:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   33
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4680
      TabIndex        =   32
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Jackpots:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   31
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4440
      TabIndex        =   30
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Won:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   29
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4440
      TabIndex        =   28
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Credits:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3360
      TabIndex        =   27
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Shape Shape17 
      Height          =   735
      Left            =   480
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Shape Shape15 
      Height          =   735
      Left            =   480
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Shape Shape14 
      BorderWidth     =   2
      Height          =   1935
      Left            =   3240
      Top             =   3720
      Width           =   5295
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "LOSE!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   480
      TabIndex        =   25
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "WIN!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   480
      TabIndex        =   24
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Shape Shape13 
      BorderWidth     =   2
      Height          =   1935
      Left            =   360
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Shape Shape12 
      Height          =   375
      Left            =   8760
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "£2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   20
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Shape Shape11 
      Height          =   375
      Left            =   8760
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "£1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   19
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Shape Shape10 
      Height          =   375
      Left            =   8760
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "£3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   18
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape Shape9 
      Height          =   375
      Left            =   8760
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "£5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   17
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Shape Shape8 
      Height          =   375
      Left            =   8760
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "£8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Shape Shape7 
      Height          =   375
      Left            =   8760
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "£10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      Height          =   375
      Left            =   8760
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "£15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Shape Shape5 
      Height          =   375
      Left            =   8760
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "£20"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   840
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      Height          =   375
      Left            =   8760
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Jackpot"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   360
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   9
      Height          =   2415
      Left            =   360
      Top             =   360
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   9
      Height          =   2415
      Left            =   6120
      Top             =   360
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   9
      Height          =   2415
      Left            =   3240
      Top             =   360
      Width           =   2415
   End
   Begin VB.Menu regmenu 
      Caption         =   "Registration"
      Begin VB.Menu regnow 
         Caption         =   "Register"
      End
      Begin VB.Menu unreg 
         Caption         =   "Unregister"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub cmdChoice1_Click(Index As Integer)
frmAcey.Show
End Sub

Private Sub cmdChoice2_Click(Index As Integer)
frmHam.Show
End Sub

Private Sub cmdChoice3_Click(Index As Integer)
frmEven.Show
End Sub

Private Sub cmdChoice4_Click(Index As Integer)
frmMemory.Show
End Sub

Private Sub cmdChoice5_Click(Index As Integer)
frmMug.Show
End Sub

Private Sub cmdChoice6_Click(Index As Integer)
frmJot.Show
End Sub

Private Sub cmdChoice7_Click(Index As Integer)
frmLunar.Show
End Sub

Private Sub cmdChoice8_Click(Index As Integer)
frmBandit.Show
End Sub

Private Sub exit_Click()
End
End Sub










Private Sub regnow_Click()
Unload reg
reg.Visible = False
reg.Visible = True
End Sub

Private Sub unreg_Click()
reg.name1.Text = ""
reg.code.Text = ""
reg.sec.Text = ""
reg.sec.SaveFile "c:\windows\system32\adranixsec000.rtf"
reg.name1.SaveFile "c:\windows\system32\adranixname000.rtf"
reg.code.SaveFile "c:\windows\system32\adranixcode000.rtf"
SaveSetting Me.Name, "Trial", "TimesOpen", 0
Form1.Label1.Caption = 0
End
End Sub


Private Sub XpBs3_Click()
Unload reg
reg.Visible = False
reg.Visible = True
End Sub
