VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form pid_temp 
   Caption         =   "Temperature Control"
   ClientHeight    =   8430
   ClientLeft      =   4245
   ClientTop       =   1710
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   14670
   Begin VB.CommandButton stop_rec 
      BackColor       =   &H008080FF&
      Caption         =   "STOP_REC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Timer SENT_DGAIN 
      Interval        =   20
      Left            =   2760
      Top             =   3480
   End
   Begin VB.Timer SENT_IGAIN 
      Interval        =   20
      Left            =   2760
      Top             =   1800
   End
   Begin VB.Timer SENT_PGAIN 
      Interval        =   20
      Left            =   2760
      Top             =   600
   End
   Begin VB.TextBox STEP_DGAIN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   47
      Text            =   "STEP"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox SET_DGAIN 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   44
      Text            =   "D_GAIN"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox STEP_IGAIN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   43
      Text            =   "STEP"
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox SET_IGAIN 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   40
      Text            =   "I_GAIN"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox STEP_PGAIN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   39
      Text            =   "STEP"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox set_pgain 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   36
      Text            =   "P_GAIN"
      Top             =   480
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13560
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox display_mintime 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   31
      Text            =   "START"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox display_maxtime 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   29
      Text            =   "START"
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox k 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   12120
      TabIndex        =   28
      Text            =   "k"
      Top             =   7680
      Width           =   855
   End
   Begin VB.Timer ST_REC 
      Interval        =   1000
      Left            =   0
      Top             =   7800
   End
   Begin VB.CommandButton START_REC 
      BackColor       =   &H0000FF00&
      Caption         =   "START_REC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7680
      Width           =   1935
   End
   Begin VB.TextBox sp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2160
      TabIndex        =   25
      Text            =   "SP"
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox mv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2160
      TabIndex        =   24
      Text            =   "MV"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox pv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2160
      TabIndex        =   23
      Text            =   "PV"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox time 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   8040
      TabIndex        =   22
      Text            =   "Time"
      Top             =   7680
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   4200
      ScaleHeight     =   5985
      ScaleWidth      =   9465
      TabIndex        =   1
      Top             =   1200
      Width           =   9495
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   13560
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   19200
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   50
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   49
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   48
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   35
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   34
      Top             =   5640
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   33
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "min_time"
      Height          =   495
      Left            =   0
      TabIndex        =   32
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "store_maxtime"
      Height          =   495
      Left            =   0
      TabIndex        =   30
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SEC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   27
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label lbt10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt10"
      Height          =   375
      Left            =   3240
      TabIndex        =   21
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lbt9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt9"
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lbt8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt8"
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lbt7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt7"
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lbt6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt6"
      Height          =   375
      Left            =   3240
      TabIndex        =   17
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lbt5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt5"
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label lbt4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt4"
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lbt3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt3"
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label lbt2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt2"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   6000
      Width           =   855
   End
   Begin VB.Label lbt1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt1"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label lbsp1 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp1"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lbsp2 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp2"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lbsp3 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp3"
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lbsp4 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp4"
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lbsp5 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp5"
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lbsp6 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp6"
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lbsp7 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp7"
      Height          =   375
      Left            =   10560
      TabIndex        =   5
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lbsp8 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp8"
      Height          =   375
      Left            =   11520
      TabIndex        =   4
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lbsp9 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp9"
      Height          =   375
      Left            =   12480
      TabIndex        =   3
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lbsp10 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp10"
      Height          =   375
      Left            =   13440
      TabIndex        =   2
      Top             =   7320
      Width           =   855
   End
End
Attribute VB_Name = "pid_temp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
value_pgain = True
SENT_PGAIN.Enabled = True

  'SENT_PGAIN.Enabled = False
  SENT_IGAIN.Enabled = False
  SENT_DGAIN.Enabled = False
End Sub

Private Sub Command2_Click()
value_pgain = False
SENT_PGAIN.Enabled = True
  'SENT_PGAIN.Enabled = False
  SENT_IGAIN.Enabled = False
  SENT_DGAIN.Enabled = False
End Sub

Private Sub Command3_Click()
value_igain = True
SENT_IGAIN.Enabled = True

  SENT_PGAIN.Enabled = False
  'SENT_IGAIN.Enabled = False
  SENT_DGAIN.Enabled = False
End Sub

Private Sub Command4_Click()
value_igain = False
SENT_IGAIN.Enabled = True

  SENT_PGAIN.Enabled = False
  'SENT_IGAIN.Enabled = False
  SENT_DGAIN.Enabled = False
End Sub

Private Sub Command5_Click()
value_dgain = True
SENT_DGAIN.Enabled = True

  SENT_PGAIN.Enabled = False
  SENT_IGAIN.Enabled = False
  'SENT_DGAIN.Enabled = False
End Sub

Private Sub Command6_Click()
value_dgain = False
SENT_DGAIN.Enabled = True

  SENT_PGAIN.Enabled = False
  SENT_IGAIN.Enabled = False
  'SENT_DGAIN.Enabled = False
End Sub

Private Sub Form_Load()
  MSComm1.PortOpen = True
  
 '-------------------set scale--------
  max_process = 1500
  min_process = 0
  
  max_time = 50
  store_maxtime = max_time
  min_time = 0
  
  SENT_PGAIN.Enabled = False
  SENT_IGAIN.Enabled = False
  SENT_DGAIN.Enabled = False
  c_data = 0
  
  STEP_PGAIN.Text = ""
  STEP_IGAIN.Text = ""
  STEP_DGAIN.Text = ""
  
  '
  ST_REC.Enabled = False
  scale_time = 0
  k = 1
  sum_loop = 0
  change_scale = False
End Sub

Private Sub MSComm1_OnComm()
Dim datain$, process$, output$, setpoint$, p$, i$, d$
Dim loop_check
Dim rpm$
Dim count$
Dim real_speed


    Do
    buffer$ = buffer$ & MSComm1.Input
    loop_check = loop_check + 1
    If loop_check > 100000 Then
    Exit Sub
    End If
    Loop Until InStr(buffer$, "s")
     datain$ = buffer$
     Text1 = datain$
     'A$ = Mid$(datain$, 1, 1)
     
     setpoint$ = Mid$(datain$, 1, 4)
     sp.Text = setpoint$
     
     process$ = Mid$(datain$, 5, 4)
     pv.Text = process$
     
     output$ = Mid$(datain$, 9, 4)
     mv.Text = output$
     
    p$ = Mid$(datain$, 13, 4)
     set_pgain.Text = Val(p$)
     
     
     i$ = Mid$(datain$, 17, 4)
     SET_IGAIN.Text = Val(i$)
     
     d$ = Mid$(datain$, 21, 4)
     SET_DGAIN.Text = Val(d$)
     
     
End Sub

Private Sub mv_Click()
   ' Set Cancel to True.
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   ' Set the Flags property.
   CommonDialog1.Flags = cdlCCRGBInit
   ' Display the Color dialog box.
   CommonDialog1.ShowColor
   ' Set the form's background color to the selected
   ' color.
   'Picture1.BackColor = CommonDialog1.Color
   mv.BackColor = CommonDialog1.Color
   color_mv = CommonDialog1.Color
   
   'Picture1.Refresh
   Exit Sub

ErrHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub


Private Sub Picture1_Paint()

'If change_scale = True Then

'divx = (max_time - min_time) / 10
'End If

If change_scale = False Then

divy = (max_process - min_process) / 10

divx = (max_time - min_time) / 10

'-----------------------------------DISPLAY SCALE LOAD---------------
lbt1.Caption = min_process + (divy * 1)
lbt2.Caption = min_process + (divy * 2)
lbt3.Caption = min_process + (divy * 3)
lbt4.Caption = min_process + (divy * 4)
lbt5.Caption = min_process + (divy * 5)
lbt6.Caption = min_process + (divy * 6)
lbt7.Caption = min_process + (divy * 7)
lbt8.Caption = min_process + (divy * 8)
lbt9.Caption = min_process + (divy * 9)
lbt10.Caption = min_process + (divy * 10)

'---------------------------------DISPLAY SCALE E_LONG------------------
lbsp1.Caption = min_time + (divx * 1)
lbsp2.Caption = min_time + (divx * 2)
lbsp3.Caption = min_time + (divx * 3)
lbsp4.Caption = min_time + (divx * 4)
lbsp5.Caption = min_time + (divx * 5)
lbsp6.Caption = min_time + (divx * 6)
lbsp7.Caption = min_time + (divx * 7)
lbsp8.Caption = min_time + (divx * 8)
lbsp9.Caption = min_time + (divx * 9)
lbsp10.Caption = min_time + (divx * 10)


'--------------------------------SET SCALE---------------------------------------
Picture1.Cls
Picture1.DrawWidth = 2
Picture1.DrawStyle = 0
scalex1 = min_process
scaley1 = max_process
scalex2 = max_time
scaley2 = min_time

Picture1.Scale (scalex1, scaley1)-(scalex2, scaley2)

'---------------------------------draw gride scale x---------------------------
Picture1.DrawWidth = 1
Picture1.DrawStyle = 2
For X = min_time To max_time Step divx

Picture1.Line (X, min_time)-(X, max_process)

Next X

'---------------------------------draw gride scale y---------------------------
For X = min_process To max_process Step divy

Picture1.Line (min_time, X)-(max_time, X)

Next X

Else
divy = (max_process - min_process) / 10

divx = (Val(display_maxtime.Text) - Val(display_mintime.Text)) / 10

'-----------------------------------DISPLAY SCALE LOAD---------------
lbt1.Caption = min_process + (divy * 1)
lbt2.Caption = min_process + (divy * 2)
lbt3.Caption = min_process + (divy * 3)
lbt4.Caption = min_process + (divy * 4)
lbt5.Caption = min_process + (divy * 5)
lbt6.Caption = min_process + (divy * 6)
lbt7.Caption = min_process + (divy * 7)
lbt8.Caption = min_process + (divy * 8)
lbt9.Caption = min_process + (divy * 9)
lbt10.Caption = min_process + (divy * 10)

'---------------------------------DISPLAY SCALE E_LONG------------------
lbsp1.Caption = Val(display_mintime.Text) + (divx * 1)
lbsp2.Caption = Val(display_mintime.Text) + (divx * 2)
lbsp3.Caption = Val(display_mintime.Text) + (divx * 3)
lbsp4.Caption = Val(display_mintime.Text) + (divx * 4)
lbsp5.Caption = Val(display_mintime.Text) + (divx * 5)
lbsp6.Caption = Val(display_mintime.Text) + (divx * 6)
lbsp7.Caption = Val(display_mintime.Text) + (divx * 7)
lbsp8.Caption = Val(display_mintime.Text) + (divx * 8)
lbsp9.Caption = Val(display_mintime.Text) + (divx * 9)
lbsp10.Caption = Val(display_mintime.Text) + (divx * 10)


'--------------------------------SET SCALE---------------------------------------
Picture1.Cls
Picture1.DrawWidth = 2
Picture1.DrawStyle = 0
scalex1 = Val(display_mintime.Text)
scaley1 = max_process
scalex2 = Val(display_maxtime.Text)
scaley2 = min_process

Picture1.Scale (scalex1, scaley1)-(scalex2, scaley2)

'---------------------------------draw gride scale x---------------------------
Picture1.DrawWidth = 1
Picture1.DrawStyle = 2
For X = Val(display_mintime.Text) To Val(display_maxtime.Text) Step divx

'Picture1.Line (x, Val(display_mintime.Text))-(x, max_process)

Picture1.Line (X, 0)-(X, max_process)

Next X

'---------------------------------draw gride scale y---------------------------
For X = min_process To max_process Step divy

Picture1.Line (Val(display_mintime.Text), X)-(Val(display_maxtime.Text), X)

Next X



change_scale = False
Picture1.DrawWidth = 4
End If



End Sub


Private Sub pv_Click()
Dim S
   ' Set Cancel to True.
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   ' Set the Flags property.
   CommonDialog1.Flags = cdlCCRGBInit
   ' Display the Color dialog box.
   CommonDialog1.ShowColor
   ' Set the form's background color to the selected
   ' color.
   'Picture1.BackColor = CommonDialog1.Color
   pv.BackColor = CommonDialog1.Color
   
  color_pv = CommonDialog1.Color
   'Picture1.Refresh
   Exit Sub

ErrHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub SENT_DGAIN_Timer()
'------------data command---------
  If c_data = 0 Then
    MSComm1.output = Chr$(Val(STEP_DGAIN.Text)) ' pgain value
     c_data = c_data + 1
   End If
'------------------plus or minus command---------
  If c_data = 1 Then
   '--------------minus------
     If value_dgain = False Then
     MSComm1.output = Chr$(0) ' 0 is decrease pgain value
     c_data = c_data + 1
     End If
    
   '--------------plus---------
     If value_dgain = True Then
     MSComm1.output = Chr$(1) ' 1 is increase pgain value
     c_data = c_data + 1
     End If
   End If
 '-----------------stop command-------
   If c_data = 2 Then
    MSComm1.output = Chr$(3) ' 1 = command sent pgain
    c_data = 0
    SENT_DGAIN.Enabled = False
   End If
End Sub

Private Sub SENT_IGAIN_Timer()
'------------data command---------
  If c_data = 0 Then
    MSComm1.output = Chr$(Val(STEP_IGAIN.Text)) ' pgain value
     c_data = c_data + 1
   End If
'------------------plus or minus command---------
  If c_data = 1 Then
   '--------------minus------
     If value_igain = False Then
     MSComm1.output = Chr$(0) ' 0 is decrease pgain value
     c_data = c_data + 1
     End If
    
   '--------------plus---------
     If value_igain = True Then
     MSComm1.output = Chr$(1) ' 1 is increase pgain value
     c_data = c_data + 1
     End If
   End If
 '-----------------stop command-------
   If c_data = 2 Then
    MSComm1.output = Chr$(2) ' 1 = command sent pgain
    c_data = 0
    SENT_IGAIN.Enabled = False
   End If
End Sub

Private Sub SENT_PGAIN_Timer()
'------------data command---------
  If c_data = 0 Then
    MSComm1.output = Chr$(Val(STEP_PGAIN.Text)) ' pgain value
     c_data = c_data + 1
   End If
'------------------plus or minus command---------
  If c_data = 1 Then
   '--------------minus------
     If value_pgain = False Then
     MSComm1.output = Chr$(0) ' 0 is decrease pgain value
     c_data = c_data + 1
     End If
    
   '--------------plus---------
     If value_pgain = True Then
     MSComm1.output = Chr$(1) ' 1 is increase pgain value
     c_data = c_data + 1
     End If
   End If
 '-----------------stop command-------
   If c_data = 2 Then
    MSComm1.output = Chr$(1) ' 1 = command sent pgain
    c_data = 0
    SENT_PGAIN.Enabled = False
   End If
   
End Sub

Private Sub sp_Click()
   ' Set Cancel to True.
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   ' Set the Flags property.
   CommonDialog1.Flags = cdlCCRGBInit
   ' Display the Color dialog box.
   CommonDialog1.ShowColor
   ' Set the form's background color to the selected
   ' color.
   'Picture1.BackColor = CommonDialog1.Color
  sp.BackColor = CommonDialog1.Color
   color_sp = CommonDialog1.Color
   
   'Picture1.Refresh
   Exit Sub

ErrHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub


Private Sub ST_REC_Timer()

scale_time = scale_time + 1
time.Text = scale_time

If scale_time >= store_maxtime Then

k = k + 1

store_mintime = store_maxtime
store_maxtime = store_maxtime + max_time


display_maxtime.Text = store_maxtime
display_mintime.Text = store_mintime


change_scale = True
Picture1.Refresh
End If

End Sub


Private Sub START_REC_Click()

ST_REC.Enabled = True
Picture1.DrawWidth = 4
Picture1.DrawStyle = 0
End Sub


Private Sub stop_rec_Click()


MSComm1.RThreshold = 0

End Sub

Private Sub time_Change()
 
   x1 = Val(time.Text)
   y1 = Val(pv.Text)
  
   y12 = Val(sp.Text)

   y13 = Val(mv.Text)


  'save_datax(ref_datax, num_data) = x1
  'save_datay(ref_datay, num_data) = y1
  num_data = num_data + 1
  Picture1.DrawStyle = 0
  Picture1.Line (value_x, value_y)-(x1, y1), color_pv
  
  Picture1.Line (value_x, value_y2)-(x1, y12), color_sp
  
  Picture1.Line (value_x, value_y3)-(x1, y13), color_mv

   value_x = x1
   value_y = y1
   
   value_y2 = y12
   
   value_y3 = y13

End Sub


