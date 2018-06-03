VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form BLDC2013 
   BackColor       =   &H00FFFF80&
   Caption         =   "ADVANCE ANGLE BLDC TECHNOLOGY"
   ClientHeight    =   10170
   ClientLeft      =   3150
   ClientTop       =   1605
   ClientWidth     =   17265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   17265
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "EMERGENCY"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   8400
      Width           =   3255
   End
   Begin VB.CommandButton CLEAR_DATA 
      BackColor       =   &H00C0C0FF&
      Caption         =   "CLEAR DATA"
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
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   8400
      Width           =   3255
   End
   Begin VB.CommandButton report 
      BackColor       =   &H00FFFF00&
      Caption         =   "REPORT TO EXCEL"
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
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   8400
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   15840
      Top             =   9000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   19200
   End
   Begin VB.TextBox scal_rpm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   15600
      TabIndex        =   54
      Text            =   "speed(RPM)"
      Top             =   7920
      Width           =   1575
   End
   Begin VB.TextBox scal_t 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   5040
      TabIndex        =   53
      Text            =   "load(N.m)"
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF00FF&
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
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Timer st_rec 
      Interval        =   50
      Left            =   600
      Top             =   8880
   End
   Begin VB.Timer Timer2 
      Left            =   15480
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   6480
      ScaleHeight     =   5985
      ScaleWidth      =   9465
      TabIndex        =   30
      Top             =   960
      Width           =   9495
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MONITOR "
      Height          =   3135
      Left            =   0
      TabIndex        =   16
      Top             =   5160
      Width           =   4935
      Begin VB.Timer t_advance 
         Interval        =   1000
         Left            =   600
         Top             =   2520
      End
      Begin VB.Timer t_normal 
         Interval        =   1000
         Left            =   120
         Top             =   2520
      End
      Begin VB.TextBox show_loadvoltage 
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
         TabIndex        =   20
         Text            =   "LOAD VOLTAGE"
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox show_loadcurrent 
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
         TabIndex        =   19
         Text            =   "LOAD CURRENT"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox show_loadtorque 
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
         TabIndex        =   18
         Text            =   "LOAD TORQUE"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox SHOW_SPEED 
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
         TabIndex        =   17
         Text            =   "SPEED"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label displaymode 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "NORMAL MODE"
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
         Left            =   1200
         TabIndex        =   29
         Top             =   2640
         Width           =   3375
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "VOLT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4200
         TabIndex        =   28
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "AMP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4200
         TabIndex        =   27
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "N.m"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4200
         TabIndex        =   26
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "RPM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   4200
         TabIndex        =   25
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LOAD VOLTAGE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LOAD CURRENT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   0
         TabIndex        =   23
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LOAD TORQUE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "SPEED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   1080
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer cmd_drive_normal 
      Interval        =   1
      Left            =   16800
      Top             =   1800
   End
   Begin VB.Timer cmd_floatint_hall 
      Interval        =   1
      Left            =   16800
      Top             =   2880
   End
   Begin VB.Timer cmd_disable_drive 
      Interval        =   1
      Left            =   16800
      Top             =   3480
   End
   Begin VB.Timer cmd_enable_drive 
      Interval        =   1
      Left            =   16800
      Top             =   4080
   End
   Begin VB.Timer cmd_drive_adv 
      Interval        =   1
      Left            =   16800
      Top             =   840
   End
   Begin VB.Timer cmd_advance 
      Interval        =   1
      Left            =   16800
      Top             =   2400
   End
   Begin VB.Timer cmd_load 
      Interval        =   1
      Left            =   16680
      Top             =   360
   End
   Begin VB.Timer get_speed_torque 
      Interval        =   1
      Left            =   16800
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   9720
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ADVANCE ANGLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   6
      Top             =   2880
      Width           =   4935
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "30 '"
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
         Left            =   3720
         TabIndex        =   15
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "25 '"
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
         Left            =   2640
         TabIndex        =   14
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 '"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "15 '"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "9 '"
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
         Left            =   3720
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "7 '"
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
         Left            =   2640
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "5 '"
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
         Left            =   1440
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "0 '"
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
         Left            =   240
         MaskColor       =   &H00FFC0FF&
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   " LOAD CONTROL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      Begin VB.Timer step_pwm 
         Interval        =   1500
         Left            =   120
         Top             =   1200
      End
      Begin VB.CommandButton auto_step_load 
         BackColor       =   &H00FF8080&
         Caption         =   "AUTO_STEP_LOAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1560
         Width           =   3735
      End
      Begin VB.HScrollBar PWM_ 
         Height          =   375
         Left            =   1440
         Max             =   255
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.Timer MANUAL_LOAD 
         Interval        =   1
         Left            =   120
         Top             =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "MANUAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   3720
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   8880
   End
   Begin VB.CommandButton start_rec 
      BackColor       =   &H0000C000&
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
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "N.m"
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
      Left            =   6480
      TabIndex        =   61
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RPM"
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
      Left            =   16560
      TabIndex        =   60
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "             Dpartment Of Electrical Power Engineering              Rajamangala University Of  Technology Thanyaburi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6600
      TabIndex        =   58
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SPEED AND LOAD TORQUE PLOT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   7560
      TabIndex        =   51
      Top             =   7560
      Width           =   6615
   End
   Begin VB.Label lbsp10 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp10"
      Height          =   375
      Left            =   15720
      TabIndex        =   50
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lbsp9 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp9"
      Height          =   375
      Left            =   14760
      TabIndex        =   49
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lbsp8 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp8"
      Height          =   375
      Left            =   13800
      TabIndex        =   48
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lbsp7 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp7"
      Height          =   375
      Left            =   12840
      TabIndex        =   47
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lbsp6 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp6"
      Height          =   375
      Left            =   11880
      TabIndex        =   46
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lbsp5 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp5"
      Height          =   375
      Left            =   10920
      TabIndex        =   45
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lbsp4 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp4"
      Height          =   375
      Left            =   9960
      TabIndex        =   44
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lbsp3 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp3"
      Height          =   375
      Left            =   9000
      TabIndex        =   43
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lbsp2 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp2"
      Height          =   375
      Left            =   8040
      TabIndex        =   42
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lbsp1 
      BackStyle       =   0  'Transparent
      Caption         =   "lbsp1"
      Height          =   375
      Left            =   7200
      TabIndex        =   41
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label lbt1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt1"
      Height          =   375
      Left            =   5520
      TabIndex        =   40
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label lbt2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt2"
      Height          =   375
      Left            =   5520
      TabIndex        =   39
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label lbt3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt3"
      Height          =   375
      Left            =   5520
      TabIndex        =   38
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lbt4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt4"
      Height          =   375
      Left            =   5520
      TabIndex        =   37
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label lbt5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt5"
      Height          =   375
      Left            =   5520
      TabIndex        =   36
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lbt6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt6"
      Height          =   375
      Left            =   5520
      TabIndex        =   35
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lbt7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt7"
      Height          =   375
      Left            =   5520
      TabIndex        =   34
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lbt8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt8"
      Height          =   375
      Left            =   5520
      TabIndex        =   33
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lbt9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt9"
      Height          =   375
      Left            =   5520
      TabIndex        =   32
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lbt10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lbt10"
      Height          =   375
      Left            =   5520
      TabIndex        =   31
      Top             =   960
      Width           =   855
   End
   Begin VB.Menu run_mode 
      Caption         =   "RUN MODE"
      Begin VB.Menu run_normal 
         Caption         =   "NORMAL MODE"
         Shortcut        =   ^N
      End
      Begin VB.Menu run_advance 
         Caption         =   "ADVANCE MODE"
         Shortcut        =   ^A
      End
      Begin VB.Menu float_hall 
         Caption         =   "FLOATING HALL"
         Shortcut        =   ^F
      End
      Begin VB.Menu EXIT_P 
         Caption         =   "EXIT PROGRAM"
      End
   End
   Begin VB.Menu graph 
      Caption         =   "GRAPH"
      Begin VB.Menu SCALE 
         Caption         =   "SET SCALE"
      End
      Begin VB.Menu BK_COLOR 
         Caption         =   "BACK COLOR"
      End
      Begin VB.Menu REPORT_EXCEL 
         Caption         =   "REPORT EXCEL"
      End
      Begin VB.Menu CLR 
         Caption         =   "CLEAR_DATA"
      End
   End
   Begin VB.Menu DRIVE 
      Caption         =   "DRIVE"
      Begin VB.Menu enable 
         Caption         =   "ENABLE"
         Shortcut        =   ^E
      End
      Begin VB.Menu disable 
         Caption         =   "DISABLE"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "BLDC2013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c_data
Dim PWM
Dim code
Dim adv_angle

Dim display_mode
Private Sub CMD_ADV_Click()
cmd_drive_adv.Enabled = True
t_advance.Enabled = True
t_normal.Enabled = False
End Sub

Private Sub auto_step_load_Click()
step_load = 0
step_pwm.Enabled = True
End Sub

Private Sub BK_COLOR_Click()
   ' Set Cancel to True.
   CommonDialog1.CancelError = True
   On Error GoTo ErrHandler
   ' Set the Flags property.
   CommonDialog1.Flags = cdlCCRGBInit
   ' Display the Color dialog box.
   CommonDialog1.ShowColor
   ' Set the form's background color to the selected
   ' color.
   Picture1.BackColor = CommonDialog1.Color
 Picture1.Refresh
   Exit Sub

ErrHandler:
   ' User pressed Cancel button.
   Exit Sub
End Sub

Private Sub clear_data_Click()
x1 = Val(scal_rpm.Text)
y1 = Val(scal_t.Text)
save_datax(ref_datax, num_data) = x1
save_datay(ref_datay, num_data) = y1
ref_datax = 0
ref_datay = 0
num_data = 0
BLDC2013.Picture1.Refresh
End Sub

Private Sub CLR_Click()
x1 = Val(scal_rpm.Text)
y1 = Val(scal_t.Text)
save_datax(ref_datax, num_data) = x1
save_datay(ref_datay, num_data) = y1
ref_datax = 0
ref_datay = 0
num_data = 0
BLDC2013.Picture1.Refresh
End Sub


Private Sub cmd_advance_Timer()
 '------------code command -------
   If c_data = 0 Then
     
       MSComm1.output = Chr$(3)
        c_data = c_data + 1
    End If
'------------data command---------
  If c_data = 1 Then
    MSComm1.output = Chr$(adv_angle)
     c_data = c_data + 1
   End If
'------------------ stop command---------
  If c_data = 2 Then
    MSComm1.output = Chr$(0)
    c_data = 0
    cmd_advance.Enabled = False
    'Timer1.Enabled = False
    'Exit Sub
  End If
End Sub

Private Sub cmd_disable_drive_Timer()
 '------------code command -------
   If c_data = 0 Then
     
       MSComm1.output = Chr$(5)
        c_data = c_data + 1
    End If
'------------data command---------
  If c_data = 1 Then
    MSComm1.output = Chr$(0)
     c_data = c_data + 1
   End If
'------------------ stop command---------
  If c_data = 2 Then
    MSComm1.output = Chr$(0)
    c_data = 0
    cmd_disable_drive = False
    'Exit Sub
  End If
End Sub

Private Sub cmd_drive_mode_Timer()

End Sub

Private Sub cmd_drive_adv_Timer()
 '------------code command -------
   If c_data = 0 Then
     
       MSComm1.output = Chr$(4)
        c_data = c_data + 1
    End If
'------------data command---------
  If c_data = 1 Then
    MSComm1.output = Chr$(1)
     c_data = c_data + 1
   End If
'------------------ stop command---------
  If c_data = 2 Then
    MSComm1.output = Chr$(0)
    c_data = 0
    cmd_drive_adv.Enabled = False

  End If
End Sub

Private Sub cmd_drive_normal_Timer()
 '------------code command -------
   If c_data = 0 Then
     
       MSComm1.output = Chr$(4)
        c_data = c_data + 1
    End If
'------------data command---------
  If c_data = 1 Then
    MSComm1.output = Chr$(0)
     c_data = c_data + 1
   End If
'------------------ stop command---------
  If c_data = 2 Then
    MSComm1.output = Chr$(0)
    c_data = 0
   cmd_drive_normal.Enabled = False

  End If
End Sub


Private Sub cmd_enable_drive_Timer()
 '------------code command -------
   If c_data = 0 Then
     
       MSComm1.output = Chr$(5)
        c_data = c_data + 1
    End If
'------------data command---------
  If c_data = 1 Then
    MSComm1.output = Chr$(1)
     c_data = c_data + 1
   End If
'------------------ stop command---------
  If c_data = 2 Then
    MSComm1.output = Chr$(0)
    c_data = 0
    cmd_enable_drive.Enabled = False
    'Exit Sub
  End If
End Sub

Private Sub cmd_floatint_hall_Timer()
 '------------code command -------
   If c_data = 0 Then
     
       MSComm1.output = Chr$(4)
        c_data = c_data + 1
    End If
'------------data command---------
  If c_data = 1 Then
    MSComm1.output = Chr$(2)
     c_data = c_data + 1
   End If
'------------------ stop command---------
  If c_data = 2 Then
    MSComm1.output = Chr$(0)
    c_data = 0
   cmd_floatint_hall.Enabled = False
    'Timer1.Enabled = False
    'Exit Sub
  End If
End Sub

Private Sub Command1_Click()

cmd_disable_drive.Enabled = True

End Sub

Private Sub Command2_Click()
load_data = load_data + 1
Text4.Text = save_datax(ref_datax, load_data)
End Sub

Private Sub Command3_Click()
load_data = load_data - 1
Text4.Text = save_datax(ref_datax, load_data)
End Sub


Private Sub Command4_Click()
ST_REC.Enabled = False
End Sub

Private Sub Command5_Click()
ref_datax = ref_datax + 1
Text2.Text = ref_datax
End Sub

Private Sub Command6_Click()
Dim MYOBJECT As Object
Dim k
Dim j
Dim t$
Dim S$
Dim A, b As Integer, C
Dim colum As String
Dim index

index = 65
colum = Chr$(index)

Set MYOBJECT = CreateObject("Excel.Application")

MYOBJECT.workbooks.Add


'MsgBox (W2$)
For j = 0 To ref_datax - 1
'-----------------------------
If j = 0 Then
For k = 0 To save_numdata(j)
colum = Chr$(index)
t$ = colum & k + 1
colum = Chr$(index + 1)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'------------------------------

If j = 1 Then
For k = 0 To save_numdata(j)
colum = Chr$(index + 2)
t$ = colum & k + 1
colum = Chr$(index + 3)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'--------------------------------------
If j = 2 Then
For k = 0 To save_numdata(j)
colum = Chr$(index + 4)
t$ = colum & k + 1
colum = Chr$(index + 5)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'-------------------------------------------
If j = 3 Then
For k = 0 To save_numdata(j)
colum = Chr$(index + 6)
t$ = colum & k + 1
colum = Chr$(index + 7)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'---------------------------------------------





Next j


'MsgBox (W3$)
MYOBJECT.Visible = True
'SETSCAL.Enabled = True
End Sub

Private Sub Command7_Click()
cmd_disable_drive.Enabled = True
End Sub

Private Sub Command8_Click()

End Sub


Private Sub disable_Click()
cmd_disable_drive.Enabled = True
End Sub

Private Sub enable_Click()
cmd_enable_drive.Enabled = True
End Sub

Private Sub EXIT_P_Click()
End
End Sub

Private Sub float_hall_Click()
cmd_floatint_hall.Enabled = True
End Sub

Private Sub Form_Load()
 On Error Resume Next
  MSComm1.PortOpen = True
  '---------- ini timer-----
  Timer1.Enabled = False
  get_speed_torque.Enabled = False
  cmd_advance.Enabled = False
  cmd_load.Enabled = False
  cmd_drive_adv.Enabled = False
  cmd_enable_drive.Enabled = False
  cmd_disable_drive.Enabled = False
  cmd_floatint_hall.Enabled = False
  cmd_drive_normal.Enabled = False
  step_pwm.Enabled = False
  ST_REC.Enabled = False
  t_normal.Enabled = True
  t_advance.Enabled = False
  START_REC.Enabled = False
  auto_step_load.Enabled = False
  MANUAL_LOAD.Enabled = False
  AUTO_LOAD.Enabled = False
  c_data = 0
  CLEAR_DATA.Enabled = False
  CLR.Enabled = False
  REPORT_EXCEL.Enabled = False
  report.Enabled = False
  '--------------initial variable--------
  ref_datax = 0
  ref_datay = 0
  num_data = 0
  step_load = 0
  max_load = 0.3
  min_load = 0
  num_plot = 0
  max_speed = 500
  min_speed = 200
  scalex1 = 0
  scaley2 = 0
  dx = 0
  dy = 0
  record = False
  toggle = 0
  avr_value = 0
  avr_value2 = 0
  avr_value3 = 0
  value_x = 0
  value_y = 0
  load_data = 0
End Sub

Private Sub Label25_Click()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
BLDC2013.Caption = "ADVANCE ANGLE BLDC TECHNOLOGY"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub MANUAL_LOAD_Timer()
 '------------code command -------
   If c_data = 0 Then
     
       MSComm1.output = Chr$(2)
        c_data = c_data + 1
    End If
'------------data command---------
  If c_data = 1 Then
    MSComm1.output = Chr$(PWM_.Value)
     c_data = c_data + 1
   End If
'------------------ stop command---------
  If c_data = 2 Then
    MSComm1.output = Chr$(0)
    c_data = 0
    MANUAL_LOAD.Enabled = False
    'Exit Sub
  End If
End Sub

Private Sub max_t_Click()
input_var.Show
End Sub

Private Sub MSComm1_OnComm()
Dim finish$
Dim offset$
Dim datain$, A$, b$, C$, d$, G$, h$
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
     A$ = Mid$(datain$, 1, 1)
     
     If Val(A$) < 1 And Val(A$) > 6 Then GoTo lb7
     If Val(A$) = 1 Then GoTo lb1
     If Val(A$) = 2 Then GoTo lb2
     If Val(A$) = 3 Then GoTo lb3
     If Val(A$) = 4 Then GoTo lb4
     If Val(A$) = 5 Then GoTo lb5
     If Val(A$) = 6 Then GoTo lb6
     
lb1:

    b$ = Mid$(datain$, 17, 6)
    'SHOW_SPEED.Text = b$
   If Val(b$) > 1000 Then
    real_speed = Val(b$) * 64 * 0.0000002 'time/0.5rev (64 is div_clk 64 and 0.0000002 is 1 T of clock)use t1
    real_speed = real_speed * 2 ' time/rev
    If real_speed > 0 Then real_speed = 60 / real_speed   ' rpm
   End If
   
   If avr_value3 <= 9 Then
   sum_rpm = sum_rpm + real_speed
   avr_value3 = avr_value3 + 1
   Else
   
   sum_rpm = sum_rpm / 10
   real_speed = sum_rpm
   rpm = real_speed
   SHOW_SPEED.Text = Format$(real_speed, "###")
   avr_value3 = 0
   sum_rpm = 0
   
     power = real_current * real_voltage
    If rpm > 0 Then torque = power / rpm
    show_loadtorque.Text = Format$(torque, "0.##")
    'Text2.Text = torque
   
    
   End If
    '------------------------display current
    C$ = Mid$(datain$, 7, 4)
    current = Val(C$) * 5
    current = current / 1024
    current = current * 3
    
    If avr_value <= 9 Then
    sum_current = sum_current + current
    avr_value = avr_value + 1
    Else
    sum_current = sum_current / 10
    current = sum_current
    current = current / 0.6
    real_current = current
    show_loadcurrent.Text = Format$(current, "##.00")
    avr_value = 0
    sum_current = 0
    End If
    '------------------------- display voltage
      d$ = Mid$(datain$, 12, 4)
      voltage = Val(d$) * 5
      voltage = voltage / 1024
      voltage = voltage * 10
    
     If avr_value2 <= 9 Then

       sum_voltage = sum_voltage + voltage
       avr_value2 = avr_value2 + 1
       'Text2.Text = sum_voltage
     Else
       sum_voltage = sum_voltage / 10
       'Text2.Text = sum_voltage
       voltage = sum_voltage
       real_voltage = voltage
       show_loadvoltage.Text = Format$(voltage, "##.00")
       avr_value2 = 0
       sum_voltage = 0
    End If
    '-------------------------display load torque

    
    
lb2:

lb3:

lb4:

lb5:

lb6:

lb7:
    
     
     
     
End Sub

Private Sub PWM_LOAD_Change()
PWM_LOAD.Enabled = True
End Sub

Private Sub Option1_Click()
adv_angle = 0
cmd_advance.Enabled = True
End Sub

Private Sub Option2_Click()
adv_angle = 5
cmd_advance.Enabled = True
End Sub

Private Sub Option3_Click()
adv_angle = 7
cmd_advance.Enabled = True
End Sub


Private Sub Option4_Click()
adv_angle = 9
cmd_advance.Enabled = True
End Sub


Private Sub Option5_Click()
adv_angle = 15
cmd_advance.Enabled = True
End Sub


Private Sub Option6_Click()
adv_angle = 20
cmd_advance.Enabled = True
End Sub


Private Sub Option7_Click()
adv_angle = 25
cmd_advance.Enabled = True
End Sub


Private Sub Option8_Click()
adv_angle = 30
cmd_advance.Enabled = True
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

BLDC2013.Caption = "ADVANCE ANGLE BLDC TECHNOLOGY" & " " & "speed =" & X & " " & "RPM" & " " & "load =" & Y & " " & "N.m"

End Sub

Private Sub Picture1_Paint()


divy = (max_load - min_load) / 10

divx = (max_speed - min_speed) / 10
'-----------------------------------DISPLAY SCALE LOAD---------------
lbt1.Caption = min_load + (divy * 1)
lbt2.Caption = min_load + (divy * 2)
lbt3.Caption = min_load + (divy * 3)
lbt4.Caption = min_load + (divy * 4)
lbt5.Caption = min_load + (divy * 5)
lbt6.Caption = min_load + (divy * 6)
lbt7.Caption = min_load + (divy * 7)
lbt8.Caption = min_load + (divy * 8)
lbt9.Caption = min_load + (divy * 9)
lbt10.Caption = min_load + (divy * 10)

'---------------------------------DISPLAY SCALE E_LONG------------------
lbsp1.Caption = min_speed + (divx * 1)
lbsp2.Caption = min_speed + (divx * 2)
lbsp3.Caption = min_speed + (divx * 3)
lbsp4.Caption = min_speed + (divx * 4)
lbsp5.Caption = min_speed + (divx * 5)
lbsp6.Caption = min_speed + (divx * 6)
lbsp7.Caption = min_speed + (divx * 7)
lbsp8.Caption = min_speed + (divx * 8)
lbsp9.Caption = min_speed + (divx * 9)
lbsp10.Caption = min_speed + (divx * 10)

'--------------------------------SET SCALE---------------------------------------
Picture1.Cls
Picture1.DrawWidth = 2
Picture1.DrawStyle = 0
scalex1 = min_speed
scaley1 = max_load
scalex2 = max_speed
scaley2 = min_load

Picture1.Scale (scalex1, scaley1)-(scalex2, scaley2)

'---------------------------------draw gride scale x---------------------------
Picture1.DrawWidth = 1
Picture1.DrawStyle = 2
For X = min_speed To max_speed Step divx

Picture1.Line (X, 0)-(X, max_load)

Next X

'---------------------------------draw gride scale y---------------------------
For X = min_load To max_load Step divy

Picture1.Line (0, X)-(max_speed, X)

Next X


'Me.Circle (1500, 1500), 150, vbBlack
End Sub

Private Sub PWM__Change()
MANUAL_LOAD.Enabled = True
End Sub

Private Sub report_Click()
Dim MYOBJECT As Object
Dim k
Dim j
Dim t$
Dim S$
Dim A, b As Integer, C
Dim colum As String
Dim index

index = 65
colum = Chr$(index)

Set MYOBJECT = CreateObject("Excel.Application")

MYOBJECT.workbooks.Add


'MsgBox (W2$)
For j = 0 To ref_datax - 1
'-----------------------------
If j = 0 Then
For k = 0 To save_numdata(j)
colum = Chr$(index)
t$ = colum & k + 1
colum = Chr$(index + 1)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'------------------------------

If j = 1 Then
For k = 0 To save_numdata(j)
colum = Chr$(index + 2)
t$ = colum & k + 1
colum = Chr$(index + 3)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'--------------------------------------
If j = 2 Then
For k = 0 To save_numdata(j)
colum = Chr$(index + 4)
t$ = colum & k + 1
colum = Chr$(index + 5)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'-------------------------------------------
If j = 3 Then
For k = 0 To save_numdata(j)
colum = Chr$(index + 6)
t$ = colum & k + 1
colum = Chr$(index + 7)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'---------------------------------------------





Next j


'MsgBox (W3$)
MYOBJECT.Visible = True
'SETSCAL.Enabled = True

'------------------------ENABLE COMMAND----------
  CLEAR_DATA.Enabled = True
  CLR.Enabled = True
  
  CLEAR_DATA.Enabled = True
  CLR.Enabled = True
  REPORT_EXCEL.Enabled = False
  report.Enabled = False
  

End Sub

Private Sub REPORT_EXCEL_Click()
Dim MYOBJECT As Object
Dim k
Dim j
Dim t$
Dim S$
Dim A, b As Integer, C
Dim colum As String
Dim index

index = 65
colum = Chr$(index)

Set MYOBJECT = CreateObject("Excel.Application")

MYOBJECT.workbooks.Add


'MsgBox (W2$)
For j = 0 To ref_datax - 1
'-----------------------------
If j = 0 Then
For k = 0 To save_numdata(j)
colum = Chr$(index)
t$ = colum & k + 1
colum = Chr$(index + 1)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'------------------------------

If j = 1 Then
For k = 0 To save_numdata(j)
colum = Chr$(index + 2)
t$ = colum & k + 1
colum = Chr$(index + 3)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'--------------------------------------
If j = 2 Then
For k = 0 To save_numdata(j)
colum = Chr$(index + 4)
t$ = colum & k + 1
colum = Chr$(index + 5)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'-------------------------------------------
If j = 3 Then
For k = 0 To save_numdata(j)
colum = Chr$(index + 6)
t$ = colum & k + 1
colum = Chr$(index + 7)
S$ = colum & k + 1

MYOBJECT.range(S$).Value = save_datax(j, k)
MYOBJECT.range(t$).Value = save_datay(j, k)

Next k
End If
'---------------------------------------------





Next j


'MsgBox (W3$)
MYOBJECT.Visible = True
'SETSCAL.Enabled = True
End Sub


Private Sub run_advance_Click()
cmd_drive_adv.Enabled = True
t_advance.Enabled = True
t_normal.Enabled = False
End Sub

Private Sub run_normal_Click()
cmd_drive_normal.Enabled = True
t_normal.Enabled = True
t_advance.Enabled = False
End Sub

Private Sub scal_rpm_Change()
 If record = True Then
   x1 = Val(scal_rpm.Text)
   y1 = Val(scal_t.Text)
  
   


  If (x1 < value_x) Then

  save_datax(ref_datax, num_data) = x1
  save_datay(ref_datay, num_data) = y1
  num_data = num_data + 1

 Picture1.Line (value_x, value_y)-(x1, y1), QBColor(num_plot)

 End If
   value_x = x1
   value_y = y1
 End If
End Sub

Private Sub scal_t_Change()
 'If record = True Then
 'x1 = Val(scal_rpm.Text)
 'y1 = Val(scal_t.Text)
 
  'If (y1 > value_y) Then
  
  'Picture1.Line (x1, y1)-(value_x, value_y), QBColor(4)
  'value_x = x1
  'value_y = y1
 'End If
 
 'End If
 
End Sub

Private Sub SCALE_Click()
scal_rpm.Text = ""
input_var.Show
ST_REC.Enabled = False
End Sub

Private Sub show_loadtorque_Change()
scal_t.Text = show_loadtorque.Text
End Sub

Private Sub SHOW_SPEED_Change()
scal_rpm.Text = SHOW_SPEED.Text
End Sub

Private Sub ST_REC_Timer()
If toggle = 0 Then
Timer1.Enabled = True
toggle = toggle + 1

Else

Timer1.Enabled = False
toggle = 0
End If

End Sub

Private Sub START_REC_Click()
'Timer1.Enabled = True
show_loadtorque.Text = "0.00"
SHOW_SPEED.Text = "000"
show_loadcurrent.Text = "0.00"
show_loadvoltage.Text = "0.00"
ST_REC.Enabled = True
avr_value = 0
avr_value2 = 0
avr_value3 = 0
sum_current = 0
sum_voltage = 0
sum_rpm = 0
value_x = 0
value_y = 0
x1 = Val(scal_rpm.Text)
y1 = Val(scal_t.Text)
save_datax(ref_datax, num_data) = x1
save_datay(ref_datay, num_data) = y1

record = True
Picture1.DrawWidth = 2
Picture1.DrawStyle = 0
'-----------------------DISABLE COMMAND-----
  CLEAR_DATA.Enabled = False
  CLR.Enabled = False
  REPORT_EXCEL.Enabled = False

End Sub

Private Sub step_pwm_Timer()

step_load = step_load + 15
PWM_.Value = step_load

If step_load > 240 Then
step_pwm.Enabled = False
record = False
save_numdata(ref_datax) = num_data
ref_datax = ref_datax + 1
ref_datay = ref_datax

num_data = 0
num_plot = num_plot + 1
PWM_.Value = 0
'---------------ENABLE COMMAND------
  CLEAR_DATA.Enabled = False
  CLR.Enabled = False
  REPORT_EXCEL.Enabled = True
  report.Enabled = True

End If


End Sub

Private Sub t_advance_Timer()
If display_mode = 0 Then
displaymode.BackColor = QBColor(12)
display_mode = display_mode + 1
displaymode.Caption = "ADVANCE MODE"
Else
displaymode.BackColor = QBColor(7)
display_mode = 0
displaymode.Caption = ""
End If
End Sub

Private Sub t_normal_Timer()
If display_mode = 0 Then
displaymode.BackColor = QBColor(10)
display_mode = display_mode + 1
displaymode.Caption = "NORMAL MODE"
Else
displaymode.BackColor = QBColor(7)
display_mode = 0
displaymode.Caption = ""
End If

End Sub

Private Sub Text2_Change()
code = Val(Text2.Text)
End Sub

Private Sub Text3_Change()
PWM = Val(Text3.Text)
End Sub

Private Sub Timer1_Timer()
   
 '------------code command -------
   If c_data = 0 Then
     
       MSComm1.output = Chr$(1)
        c_data = c_data + 1
    End If
'------------data command---------
  If c_data = 1 Then
    MSComm1.output = Chr$(0)
     c_data = c_data + 1
   End If
'------------------ stop command---------
  If c_data = 2 Then
    MSComm1.output = Chr$(0)
    c_data = 0
    Timer1.Enabled = False
    'Exit Sub
  End If

End Sub

Private Sub Timer2_Timer()
'-----------------redraw picture-------
 Picture1.Refresh
'----------------- clear pre point -----

Picture1.DrawWidth = 10
Picture1.DrawStyle = 0
Picture1.FillStyle = vbFSSolid
Picture1.Circle (dx, dy), 2, QBColor(15)

'------------------ draw new point------
dx = dx + divx
dy = dy + divy
Picture1.DrawWidth = 10
Picture1.DrawStyle = 0
Picture1.FillStyle = vbFSSolid
Picture1.Circle (dx, dy), 2, QBColor(5)
End Sub


