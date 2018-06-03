VERSION 5.00
Begin VB.Form input_var 
   BackColor       =   &H00FFFFC0&
   Caption         =   "SET SCALE"
   ClientHeight    =   3885
   ClientLeft      =   7845
   ClientTop       =   2640
   ClientWidth     =   8790
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   8790
   Begin VB.TextBox input_maxload 
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
      Height          =   615
      Left            =   4440
      TabIndex        =   11
      Text            =   "MIN"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox input_maxspeed 
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
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Text            =   "MIN"
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
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
      Left            =   4440
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox input_minload 
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
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Text            =   "MAX "
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox input_minspeed 
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
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Text            =   "MAX "
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MAX"
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
      Left            =   4680
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MIN"
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
      Left            =   2400
      TabIndex        =   8
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6360
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LOAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SPEED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "input_var"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Val(input_maxspeed.Text) > Val(input_minspeed.Text) And Val(input_maxload.Text) > Val(input_minload.Text) Then
max_speed = Val(input_maxspeed.Text)
min_speed = Val(input_minspeed.Text)
max_load = Val(input_maxload.Text)
min_load = Val(input_minload.Text)

BLDC2013.Picture1.Refresh

Unload Me

Else

MsgBox "Input data incorrect", , "please check data"

End If

BLDC2013.st_rec.Enabled = True
BLDC2013.show_loadtorque.Text = "0.00"
BLDC2013.SHOW_SPEED.Text = "000"
BLDC2013.show_loadcurrent.Text = "0.00"
BLDC2013.show_loadvoltage.Text = "0.00"
BLDC2013.start_rec.Enabled = True
BLDC2013.auto_step_load.Enabled = True
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
input_minspeed.Text = ""
input_minload.Text = ""
input_maxload.Text = ""
input_maxspeed.Text = ""
End Sub

Private Sub min_load_Change()

End Sub


