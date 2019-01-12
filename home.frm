VERSION 5.00
Begin VB.Form home 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      Height          =   8415
      Left            =   0
      Picture         =   "home.frx":0000
      ScaleHeight     =   8355
      ScaleWidth      =   11835
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Shape Shape3 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   2
         DrawMode        =   5  'Not Copy Pen
         Height          =   495
         Left            =   10920
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "i"
         BeginProperty Font 
            Name            =   "Pacifico"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Left            =   11040
         TabIndex        =   6
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " Sign Up"
         BeginProperty Font 
            Name            =   "Aldrich"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9360
         TabIndex        =   5
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Height          =   975
         Left            =   9360
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Height          =   975
         Left            =   720
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "   LOGIN"
         BeginProperty Font 
            Name            =   "Aldrich"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   720
         TabIndex        =   4
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label title 
         BackStyle       =   0  'Transparent
         Caption         =   "Hashiko"
         BeginProperty Font 
            Name            =   "Pacifico"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   1095
         Left            =   480
         TabIndex        =   2
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label title2 
         BackStyle       =   0  'Transparent
         Caption         =   "Matrimonial"
         BeginProperty Font 
            Name            =   "Pacifico"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   855
         Left            =   1800
         TabIndex        =   1
         Top             =   960
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub login_Click()

End Sub

Private Sub Label2_Click()
Unload Me
login.Show
End Sub

Private Sub loginbtn_Click()
Unload Me
login.Show
End Sub



Private Sub Label3_Click()
Unload Me
Rsignup.Show
End Sub

Private Sub signbtn_Click()
Rsignup.Show
End Sub

Private Sub Label4_Click()
frmAbout.Show

End Sub

