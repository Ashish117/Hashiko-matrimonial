VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SignUp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create a new Profile"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "signup.frx":0000
   ScaleHeight     =   8190
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   9360
      ScaleHeight     =   2415
      ScaleWidth      =   2295
      TabIndex        =   13
      Top             =   1200
      Width           =   2295
      Begin VB.Label up 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Display Picture"
         BeginProperty Font 
            Name            =   "Aldrich"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sign Up"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox email 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9240
      TabIndex        =   9
      Text            =   "Email"
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox loc 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Text            =   "Location"
      Top             =   6000
      Width           =   2415
   End
   Begin VB.TextBox country 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9240
      TabIndex        =   7
      Text            =   "Country"
      Top             =   5040
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker datepk 
      Height          =   495
      Left            =   9480
      TabIndex        =   6
      Top             =   4200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Exo 2 Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483632
      CalendarTitleBackColor=   16777215
      Format          =   96272385
      CurrentDate     =   43231
   End
   Begin VB.ComboBox Caste 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   600
      TabIndex        =   5
      Text            =   "Caste"
      Top             =   6960
      Width           =   2415
   End
   Begin VB.ComboBox Religion 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   600
      TabIndex        =   4
      Text            =   "Religion"
      Top             =   6000
      Width           =   2415
   End
   Begin VB.ComboBox Gender 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   600
      TabIndex        =   3
      Text            =   "Gender"
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox uname 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Text            =   " Name"
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox pass 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Text            =   "Password"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox lname 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Text            =   "Login Name"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label sversion 
      BackStyle       =   0  'Transparent
      Caption         =   "v1.1 -a"
      BeginProperty Font 
         Name            =   "Exo 2"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   11160
      TabIndex        =   15
      Top             =   120
      Width           =   615
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
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   360
      TabIndex        =   12
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label title 
      BackStyle       =   0  'Transparent
      Caption         =   "Hashiko"
      BeginProperty Font 
         Name            =   "Pacifico"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   855
      Left            =   360
      TabIndex        =   11
      Top             =   0
      Width           =   1695
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H8000000D&
      BorderWidth     =   9
      Height          =   735
      Left            =   5040
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   9240
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   9240
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   9240
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   9480
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   600
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   600
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   600
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   600
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   600
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   600
      Top             =   2160
      Width           =   2415
   End
End
Attribute VB_Name = "signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub

Private Sub lname_Change()
Shape1.BorderColor = 8454016
End Sub

Private Sub lname_Click()
lname.Text = ""
Shape1.BorderColor = &HFF&

End Sub

Private Sub lname_LostFocus()
If lname.Text = "" Then
lname.Text = "Login Name"
End If
End Sub

Private Sub pass_Change()
Shape2.BorderColor = 8454016
End Sub

Private Sub pass_GotFocus()
pass.Text = ""
Shape2.BorderColor = &HFF&

End Sub

Private Sub pass_LostFocus()
If pass.Text = "" Then
pass.Text = "Password"
End If
End Sub



Private Sub uname_Change()
Shape3.BorderColor = 8454016
End Sub

Private Sub uname_GotFocus()
uname.Text = ""
Shape3.BorderColor = &HFF&

End Sub

Private Sub uname_LostFocus()
If uname.Text = "" Then
uname.Text = "Name"
End If
End Sub

Private Sub Gender_Change()
Shape4.BorderColor = 8454016
End Sub

Private Sub Gender_GotFocus()
Shape4.BorderColor = &HFF&
End Sub

Private Sub Religion_Change()
Shape5.BorderColor = 8454016
End Sub

Private Sub Religion_GotFocus()
Shape5.BorderColor = &HFF&
End Sub

Private Sub Caste_Change()
Shape6.BorderColor = 8454016
End Sub

Private Sub Caste_GotFocus()
Shape6.BorderColor = &HFF&
End Sub

Private Sub datepk_Change()
Shape7.BorderColor = 8454016
End Sub

Private Sub datepk_GotFocus()
Shape7.BorderColor = &HFF&
End Sub

Private Sub country_Change()
Shape8.BorderColor = 8454016
End Sub

Private Sub country_GotFocus()
country.Text = ""
Shape8.BorderColor = &HFF&

End Sub

Private Sub country_LostFocus()
If country.Text = "" Then
country.Text = "Country"
End If
End Sub

Private Sub loc_Change()
Shape9.BorderColor = 8454016
End Sub

Private Sub loc_GotFocus()
loc.Text = ""
Shape9.BorderColor = &HFF&

End Sub

Private Sub loc_LostFocus()
If loc.Text = "" Then
loc.Text = "Location"
End If
End Sub

Private Sub email_Change()
Shape10.BorderColor = 8454016
End Sub

Private Sub email_GotFocus()
email.Text = ""
Shape10.BorderColor = &HFF&

End Sub

Private Sub email_LostFocus()
If email.Text = "" Then
email.Text = "Email"
End If
End Sub

