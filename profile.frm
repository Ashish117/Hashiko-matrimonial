VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form profile 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   8880
      ScaleHeight     =   2415
      ScaleWidth      =   2295
      TabIndex        =   11
      Top             =   600
      Width           =   2295
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1935
      End
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
      Left            =   8760
      TabIndex        =   8
      Text            =   "Email"
      Top             =   6600
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
      Left            =   8760
      TabIndex        =   7
      Text            =   "Location"
      Top             =   5640
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
      Left            =   8760
      TabIndex        =   6
      Text            =   "Country"
      Top             =   4680
      Width           =   2415
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
      Left            =   120
      TabIndex        =   5
      Text            =   "Caste"
      Top             =   6600
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
      Left            =   120
      TabIndex        =   4
      Text            =   "Religion"
      Top             =   5640
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
      Left            =   120
      TabIndex        =   3
      Text            =   "Gender"
      Top             =   4680
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
      Left            =   120
      TabIndex        =   2
      Text            =   " Name"
      Top             =   3720
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
      Left            =   120
      TabIndex        =   1
      Text            =   "Password"
      Top             =   2760
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
      Left            =   120
      TabIndex        =   0
      Text            =   "Username"
      Top             =   1800
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker datepk 
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   3840
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
      Format          =   109051905
      CurrentDate     =   43231
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000C0&
      Caption         =   "NOT COMPLETED"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   13
      Top             =   3120
      Width           =   3855
   End
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
      Left            =   8880
      TabIndex        =   10
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim confirm As Integer

Private Sub Firstbtn_Click()
rs.MoveFirst
display
End Sub

Private Sub lastbtn_Click()
rs.MoveLast
display
End Sub

Private Sub nextbtn_Click()
rs.MoveNext
If Not rs.EOF Then
display
Else
rs.MoveFirst
display
End If
End Sub

Private Sub Previousbtn_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Else
display
End If
End Sub

Private Sub Form_Load()
con.Open "Provider=MSDASQL.1;User ID=sa;Password=alpine;Persist Security Info=True;User ID=sa;Data Source=SQLServer;Initial Catalog=HashDB"
rs.Open "Select * from profiles", con, adOpenDynamic, adLockPessimistic
End Sub

Sub display()
lname.Text = rs!UserName
pass.Text = rs!Password
uname.Text = rs!Name

End Sub
