VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11865
   LinkTopic       =   "Login to your profile"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   9000
      Left            =   -120
      Picture         =   "login.frx":0000
      ScaleHeight     =   9000
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   -120
      Width           =   12000
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   855
         Left            =   8280
         Top             =   7200
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1508
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=MSDASQL.1;User ID=sa;Password=alpine;Persist Security Info=True;User ID=sa;Data Source=SQLServer;Initial Catalog=HashDB"
         OLEDBString     =   "Provider=MSDASQL.1;User ID=sa;Password=alpine;Persist Security Info=True;User ID=sa;Data Source=SQLServer;Initial Catalog=HashDB"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from profiles"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1320
         Picture         =   "login.frx":1BDB8
         ScaleHeight     =   495
         ScaleWidth      =   1695
         TabIndex        =   6
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox pass 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00404040&
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "Password"
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox lname 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "Username"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C000C0&
         BorderWidth     =   3
         Height          =   495
         Left            =   1320
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Login to "
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
         Height          =   975
         Left            =   5640
         TabIndex        =   5
         Top             =   120
         Width           =   1815
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
         ForeColor       =   &H00FF8080&
         Height          =   855
         Left            =   7320
         TabIndex        =   4
         Top             =   120
         Width           =   1695
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
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   9000
         TabIndex        =   3
         Top             =   120
         Width           =   2655
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Height          =   495
         Left            =   1080
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Height          =   495
         Left            =   1080
         Top             =   2160
         Width           =   2415
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic As String

Sub profup()
Rsignup.update.Visible = True
Rsignup.Shape18.Visible = True
Rsignup.Delete.Visible = True
Rsignup.lname.Text = Adodc1.Recordset!UserName
Rsignup.uname.Text = Adodc1.Recordset!Name
Rsignup.pass.Text = Adodc1.Recordset!Password
Rsignup.Gender.Text = Adodc1.Recordset!Gender
Rsignup.religion.Text = Adodc1.Recordset!religion
Rsignup.Caste.Text = Adodc1.Recordset!Caste
Rsignup.Job.Text = Adodc1.Recordset!Job
Rsignup.aboutme.Text = Adodc1.Recordset!about
Rsignup.fbuser.Text = Adodc1.Recordset!fbuser
Rsignup.instauser.Text = Adodc1.Recordset!instauser
Rsignup.datepk.Value = Adodc1.Recordset!DOB
Rsignup.country.Text = Adodc1.Recordset!country
Rsignup.loc.Text = Adodc1.Recordset!location
Rsignup.email.Text = Adodc1.Recordset!email
fileName = Adodc1.Recordset!DP
location = (App.Path & "\DB\" & fileName)
Rsignup.dis.Picture = LoadPicture(location)
End Sub
Sub display()

'profile.lname.Text = Adodc1.Recordset!UserName
'profile.Picture1.Picture = LoadPicture(Adodc1.Recordset!dis)
PROFILES.myprofile.Caption = Adodc1.Recordset!Name
PROFILES.jobe.Text = Adodc1.Recordset!Job
PROFILES.rel.Text = Adodc1.Recordset!religion
'Card.fb.Text = Adodc1.Recordset!fbuser
'Card.insta.Text = Adodc1.Recordset!instauser
If Adodc1.Recordset!Gender = "Male" Then
PROFILES.gen.Text = "Female"
'PROFILES.smartmodeoff
Else
PROFILES.gen.Text = "Male"
'PROFILES.loadmale

End If
PROFILES.smartmode
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
lname.Text = "Username"
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



Private Sub Picture2_Click()
If lname.Text = "Username" Then
dialog.Show
dialog.msg.Caption = "Type in your username!"
'Beep
Exit Sub
ElseIf pass.Text = "Password" Then
dialog.Show
dialog.msg.Caption = "Type in your password!"
Exit Sub
ElseIf lname.Text = "admin" And pass.Text = "admin" Then
Admin.Show
Exit Sub
Else
Adodc1.RecordSource = "Select * from profiles where username ='" & lname.Text & "' and password ='" & pass.Text & "' "
'End If
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Login Failed.. Please login with correct credentials", vbCritical
login.Show
lname.Text = "Username"
pass.Text = "Password"
Else
'MsgBox "Well Done..Login Successful", vbInformation
PROFILES.Show
display
login.Hide
End If
End If
End Sub

