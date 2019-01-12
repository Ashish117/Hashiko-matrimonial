VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Begin VB.Form Rsignup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Signup"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox fn 
      Height          =   285
      Left            =   7920
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Delete 
      BackColor       =   &H000000C0&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton update 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   20
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox instauser 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   3120
      MaxLength       =   15
      TabIndex        =   19
      Text            =   "Instagram username"
      Top             =   8280
      Width           =   2415
   End
   Begin VB.TextBox fbuser 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   240
      MaxLength       =   15
      TabIndex        =   18
      Text            =   "Facebook username"
      Top             =   8280
      Width           =   2415
   End
   Begin VB.ComboBox job 
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
      Left            =   240
      TabIndex        =   17
      Text            =   "Profession"
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox age 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   11040
      TabIndex        =   16
      Text            =   "age"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox aboutme 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      MaxLength       =   33
      TabIndex        =   15
      Text            =   "About"
      ToolTipText     =   "Something about you!"
      Top             =   7680
      Width           =   3735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   4920
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
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
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SQLServer;Initial Catalog=HashDB"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;Data Source=SQLServer;Initial Catalog=HashDB"
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
      Left            =   240
      MaxLength       =   8
      TabIndex        =   10
      Text            =   "Username"
      ToolTipText     =   "Username used to login to your profile"
      Top             =   1800
      Width           =   2415
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
      Left            =   240
      MaxLength       =   8
      TabIndex        =   9
      Text            =   "Password"
      Top             =   2640
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
      Left            =   240
      MaxLength       =   8
      TabIndex        =   8
      Text            =   "Name"
      Top             =   3480
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
      Left            =   240
      TabIndex        =   7
      Text            =   "Gender"
      Top             =   4320
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
      ItemData        =   "Rsignup1.frx":0000
      Left            =   240
      List            =   "Rsignup1.frx":0002
      TabIndex        =   6
      Text            =   "Religion"
      Top             =   5160
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
      Left            =   240
      TabIndex        =   5
      Text            =   "Caste"
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
      Left            =   8880
      MaxLength       =   8
      TabIndex        =   3
      Text            =   "Country"
      Top             =   5040
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
      Left            =   8880
      MaxLength       =   16
      TabIndex        =   2
      Text            =   "Location"
      Top             =   6000
      Width           =   2415
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
      Left            =   8880
      MaxLength       =   15
      TabIndex        =   1
      Text            =   "Email"
      Top             =   6960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9120
      MaskColor       =   &H00C0C0FF&
      Picture         =   "Rsignup1.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker datepk 
      Height          =   495
      Left            =   8880
      TabIndex        =   4
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   109051907
      CurrentDate     =   43231
   End
   Begin VB.Shape Shape18 
      BorderColor     =   &H00000000&
      BorderWidth     =   6
      Height          =   615
      Left            =   6720
      Top             =   7920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image dis 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2295
   End
   Begin VB.Shape Shape17 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   375
      Left            =   3120
      Top             =   8280
      Width           =   2415
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   375
      Left            =   240
      Top             =   8280
      Width           =   2415
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   240
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   11040
      Top             =   4200
      Width           =   615
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   375
      Left            =   240
      Top             =   7680
      Width           =   3735
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      Height          =   495
      Left            =   9360
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   9600
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00000000&
      BorderWidth     =   6
      Height          =   615
      Left            =   9120
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   240
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   240
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   240
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   240
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   240
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   240
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   8880
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   8880
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   8880
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   5
      Height          =   495
      Left            =   8880
      Top             =   6960
      Width           =   2415
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
      Left            =   240
      TabIndex        =   13
      Top             =   0
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
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   2655
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
      Left            =   10800
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin Project1.PictureG PictureG1 
      Height          =   9000
      Left            =   0
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   15875
      GIF             =   "Rsignup1.frx":36E6
   End
End
Attribute VB_Name = "Rsignup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim mStream As New ADODB.Stream

Function getFName(pf) As String: getFName = Mid(pf, InStrRev(pf, "\") + 1): End Function

Public Function AgeYears(ByVal datBirthDate As Date) As Integer

  On Error GoTo PROC_ERR

  Dim intYears As Integer

  intYears = Year(Now) - Year(datBirthDate)

  If DateSerial(Year(Now), Month(datBirthDate), Day(datBirthDate)) > Now Then
   ' Subtract a year if birthday hasn't arrived this year
    intYears = intYears - 1
  End If

  AgeYears = intYears

PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , "modDateTime.AgeYears"
  Resume PROC_EXIT
End Function



Private Sub aboutme_Change()
Shape13.BorderColor = 8454016
End Sub

Private Sub aboutme_GotFocus()
aboutme.Text = ""
Shape13.BorderColor = &HFF&

End Sub

Private Sub aboutme_LostFocus()
If aboutme.Text = "" Then
aboutme.Text = "About"
End If
End Sub



Private Sub Command1_Click()
'profile.Show
Dim fileName As String
If dis.Picture <> 0 Then
Dim fso As New FileSystemObject

fileName = fso.GetFileName(camfrm.pic.Caption)
Adodc1.Refresh
Adodc1.Recordset.AddNew
If lname.Text = "Username" Then
Shape1.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "A username is required to login!"
Exit Sub
Else
Adodc1.Recordset.Fields("username") = lname.Text
End If
If pass.Text = "Password" Then
Shape2.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "A password is required to login!"
Exit Sub
Else
Adodc1.Recordset.Fields("password") = pass.Text
End If
If uname.Text = "Name" Then
Shape3.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "A Name is required to Signup!"
Exit Sub
Else
Adodc1.Recordset.Fields("Name") = uname.Text
End If
If Gender.Text = "Gender" Then
Shape4.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Select your Gender to prove atleast you are a human being!"
Exit Sub
Else
Adodc1.Recordset.Fields("Gender") = Gender.Text
End If

If Religion.Text = "Religion" Then
Shape5.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Select your religion.. this is not an athiest matrimonial"
Exit Sub
Else
Adodc1.Recordset.Fields("Religion") = Religion.Text
End If

If Caste.Text = "" Then
Shape6.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Select your Caste to continue.."
Exit Sub
Else
Adodc1.Recordset.Fields("Caste") = Caste.Text
End If

If age.Text = "age" Then
Shape7.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Fill your Date Of Birth"
Exit Sub
Else
Adodc1.Recordset.Fields("DOB") = datepk.Value
Adodc1.Recordset.Fields("age") = age.Text
End If

If country.Text = "Country" Then
Shape8.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Country is required"
Exit Sub
Else
Adodc1.Recordset.Fields("Country") = country.Text
End If

If loc.Text = "Location" Then
Shape9.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Location is required"
Exit Sub
Else
Adodc1.Recordset.Fields("Location") = loc.Text
End If

If email.Text = "Location" Then
Shape10.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Email id is required"
Exit Sub
Else
Adodc1.Recordset.Fields("Email") = email.Text
End If

If aboutme.Text = "About" Then
Shape13.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "This field is required"
Exit Sub
Else
Adodc1.Recordset.Fields("about") = aboutme.Text
End If

If aboutme.Text = "Profession" Then
Shape15.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "You serious? cannot proceed without having a job"
Exit Sub
Else
Adodc1.Recordset.Fields("job") = job.Text
End If

If fbuser.Text = "Facebook username" Then
Shape16.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Input Facebook username to continue!"
Exit Sub
Else
Adodc1.Recordset.Fields("fbuser") = fbuser.Text
End If

If instauser.Text = "Instagram username" Then
Shape17.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Input Instagram username to continue!"
Exit Sub
Else
Adodc1.Recordset.Fields("instauser") = instauser.Text
End If

Adodc1.Recordset.Fields("DP") = fileName

'test

'MsgBox getFName(camfrm.pic.Caption)

location = (App.Path & "\DB\" & fileName)


SavePicture dis.Picture, location


'test close


Adodc1.Recordset.update
Adodc1.Recordset.Close
Unload Me
donesplsh.Show
Else
'Unload Me
'MsgBox "A Display picture is required to complete the registration", vbOKOnly
dialog.Show
dialog.msg.Caption = "A Display picture is required to complete the registration"


End If
End Sub


Private Sub country_KeyPress(KeyAscii As Integer)
If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii = "08") Or (KeyAscii = "32") Or (KeyAscii = "46") Then
country.Locked = False
Else
country.Locked = True
'MsgBox "ONLY ALPHABETS ARE ALLOWDED..!!"
dialog.Show
dialog.msg.Caption = "ONLY ALPHABETS ARE ALLOWDED..!!"
End If
End Sub



Private Sub loc_KeyPress(KeyAscii As Integer)
If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii = "08") Or (KeyAscii = "32") Or (KeyAscii = "46") Then
loc.Locked = False
Else
loc.Locked = True
'MsgBox "ONLY ALPHABETS ARE ALLOWDED..!!"
dialog.Show
dialog.msg.Caption = "ONLY ALPHABETS ARE ALLOWDED..!!"
End If
End Sub



Private Sub uname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii = "08") Or (KeyAscii = "32") Or (KeyAscii = "46") Then
uname.Locked = False
Else
uname.Locked = True
'MsgBox "ONLY ALPHABETS ARE ALLOWDED..!!"
dialog.Show
dialog.msg.Caption = "ONLY ALPHABETS ARE ALLOWDED..!!"
End If
End Sub



Private Sub Delete_Click()
confirmation = MsgBox("Do you want to delete this profile?", vbYesNo + vbCritical, "Delete Profile Confirmation")

If confirmation = vbYes Then

Adodc1.RecordSource = login.Adodc1.RecordSource
login.Adodc1.Recordset.Delete
'login.Adodc1.Recordset.update
'login.Adodc1.Recordset.Close
MsgBox "Profile has been deleted Successfully", vbInformation, "Message"
Unload Me
Unload PROFILES
Unload login
home.Show
Else

MsgBox "Profile Not Deleted !!!", vbInformation, "Message"

End If
End Sub

Private Sub Form_Load()
Gender.AddItem ("Male")
Gender.AddItem ("Female")
Religion.AddItem ("Christian")
Religion.AddItem ("Hindu")
Religion.AddItem ("Muslim")
Religion.AddItem ("Others")
'Profession
job.AddItem ("Architect")
job.AddItem ("Consultant")
job.AddItem ("Designer")
job.AddItem ("Doctor")
job.AddItem ("Manager")
job.AddItem ("Supervisor")
job.AddItem ("Pilot")
job.AddItem ("Nurse")
job.AddItem ("Software Eng")
job.AddItem ("Hardware ENg")
End Sub

Private Sub job_Click()
Shape15.BorderColor = 8454016
End Sub

Private Sub job_GotFocus()
Shape15.BorderColor = &HFF&
End Sub

Private Sub Label1_Click()
camfrm.Show

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


Private Sub Religion_Click()
Shape5.BorderColor = 8454016
If Religion.Text = "Christian" Then
Caste.Clear
Caste.AddItem ("Catholic")
Caste.AddItem ("Marthoma")
Caste.AddItem ("Orthodox")
Caste.AddItem ("Rc")
Caste.AddItem ("Ipc")
Caste.AddItem ("Jacobite")
Caste.AddItem ("Pentecost")
ElseIf Religion.Text = "Hindu" Then
Caste.Clear
Caste.AddItem ("Brahmin")
Caste.AddItem ("Nair")
Caste.AddItem ("Iyer")
Caste.AddItem ("Menon")
Caste.AddItem ("Warrier")
ElseIf Religion.Text = "Muslim" Then
Caste.Clear
Caste.AddItem ("Shiya")
Caste.AddItem ("Sunni")
Else
Caste.Clear
Caste.AddItem ("Any")
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

Private Sub fbuser_Change()
Shape16.BorderColor = 8454016
End Sub

Private Sub fbuser_GotFocus()
fbuser.Text = ""
Shape16.BorderColor = &HFF&

End Sub

Private Sub fbuser_LostFocus()
If fbuser.Text = "" Then
fbuser.Text = "Facebook username"
End If
End Sub
Private Sub instauser_Change()
Shape17.BorderColor = 8454016
End Sub

Private Sub instauser_GotFocus()
instauser.Text = ""
Shape17.BorderColor = &HFF&

End Sub

Private Sub instauser_LostFocus()
If instauser.Text = "" Then
instauser.Text = "Instagram username"
End If
End Sub

Private Sub Gender_Click()
Shape4.BorderColor = 8454016
End Sub

Private Sub Gender_GotFocus()
Shape4.BorderColor = &HFF&
End Sub



Private Sub Religion_GotFocus()
Shape5.BorderColor = &HFF&
End Sub

Private Sub Caste_Click()
Shape6.BorderColor = 8454016
End Sub

Private Sub Caste_GotFocus()
Shape6.BorderColor = &HFF&
End Sub

Private Sub datepk_Change()
Shape7.BorderColor = 8454016
Dim CurrentAge As Integer
CurrentAge = AgeYears(datepk.Value)

age.Text = CurrentAge
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


Private Sub update_Click()

Dim fileName As String
If dis.Picture <> 0 Then
Dim fso As New FileSystemObject
'Adodc1.RecordSource = "select * from profiles where username = '" & lname.Text & "' And password = '" & pass.Text & "'"
Adodc1.RecordSource = login.Adodc1.RecordSource
'Adodc1.Recordset.Open "select * from profiles where username = '" & lname.Text & "' And password = '" & pass.Text & "'", con, adOpenDynamic, adLockOptimistic
'While login.Adodc1.Recordset.EOF <> True

Adodc1.Refresh

'Do Until rs.EOF
fileName = fso.GetFileName(camfrm.pic.Caption)
If lname.Text = "Username" Then
Shape1.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "A username is required to login!"
Exit Sub
Else

Adodc1.Recordset.Fields("username") = lname.Text
'Wend
End If
If pass.Text = "Password" Then
Shape2.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "A password is required to login!"
Exit Sub
Else
Adodc1.Recordset.Fields("password") = pass.Text
End If
If uname.Text = "Name" Then
Shape3.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "A Name is required to Signup!"
Exit Sub
Else
Adodc1.Recordset.Fields("Name") = uname.Text
End If
If Gender.Text = "Gender" Then
Shape4.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Select your Gender to prove atleast you are a human being!"
Exit Sub
Else
Adodc1.Recordset.Fields("Gender") = Gender.Text
End If

If Religion.Text = "Religion" Then
Shape5.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Select your religion.. this is not an athiest matrimonial"
Exit Sub
Else
Adodc1.Recordset.Fields("Religion") = Religion.Text
End If

If Caste.Text = "" Then
Shape6.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Select your Caste to continue.."
Exit Sub
Else
Adodc1.Recordset.Fields("Caste") = Caste.Text
End If

If age.Text = "age" Then
Shape7.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Fill your Date Of Birth"
Exit Sub
Else
Adodc1.Recordset.Fields("DOB") = datepk.Value
Adodc1.Recordset.Fields("age") = age.Text
End If

If country.Text = "Country" Then
Shape8.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Country is required"
Exit Sub
Else
Adodc1.Recordset.Fields("Country") = country.Text
End If

If loc.Text = "Location" Then
Shape9.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Location is required"
Exit Sub
Else
Adodc1.Recordset.Fields("Location") = loc.Text
End If

If email.Text = "Location" Then
Shape10.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Email id is required"
Exit Sub
Else
Adodc1.Recordset.Fields("Email") = email.Text
End If

If aboutme.Text = "About" Then
Shape13.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "This field is required"
Exit Sub
Else
Adodc1.Recordset.Fields("about") = aboutme.Text
End If

If aboutme.Text = "Profession" Then
Shape15.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "You serious? cannot proceed without having a job"
Exit Sub
Else
Adodc1.Recordset.Fields("job") = job.Text
End If

If fbuser.Text = "Facebook username" Then
Shape16.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Input Facebook username to continue!"
Exit Sub
Else
Adodc1.Recordset.Fields("fbuser") = fbuser.Text
End If

If instauser.Text = "Instagram username" Then
Shape17.BorderColor = &HFF&
dialog.Show
dialog.msg.Caption = "Input Instagram username to continue!"
Exit Sub
Else
Adodc1.Recordset.Fields("instauser") = instauser.Text
End If

'Adodc1.Recordset.Fields("DP") = fileName
fn.Text = Adodc1.Recordset.Fields("DP")
'test

'MsgBox getFName(camfrm.pic.Caption)

location = (App.Path & "\DB\" & fn)


SavePicture dis.Picture, location


'test close


Adodc1.Recordset.update
Adodc1.Recordset.Close
Unload Me
confirm = MsgBox("You need to re-login to continue", vbOKOnly + vbInformation, "PROFILE UPDATED")
If confirm = vbOK Then
Unload PROFILES
Unload login
'dialog.Show
'dialog.msg.Caption = "You need to re-login to continue"
login.Show
Else
End If
Else
'Unload Me
'MsgBox "A Display picture is required to complete the registration", vbOKOnly
dialog.Show
dialog.msg.Caption = "A Display picture is required to complete the registration"


End If

End Sub

