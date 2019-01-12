VERSION 5.00
Object = "{BF38D12B-22A9-4B10-B26E-019F2B5F9C22}#1.0#0"; "Ani Gif.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PROFILES 
   BackColor       =   &H00000000&
   Caption         =   "iMatching"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10200
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox rel 
      Height          =   285
      Left            =   12240
      TabIndex        =   12
      Text            =   "rel"
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox jobe 
      Height          =   285
      Left            =   12240
      TabIndex        =   9
      Text            =   "job"
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox gen 
      Height          =   285
      Left            =   12240
      TabIndex        =   8
      Text            =   "gen"
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00613737&
      BorderStyle     =   0  'None
      Height          =   9180
      Left            =   14880
      ScaleHeight     =   9180
      ScaleWidth      =   15000
      TabIndex        =   2
      Top             =   1200
      Width           =   15000
      Begin VB.ComboBox religion 
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
         Left            =   480
         TabIndex        =   7
         Text            =   "Religion"
         Top             =   2280
         Width           =   2415
      End
      Begin VB.ComboBox Job 
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
         Left            =   480
         TabIndex        =   6
         Text            =   "Profession"
         Top             =   5040
         Width           =   2415
      End
      Begin VB.ComboBox Castee 
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
         Left            =   480
         TabIndex        =   5
         Text            =   "Caste"
         Top             =   3600
         Width           =   2415
      End
      Begin VB.CommandButton srch 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Exo 2 Medium"
            Size            =   20.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   4
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H8000000D&
         BorderWidth     =   6
         Height          =   615
         Left            =   1920
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Height          =   495
         Left            =   480
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Height          =   495
         Left            =   480
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   5
         Height          =   495
         Left            =   480
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         Height          =   735
         Left            =   1920
         Shape           =   4  'Rounded Rectangle
         Top             =   -120
         Width           =   1935
      End
      Begin VB.Label filtr 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "Exo 2 Medium"
            Size            =   30
            Charset         =   0
            Weight          =   500
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Left            =   1920
         TabIndex        =   3
         Top             =   -120
         Width           =   1935
      End
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   9135
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   16113
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   -2147483641
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Aldrich"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "Form1.frx":8F87
   End
   Begin VB.Image logout 
      Height          =   450
      Left            =   19800
      Picture         =   "Form1.frx":1360F
      Stretch         =   -1  'True
      ToolTipText     =   "Logout"
      Top             =   480
      Width           =   510
   End
   Begin VB.Label Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10320
      TabIndex        =   11
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Smart"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image ton 
      Height          =   1005
      Left            =   8400
      Picture         =   "Form1.frx":13E00
      Stretch         =   -1  'True
      ToolTipText     =   "Enable or Disable Smart mode"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image toff 
      Height          =   885
      Left            =   8400
      Picture         =   "Form1.frx":1653A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin Project1.PictureG PictureG1 
      Height          =   1005
      Left            =   15000
      ToolTipText     =   "Messages and Notifications"
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1773
      GIF             =   "Form1.frx":17E47
      Stretch         =   3
   End
   Begin Project1.PictureG love 
      Height          =   1005
      Left            =   16680
      ToolTipText     =   "Requests"
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1773
      GIF             =   "Form1.frx":7E8AD
      Stretch         =   3
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      BorderWidth     =   16
      FillColor       =   &H00FFC0C0&
      Height          =   9135
      Left            =   120
      Top             =   1320
      Width           =   14655
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   135
      Left            =   17760
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   135
   End
   Begin VB.Label myprofile 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   18000
      TabIndex        =   1
      ToolTipText     =   "Your Profile"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      FillColor       =   &H00FFFFC0&
      Height          =   1215
      Left            =   -120
      Top             =   0
      Width           =   20655
   End
End
Attribute VB_Name = "PROFILES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim mStream As New ADODB.Stream



Private Sub logout_Click()
Unload login
Unload Me
home.Show
End Sub

Private Sub myprofile_Click()
Rsignup.Show
login.profup
End Sub

Private Sub srch_Click()

ListView.ListItems.Clear
rs.Close
rs.Open "Select * from profiles where Caste='" & Castee.Text & "' and job='" & Job.Text & "' and Gender ='" & gen.Text & "'", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView.ListItems.Add(, , rs!Name)
list.SubItems(1) = rs!age
list.SubItems(2) = rs!religion
list.SubItems(3) = rs!Caste
list.SubItems(4) = rs!Job
list.SubItems(5) = rs!about
list.SubItems(6) = rs!DP
list.SubItems(7) = rs!fbuser
list.SubItems(8) = rs!instauser
list.SubItems(9) = rs!Reported
rs.MoveNext
Loop

End Sub

Private Sub job_Click()
ListView.ListItems.Clear
rs.Close
rs.Open "Select * from profiles where job='" & Job.Text & "' and Gender ='" & gen.Text & "'", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView.ListItems.Add(, , rs!Name)
list.SubItems(1) = rs!age
list.SubItems(2) = rs!religion
list.SubItems(3) = rs!Caste
list.SubItems(4) = rs!Job
list.SubItems(5) = rs!about
list.SubItems(6) = rs!DP
list.SubItems(7) = rs!fbuser
list.SubItems(8) = rs!instauser
list.SubItems(9) = rs!Reported
rs.MoveNext
Loop
End Sub


Private Sub Castee_Click()
ListView.ListItems.Clear
rs.Close
rs.Open "Select * from profiles where Caste='" & Castee.Text & "' and Gender ='" & gen.Text & "'", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView.ListItems.Add(, , rs!Name)
list.SubItems(1) = rs!age
list.SubItems(2) = rs!religion
list.SubItems(3) = rs!Caste
list.SubItems(4) = rs!Job
list.SubItems(5) = rs!about
list.SubItems(6) = rs!DP
list.SubItems(7) = rs!fbuser
list.SubItems(8) = rs!instauser
list.SubItems(9) = rs!Reported
rs.MoveNext
Loop
End Sub



Private Sub Form_Load()
Picture1.BackColor = RGB(37, 15, 45)
'religion
religion.AddItem ("Christian")
religion.AddItem ("Hindu")
religion.AddItem ("Muslim")
religion.AddItem ("Others")

'Profession
Job.AddItem ("Architect")
Job.AddItem ("Consultant")
Job.AddItem ("Designer")
Job.AddItem ("Doctor")
Job.AddItem ("Manager")
Job.AddItem ("Supervisor")
Job.AddItem ("Pilot")
Job.AddItem ("Nurse")
Job.AddItem ("Software Eng")
Job.AddItem ("Hardware ENg")
'Me.BackColor = RGB(142, 68, 173)
ListView.ForeColor = RGB(255, 70, 255)
With ListView.ColumnHeaders
.Add , , "Name", Width / 7, lvwColumnLeft
.Add , , "Age", Width / 9, lvwColumnCenter
.Add , , "Religion", Width / 7, lvwColumnCenter
.Add , , "Caste", Width / 7, lvwColumnCenter
.Add , , "Profession", Width / 6, lvwColumnCenter
.Add , , "about", Width / 6, lvwColumnCenter
.Add , , "dp", Width / 6, lvwColumnCenter
.Add , , "fb", Width / 6, lvwColumnCenter
.Add , , "insta", Width / 6, lvwColumnCenter
.Add , , "Reported", Width / 6, lvwColumnCenter
End With
'loaddata
End Sub

Sub dbconnection()
If Not (con Is Nothing) Then
  If (con.State And adStateOpen) = adStateOpen Then con.Close
  Set con = Nothing
End If
con.Open "Provider=MSDASQL.1;User ID=sa;Password=alpine;Persist Security Info=True;User ID=sa;Data Source=SQLServer;Initial Catalog=HashDB"
End Sub

Sub loaddata()
Dim list As ListItem
ListView.ListItems.Clear
dbconnection
rs.Open "Select * from profiles where Gender ='" & gen.Text & "'", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView.ListItems.Add(, , rs!Name)
list.SubItems(1) = rs!age
list.SubItems(2) = rs!religion
list.SubItems(3) = rs!Caste
list.SubItems(4) = rs!Job
list.SubItems(5) = rs!about
list.SubItems(6) = rs!DP
list.SubItems(7) = rs!fbuser
list.SubItems(8) = rs!instauser
list.SubItems(9) = rs!Reported
rs.MoveNext
Loop

End Sub

Sub loadmale()
Dim list As ListItem
ListView.ListItems.Clear
dbconnection
'use if for each personals
Select Case jobe.Text
Case "Architect"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Male' and job IN ('Architect','Nurse','Consultant', 'Designer', 'Hardware ENg')", con, adOpenDynamic, adLockOptimistic
Case "Consultant"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Male' and job IN ('Designer','Manager','Supervisor','Nurse')", con, adOpenDynamic, adLockOptimistic
Case "Designer"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Male' and job IN ('Consultant','Architect','Manager','Software Eng')", con, adOpenDynamic, adLockOptimistic
Case "Doctor"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Male' and job IN ('Architect','Doctor','Designer','Software Eng','Pilot')", con, adOpenDynamic, adLockOptimistic
Case "Manager"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Male' and job IN ('Nurse','Software Eng','Hardware ENg','Consultant')", con, adOpenDynamic, adLockOptimistic
Case "Supervisor"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Male' and job IN ('Nurse','Supervisor','Hardware ENg','Consultant')", con, adOpenDynamic, adLockOptimistic
Case "Pilot"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Male' and job IN ('Software Eng','Pilot','Hardware ENg','Consultant')", con, adOpenDynamic, adLockOptimistic
Case "Nurse"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Male' and job IN ('Nurse','Consultant')", con, adOpenDynamic, adLockOptimistic
Case "Software Eng"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Male' and job IN ('Nurse','Software Eng','Hardware ENg','Doctor')", con, adOpenDynamic, adLockOptimistic
Case "Hardware ENg"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Male' and job IN ('Nurse','Software Eng','Hardware ENg','Doctor')", con, adOpenDynamic, adLockOptimistic
Case Else
Unload login
Unload PROFILES
dialog.Show
dialog.msg.Caption = "Why are you still here? It's over.. No one needs you"
Exit Sub
End Select
'rs.Open "Select * from profiles where Gender='Male'", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView.ListItems.Add(, , rs!Name)
list.SubItems(1) = rs!age
list.SubItems(2) = rs!religion
list.SubItems(3) = rs!Caste
list.SubItems(4) = rs!Job
list.SubItems(5) = rs!about
list.SubItems(6) = rs!DP
list.SubItems(7) = rs!fbuser
list.SubItems(8) = rs!instauser
list.SubItems(9) = rs!Reported
rs.MoveNext
Loop

End Sub

Sub loadfemale()
Dim list As ListItem
ListView.ListItems.Clear
dbconnection
'use if for each personals
Select Case jobe.Text
Case "Architect"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Female' and job IN ('Architect','Nurse','Consultant', 'Designer', 'Hardware ENg')", con, adOpenDynamic, adLockOptimistic
Case "Consultant"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Female' and job IN ('Designer','Manager','Supervisor','Nurse')", con, adOpenDynamic, adLockOptimistic
Case "Designer"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Female' and job IN ('Consultant','Architect','Manager','Software Eng')", con, adOpenDynamic, adLockOptimistic
Case "Doctor"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Female' and job IN ('Architect','Doctor','Designer','Software Eng','Pilot')", con, adOpenDynamic, adLockOptimistic
Case "Manager"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Female' and job IN ('Nurse','Software Eng','Hardware ENg','Consultant')", con, adOpenDynamic, adLockOptimistic
Case "Supervisor"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Female' and job IN ('Nurse','Supervisor','Hardware ENg','Consultant')", con, adOpenDynamic, adLockOptimistic
Case "Pilot"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Female' and job IN ('Software Eng','Pilot','Hardware ENg','Consultant')", con, adOpenDynamic, adLockOptimistic
Case "Nurse"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Female' and job IN ('Nurse','Consultant')", con, adOpenDynamic, adLockOptimistic
Case "Software Eng"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Female' and job IN ('Nurse','Software Eng','Hardware ENg','Doctor')", con, adOpenDynamic, adLockOptimistic
Case "Hardware ENg"
rs.Open "Select * from profiles where Religion ='" & rel.Text & "' and Gender='Female' and job IN ('Nurse','Software Eng','Hardware ENg','Doctor')", con, adOpenDynamic, adLockOptimistic
Case Else
dialog.msg.Caption = "Why are you still here? It's over.. No one needs you"
End Select
'rs.Open "Select * from profiles where Gender='Female'", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView.ListItems.Add(, , rs!Name)
list.SubItems(1) = rs!age
list.SubItems(2) = rs!religion
list.SubItems(3) = rs!Caste
list.SubItems(4) = rs!Job
list.SubItems(5) = rs!about
list.SubItems(6) = rs!DP
list.SubItems(7) = rs!fbuser
list.SubItems(8) = rs!instauser
list.SubItems(9) = rs!Reported
rs.MoveNext
Loop

End Sub





Private Sub Form_Unload(Cancel As Integer)
Unload Card
End Sub



Private Sub ListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
Me.Hide
  
Card.dis.Picture = LoadPicture(App.Path & "\DB\" & ListView.SelectedItem.SubItems(6))
Card.Disname.Caption = ListView.SelectedItem.Text
Card.age.Caption = ListView.SelectedItem.SubItems(1)
Card.religion.Caption = ListView.SelectedItem.SubItems(2)
Card.Caste.Caption = ListView.SelectedItem.SubItems(3)
Card.Job.Caption = ListView.SelectedItem.SubItems(4)
Card.about.Caption = ListView.SelectedItem.SubItems(5)
Card.fb.Text = ListView.SelectedItem.SubItems(7)
Card.insta.Text = ListView.SelectedItem.SubItems(8)
Card.Show
End Sub

Sub christian()
Dim list As ListItem
ListView.ListItems.Clear
dbconnection
rs.Open "Select * from profiles where Gender ='" & gen.Text & "' and Religion='Christian'", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView.ListItems.Add(, , rs!Name)
list.SubItems(1) = rs!age
list.SubItems(2) = rs!religion
list.SubItems(3) = rs!Caste
list.SubItems(4) = rs!Job
list.SubItems(5) = rs!about
list.SubItems(6) = rs!DP
list.SubItems(7) = rs!fbuser
list.SubItems(8) = rs!instauser
list.SubItems(9) = rs!Reported
rs.MoveNext
Loop
End Sub

Sub hindu()
Dim list As ListItem
ListView.ListItems.Clear
dbconnection
rs.Open "Select * from profiles where Gender ='" & gen.Text & "' and Religion='Hindu'", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView.ListItems.Add(, , rs!Name)
list.SubItems(1) = rs!age
list.SubItems(2) = rs!religion
list.SubItems(3) = rs!Caste
list.SubItems(4) = rs!Job
list.SubItems(5) = rs!about
list.SubItems(6) = rs!DP
list.SubItems(7) = rs!fbuser
list.SubItems(8) = rs!instauser
list.SubItems(9) = rs!Reported
rs.MoveNext
Loop
End Sub

Sub muslim()
Dim list As ListItem
ListView.ListItems.Clear
dbconnection
rs.Open "Select * from profiles where Gender ='" & gen.Text & "' and Religion='Muslim'", con, adOpenDynamic, adLockOptimistic
Do Until rs.EOF
Set list = ListView.ListItems.Add(, , rs!Name)
list.SubItems(1) = rs!age
list.SubItems(2) = rs!religion
list.SubItems(3) = rs!Caste
list.SubItems(4) = rs!Job
list.SubItems(5) = rs!about
list.SubItems(6) = rs!DP
list.SubItems(7) = rs!fbuser
list.SubItems(8) = rs!instauser
list.SubItems(9) = rs!Reported
rs.MoveNext
Loop
End Sub


Private Sub Religion_Click()
'Caste
If religion.Text = "Christian" Then
christian
Castee.Clear
Castee.Text = "Caste"
Castee.AddItem ("Catholic")
Castee.AddItem ("Marthoma")
Castee.AddItem ("Orthodox")
Castee.AddItem ("Rc")
Castee.AddItem ("Ipc")
Castee.AddItem ("Jacobite")
Castee.AddItem ("Pentecost")
ElseIf religion.Text = "Hindu" Then
hindu
Castee.Clear
Castee.Text = "Caste"
Castee.AddItem ("Brahmin")
Castee.AddItem ("Nair")
Castee.AddItem ("Iyer")
Castee.AddItem ("Menon")
Castee.AddItem ("Warrier")
ElseIf religion.Text = "Muslim" Then
muslim
Castee.Clear
Castee.Text = "Caste"
Castee.AddItem ("Shiya")
Castee.AddItem ("Sunni")
Else
Castee.Clear
Castee.Text = "Caste"
Castee.AddItem ("Any")
End If
End Sub

Sub smartmode()
If gen.Text = "Male" Then
'loaddata
'con.Close
loadmale
Else
'loaddata
'con.Close
loadfemale
End If

Text1.ForeColor = RGB(46, 204, 113)
Text2.ForeColor = RGB(46, 204, 113)
filtr.ForeColor = RGB(171, 178, 185)
religion.Enabled = False
Castee.Enabled = False
Job.Enabled = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
End Sub

Sub smartmodeoff()
loaddata
Text1.ForeColor = RGB(171, 178, 185)
Text2.ForeColor = RGB(171, 178, 185)
filtr.ForeColor = RGB(46, 204, 113)
religion.Enabled = True
Castee.Enabled = True
Job.Enabled = True
Shape5.Visible = True
Shape6.Visible = True
Shape7.Visible = True
End Sub

Private Sub ton_Click()
ton.Visible = False
toff.Visible = True
smartmodeoff
'Text2.ForeColor = RGB(240, 128, 128)
End Sub

Private Sub toff_Click()
toff.Visible = False
ton.Visible = True
smartmode
End Sub
