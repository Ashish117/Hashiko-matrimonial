VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Card 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "profession"
   ClientHeight    =   9330
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Card.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5880
      Top             =   3240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=False;User ID=sa;Data Source=SQLServer;Initial Catalog=HashDB"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=False;User ID=sa;Data Source=SQLServer;Initial Catalog=HashDB"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select Reported from profiles"
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
   Begin VB.TextBox one 
      Height          =   285
      Left            =   2640
      TabIndex        =   14
      Text            =   "0"
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox insta 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox fb 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox back 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      Picture         =   "Card.frx":000C
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3720
      Picture         =   "Card.frx":0412
      ScaleHeight     =   1095
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   7320
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1320
      Picture         =   "Card.frx":147F
      ScaleHeight     =   1095
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label report 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Franklin Gothic Book"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2640
      TabIndex        =   13
      Top             =   8520
      Width           =   1005
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   12
      Height          =   1935
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image dis 
      Height          =   1935
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label job 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Profession"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   2280
      TabIndex        =   10
      Top             =   4560
      Width           =   1605
   End
   Begin VB.Label age 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "age"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   4020
      TabIndex        =   8
      Top             =   3840
      Width           =   525
   End
   Begin VB.Label religion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Religion"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   2505
      TabIndex        =   7
      Top             =   3840
      Width           =   1155
   End
   Begin VB.Label about 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   2040
      TabIndex        =   5
      Top             =   5160
      Width           =   2205
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   615
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label send 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Send Interest"
      BeginProperty Font 
         Name            =   "Exo 2"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2400
      TabIndex        =   4
      Top             =   6360
      Width           =   1395
   End
   Begin VB.Label caste 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Caste"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2685
      TabIndex        =   3
      Top             =   4200
      Width           =   765
   End
   Begin VB.Label Disname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Yourname"
      BeginProperty Font 
         Name            =   "Merienda One"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   2040
      TabIndex        =   2
      Top             =   3240
      Width           =   2145
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   8415
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINEINDEX = &HBB
Private Const EM_GETLINE = &HC4






Private Sub back_Click()
PROFILES.Show
End Sub

Private Sub Form_Load()
Disname.ForeColor = RGB(0, 128, 255)
age.ForeColor = RGB(243, 142, 188)
Me.BackColor = RGB(72, 61, 139)
send.ForeColor = RGB(155, 89, 182)
Shape2.BorderColor = RGB(0, 0, 0)


 Dim lngCount As Long
    Dim lngLineIndex As Long
    Dim lngLength As Long
    Dim strBuffer As String
    Dim lngIndex As Long
    Dim strSave As String
    
    'Get the line count
    lngCount = SendMessage(Text1.hWnd, EM_GETLINECOUNT, 0, 0)

    If lngCount > 3 Then
        With Text1
            For lngIndex = 0 To lngCount - 2
                lngLineIndex = SendMessage(.hWnd, EM_LINEINDEX, lngIndex, 0)
                lngLength = SendMessage(.hWnd, EM_LINELENGTH, lngLineIndex, 0)
                strBuffer = Space(lngLength)
                Call SendMessageStr(.hWnd, EM_GETLINE, lngIndex, ByVal strBuffer)
                strSave = strSave & strBuffer & vbCrLf
            Next
            .Text = Left$(strSave, Len(strSave) - 2)
        End With
        Exit Sub
    End If
    
    With Text1
        For lngIndex = 0 To lngCount - 1
            'Get line index of the chosen line
            lngLineIndex = SendMessage(.hWnd, EM_LINEINDEX, lngIndex, 0)
            'get line length
            lngLength = SendMessage(.hWnd, EM_LINELENGTH, lngLineIndex, 0)
'            'resize buffer
            strBuffer = Space(lngLength)
'            'get line text
            Call SendMessageStr(.hWnd, EM_GETLINE, lngIndex, ByVal strBuffer)
            about.Caption = strBuffer & vbCrLf
        
        Next
    End With


End Sub



Private Sub Picture1_Click()
Dim url
url = "http://www.facebook.com/" + fb.Text
CreateObject("Wscript.Shell").Run url
End Sub

Private Sub Picture2_Click()
Dim url
url = "http://www.instagram.com/" + insta.Text
CreateObject("Wscript.Shell").Run url
End Sub

Private Sub report_Click()
'PROFILES.ListView.SelectedItem.SubItems(9) = one.Text

'Adodc1.Recordset.AddNew

'Adodc1.Recordset.Fields("Reported").Value = one.Text
'Adodc1.Recordset.update
dialog.Show
dialog.msg.Caption = "Successfully Reported this profile!"
End Sub

Private Sub send_Click()
dialog.Show
dialog.msg.Caption = "Request has been sent"
End Sub
