VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form camfrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd5 
      Caption         =   "Browse"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   5400
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   8280
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00404040&
      ForeColor       =   &H8000000A&
      Height          =   4095
      Left            =   480
      ScaleHeight     =   4035
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   2880
         Picture         =   "camfrm.frx":0000
         ScaleHeight     =   1935
         ScaleWidth      =   1935
         TabIndex        =   7
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Smile for the Camera"
         BeginProperty Font 
            Name            =   "Aldrich"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   5295
      End
   End
   Begin VB.Label pic 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   3000
      Width           =   135
   End
End
Attribute VB_Name = "camfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hCap As Long

Private Sub cmd4_Click()
Dim sFileName As String
    Call SendMessage(hCap, WM_CAP_SET_PREVIEW, CLng(False), 0&)
    With CDialog
        .CancelError = True
        .Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
        .Filter = "Bitmap Picture(*.bmp)|*.bmp|JPEG Picture(*.jpg)|*.jpg|All Files|*.*"
        .ShowSave
        sFileName = .fileName










    End With
    Call SendMessage(hCap, WM_CAP_FILE_SAVEDIB, 0&, ByVal CStr(sFileName))
DoFinally:
    Call SendMessage(hCap, WM_CAP_SET_PREVIEW, CLng(True), 0&)
End Sub





Private Sub Cmd3_Click()
Dim temp As Long
If startcap = True Then
temp = SendMessage(hwdc, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
startcap = False
End If
End Sub



Private Sub Cmd1_Click()
hCap = capCreateCaptureWindow("Smile for the Camera", WS_CHILD Or WS_VISIBLE, 0, 0, Picture1.Height, Picture1.Width, Picture1.hWnd, 0)

  If (hwdc <> 0) Then
    temp = SendMessage(hwdc, WM_CAP_DRIVER_CONNECT, 0, 0)
    temp = SendMessage(hwdc, WM_CAP_SET_PREVIEW, 1, 0)
    temp = SendMessage(hwdc, WM_CAP_SET_PREVIEWRATE, 30, 0)
    startcap = True
    Else
    MsgBox ("No Webcam found")
  End If
End Sub







Private Sub Cmd2_Click()
Dim temp As Long
temp = SendMessage(hCap, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
End Sub



Private Sub cmd5_Click()
       
CDialog.Filter = "Picture Files|*.JPG"
CDialog.ShowOpen
pic = CDialog.fileName
Picture1.Picture = LoadPicture(pic)
Rsignup.dis = LoadPicture(pic)
If CDialog.fileName <> "" Then
     Picture1.Picture = LoadPicture(CDialog.fileName)
     Rsignup.dis = LoadPicture(CDialog.fileName)
    Label1.Visible = False
    Picture2.Visible = False
    Me.Hide

End If
End Sub

Private Sub Form_Load()
cmd1.Caption = "Start &Cam"
cmd2.Caption = "&Capture"
cmd3.Caption = "&Close Cam"
cmd4.Caption = "&Save Image"
End Sub

