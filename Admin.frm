VERSION 5.00
Begin VB.Form Admin 
   Caption         =   "-_-"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10590
   ForeColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Reported Profiles"
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   15
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Profile Reports"
      BeginProperty Font 
         Name            =   "Exo 2 Medium"
         Size            =   15
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   735
      Left            =   6120
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   735
      Left            =   1800
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "//ADMIN"
      BeginProperty Font 
         Name            =   "Aldrich"
         Size            =   38.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1215
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport1.Show

End Sub
