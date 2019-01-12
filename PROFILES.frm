VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form browser 
   Caption         =   "Hashiko Browser"
   ClientHeight    =   9750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20535
      ExtentX         =   36221
      ExtentY         =   19288
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim strLocation As String

'strLocation = "C:\Users\Ashish\Downloads\user-profile\index.html"
'strLocation = "file://" & strLocation

'WebBrowser1.Navigate (strLocation)
End Sub

