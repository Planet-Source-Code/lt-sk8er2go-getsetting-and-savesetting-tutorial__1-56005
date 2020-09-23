VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SettingsTutorial"
   ClientHeight    =   3435
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Restart Application"
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Assign me a caption!!"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   4455
   End
   Begin VB.Menu sett 
      Caption         =   "Setting"
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu get 
         Caption         =   "Get"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Command1_Click()
CaptionAssign.Show
End Sub

Private Sub Command2_Click()
Unload Me
Form1.Show
End Sub

Private Sub Form_Load()
MsgBox "Welcome to my SaveSetting and GetSetting tutorial.  If you have any questions, please email me at LTsK8eR2gO@yahoo.com and/or IM me on AIM on the s/n LTsK8eR2gO"
End Sub

Private Sub get_Click()
Command1.Caption = GetSetting("SettingsTutorial", "Form1", "buttons caption")
End Sub

Private Sub save_Click()
SaveSetting "SettingsTutorial", "Form1", "buttons caption", Command1.Caption
End Sub
