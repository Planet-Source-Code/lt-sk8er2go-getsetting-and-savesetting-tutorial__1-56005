VERSION 5.00
Begin VB.Form CaptionAssign 
   Caption         =   "Assign a Caption"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6420
   LinkTopic       =   "Form2"
   ScaleHeight     =   1020
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "cancel"
      Height          =   615
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "okie dokey"
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "New caption"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "CaptionAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Command1.Caption = CaptionAssign.Text1.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
