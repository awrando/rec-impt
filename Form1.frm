VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "test"
End Sub
