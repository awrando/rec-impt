VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu OpenMenu 
         Caption         =   "Open"
      End
      Begin VB.Menu RedMenu 
         Caption         =   "Red"
      End
      Begin VB.Menu ExitMenu 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mnuExitCmd As ExitCommand



Private Sub ExitMenu_Click()
mnuExitCmd.Command_Execute
End Sub

Private Sub Form_Load()
Set mnuExitCmd = New ExitCommand
End Sub
