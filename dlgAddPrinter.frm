VERSION 5.00
Begin VB.Form dlgAddPrinter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1335
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPtrFldNm 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   2775
   End
   Begin VB.ComboBox cmbPtrNm 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton btnPtrOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblAdFldNm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Folder Name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblAdPtrNm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Printer name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "dlgAddPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload dlgAddPrinter
End Sub
