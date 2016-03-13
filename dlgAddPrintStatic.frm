VERSION 5.00
Begin VB.Form dlgStcAddPrinter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1200
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAddPrinterFolder 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox txtAddPrinterName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblAddPrinterFolder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "FolderName:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblStcAdPtrNm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Printer Name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "dlgStcAddPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
