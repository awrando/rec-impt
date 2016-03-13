VERSION 5.00
Begin VB.Form dlgAdFld 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton rStaticField 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Static Field"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton rMappedField 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Mapped Field"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox cmbFieldPurpose 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox txtDMSFieldName 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton btnCreateField 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label FieldPurpose 
      Caption         =   "Field Purpose"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label DMSFieldName 
      Caption         =   "DMS Field"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "dlgAdFld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload dlgAdFld
End Sub

Private Sub rMappedField_Click(Index As Integer)
    dlgAdFld.Label1.Caption = "Mapped Field"
End Sub

Private Sub rStaticField_Click(Index As Integer)
    dlgAdFld.Label1.Caption = "Static Field"
End Sub
