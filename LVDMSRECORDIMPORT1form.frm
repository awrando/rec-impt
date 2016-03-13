VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.Ocx"
Begin VB.Form RecordImport 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record Import"
   ClientHeight    =   4785
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8493
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Server"
      TabPicture(0)   =   "LVDMSRECORDIMPORT1form.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CommonDialog1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DMSServerURL"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DMSSiteName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DMSUserName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DMSPassword"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DMSFolderName"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "btnTestConnection"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblServerURL"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblSiteName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblUserName"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblPassword"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblFolderName"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Printer"
      TabPicture(1)   =   "LVDMSRECORDIMPORT1form.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmPrinterGeneral"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frmPrinterLabel"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frmBarCode"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "frmCaption"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Data"
      TabPicture(2)   =   "LVDMSRECORDIMPORT1form.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "frmData"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frmGeneral"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "frmFields"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -71640
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame frmFields 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Fields"
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   3120
         TabIndex        =   62
         Top             =   2640
         Width           =   7575
         Begin ComctlLib.ListView ListView1 
            Height          =   1455
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2566
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            MousePointer    =   4
            NumItems        =   4
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Source Field"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "DMS Field"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Purpose"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Static Value"
               Object.Width           =   2646
            EndProperty
         End
         Begin VB.CommandButton btnMinus 
            Appearance      =   0  'Flat
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   720
            TabIndex        =   64
            Top             =   1680
            Width           =   315
         End
         Begin VB.CommandButton btnPlus 
            Appearance      =   0  'Flat
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   63
            Top             =   1680
            Width           =   315
         End
      End
      Begin VB.Frame frmGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "General"
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   120
         TabIndex        =   57
         Top             =   2640
         Width           =   2895
         Begin VB.CheckBox EndOnError 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Continue On Error"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox KeepLogs 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   59
            Text            =   "0"
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblDays 
            Caption         =   " Days."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   61
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblKeepLogs 
            Caption         =   "Keep Logs:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frmPrinterGeneral 
         Caption         =   "General"
         Height          =   2175
         Left            =   -74760
         TabIndex        =   43
         Top             =   360
         Width           =   3495
         Begin VB.CheckBox PrintBarCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Print Bar Code"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   2055
         End
         Begin VB.ComboBox PrinterName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   47
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox PaperBin 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   46
            Top             =   960
            Width           =   2055
         End
         Begin VB.ComboBox PaperSize 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   45
            Top             =   1320
            Width           =   2055
         End
         Begin VB.ComboBox Orientation 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   44
            Top             =   1680
            Width           =   2055
         End
         Begin VB.Label lblPrinterName 
            Caption         =   "Printer Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblPaperBin 
            Caption         =   "Paper Bin:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblPaperSize 
            Caption         =   "Paper Size:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lblOrientation 
            Caption         =   "Orientation:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1680
            Width           =   1335
         End
      End
      Begin VB.Frame frmPrinterLabel 
         Caption         =   "Label"
         Height          =   2175
         Left            =   -74760
         TabIndex        =   35
         Top             =   2520
         Width           =   3375
         Begin VB.CheckBox LabelSheet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Label Printer"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox cmbLabelDimensions 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   2160
            TabIndex        =   38
            Top             =   720
            Width           =   1095
         End
         Begin VB.ComboBox cmbMargins 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            TabIndex        =   37
            Top             =   1200
            Width           =   2175
         End
         Begin VB.ComboBox cmbPadding 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1080
            TabIndex        =   36
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label lblLabelDimension 
            Caption         =   "Label Dimensions:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label lblLabelSheetMargins 
            Caption         =   "Margins:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblPadding 
            Caption         =   "Padding:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   1680
            Width           =   975
         End
      End
      Begin VB.TextBox DMSServerURL 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73200
         TabIndex        =   34
         Top             =   480
         Width           =   8535
      End
      Begin VB.TextBox DMSSiteName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73200
         TabIndex        =   33
         Top             =   960
         Width           =   8535
      End
      Begin VB.TextBox DMSUserName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73200
         TabIndex        =   32
         Top             =   1440
         Width           =   8535
      End
      Begin VB.TextBox DMSPassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -73200
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   1920
         Width           =   8535
      End
      Begin VB.TextBox DMSFolderName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -73200
         TabIndex        =   30
         Top             =   2400
         Width           =   8535
      End
      Begin VB.CommandButton btnTestConnection 
         Appearance      =   0  'Flat
         Caption         =   "Test Connection"
         Height          =   495
         Left            =   -73200
         TabIndex        =   29
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Frame frmBarCode 
         Caption         =   "Bar Code"
         Height          =   2175
         Left            =   -71160
         TabIndex        =   22
         Top             =   360
         Width           =   6855
         Begin VB.ComboBox cmbBarCodeDimensions 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4680
            TabIndex        =   25
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox BarCodeType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   4680
            TabIndex        =   24
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox BarCodeValue 
            Enabled         =   0   'False
            Height          =   375
            Left            =   4680
            TabIndex        =   23
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label lblBarCodeDimension 
            Caption         =   "Bar Code Dimensions:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   28
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label lblBarCodeType 
            Caption         =   "Bar Code Type:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   27
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label lblBarCodeValue 
            Caption         =   "Bar Code Value:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   26
            Top             =   1320
            Width           =   1815
         End
      End
      Begin VB.Frame frmCaption 
         Caption         =   "Captions"
         Height          =   2175
         Left            =   -71160
         TabIndex        =   13
         Top             =   2520
         Width           =   6855
         Begin VB.ComboBox cmbCaptionDimensions 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            TabIndex        =   17
            Top             =   360
            Width           =   4215
         End
         Begin VB.TextBox CaptionText 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2520
            TabIndex        =   16
            Top             =   720
            Width           =   4215
         End
         Begin VB.ComboBox CaptionFontName 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            TabIndex        =   15
            Top             =   1080
            Width           =   4215
         End
         Begin VB.ComboBox CaptionFontSize 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2520
            TabIndex        =   14
            Top             =   1440
            Width           =   4215
         End
         Begin VB.Label lblCaptionDimensions 
            Caption         =   "Caption Dimensions:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   21
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblCaptionText 
            Caption         =   "Caption Text:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   20
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblFont 
            Caption         =   "Font:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lblFontSize 
            Caption         =   "Font Size:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.Frame frmData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Data Connection Details"
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10575
         Begin VB.TextBox RecordFilter 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1920
            TabIndex        =   7
            Top             =   1320
            Width           =   8535
         End
         Begin VB.TextBox TableName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1920
            TabIndex        =   6
            Text            =   "Data source or file table name"
            Top             =   960
            Width           =   8535
         End
         Begin VB.TextBox SelectString 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1920
            TabIndex        =   5
            Text            =   "Where clause for your DMS Record Set"
            Top             =   600
            Width           =   8535
         End
         Begin VB.TextBox ConnectionString 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1920
            TabIndex        =   4
            Text            =   "Connection string to data source"
            Top             =   1680
            Width           =   8535
         End
         Begin VB.TextBox ColumnHeaderFile 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1920
            TabIndex        =   3
            Top             =   240
            Width           =   7095
         End
         Begin VB.CommandButton btnBrowse 
            Appearance      =   0  'Flat
            Caption         =   "Browse"
            Height          =   375
            Left            =   9240
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblRecordFilter 
            Caption         =   "Record Filter:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label lblTableName 
            Caption         =   "Table Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblSelectString 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "SQLSelect:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblConnectionString 
            Caption         =   "Connection String:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "File:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label lblServerURL 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Server URL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   56
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblSiteName 
         Caption         =   "Site Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   55
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblUserName 
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   54
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   53
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblFolderName 
         Caption         =   "Folder Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   52
         Top             =   2400
         Width           =   1575
      End
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu OpenFile 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu SaveFile 
         Caption         =   "&Save File"
         Shortcut        =   ^S
      End
      Begin VB.Menu ExitApplication 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "RecordImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CLASS RECORDIMPORT

