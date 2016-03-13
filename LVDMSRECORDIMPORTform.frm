VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
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
   Icon            =   "LVDMSRECORDIMPORTform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   10815
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
      TabHeight       =   520
      TabCaption(0)   =   "Server"
      TabPicture(0)   =   "LVDMSRECORDIMPORTform.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFolderName(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPassword"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUserName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblSiteName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblServerURL"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "btnTestConnection"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DMSPassword"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DMSUserName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DMSSiteName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DMSServerURL"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CommonDialog1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CommonDialogHdr"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CommonDialogRt"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CommonDialogColHd"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CommonDialogDat"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DMSFolderName"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Printer/General"
      TabPicture(1)   =   "LVDMSRECORDIMPORTform.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmPrinterGeneral(1)"
      Tab(1).Control(1)=   "frmPrinterGeneral(0)"
      Tab(1).Control(2)=   "frmPrinterLabel"
      Tab(1).Control(3)=   "frmBarCode"
      Tab(1).Control(4)=   "frmCaption"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Data"
      TabPicture(2)   =   "LVDMSRECORDIMPORTform.frx":0182
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmFields"
      Tab(2).Control(1)=   "frmGeneral"
      Tab(2).Control(2)=   "frmData"
      Tab(2).Control(3)=   "frmData2"
      Tab(2).ControlCount=   4
      Begin VB.Frame frmPrinterGeneral 
         Height          =   4335
         Index           =   1
         Left            =   -74760
         TabIndex        =   101
         Top             =   360
         Width           =   10455
         Begin VB.TextBox DataFile 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            TabIndex        =   32
            Top             =   600
            Width           =   3375
         End
         Begin VB.TextBox ImageRoot 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            TabIndex        =   34
            Top             =   1080
            Width           =   3375
         End
         Begin VB.CheckBox LabelPrinter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Label Printer"
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
            Left            =   120
            TabIndex        =   38
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox BCHeight 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   41
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox BCWidth 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2760
            TabIndex        =   42
            Top             =   2400
            Width           =   735
         End
         Begin VB.TextBox BCCaption 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   45
            Top             =   3840
            Width           =   4095
         End
         Begin VB.CheckBox RunHidden 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Run Hidden"
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
            Left            =   4080
            TabIndex        =   40
            Top             =   2040
            Width           =   1935
         End
         Begin VB.CommandButton brwColHedFle 
            Appearance      =   0  'Flat
            Caption         =   "Browse"
            Height          =   375
            Left            =   9120
            TabIndex        =   37
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox ColumnHeaderFile 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   5640
            TabIndex        =   36
            Top             =   1560
            Width           =   3375
         End
         Begin VB.Frame grmGenPrint 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Printer"
            ForeColor       =   &H80000006&
            Height          =   1575
            Left            =   120
            TabIndex        =   109
            Top             =   120
            Width           =   3495
            Begin VB.CommandButton btnPtrMinus 
               Appearance      =   0  'Flat
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   30
               Top             =   1320
               Width           =   255
            End
            Begin VB.CommandButton btnPtrPlus 
               Caption         =   "+"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   1320
               Width           =   255
            End
            Begin ComctlLib.ListView ListViewPrint 
               Height          =   975
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   1720
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   327682
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   2
               BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Printer"
                  Object.Width           =   2672
               EndProperty
               BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
                  SubItemIndex    =   1
                  Key             =   ""
                  Object.Tag             =   ""
                  Text            =   "Folder"
                  Object.Width           =   1993
               EndProperty
            End
         End
         Begin VB.CommandButton btnBrseDatFle 
            Appearance      =   0  'Flat
            Caption         =   "Browse"
            Height          =   375
            Index           =   1
            Left            =   9120
            TabIndex        =   33
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox TableName 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000003&
            Height          =   285
            Index           =   1
            Left            =   5640
            TabIndex        =   31
            Text            =   "Data source or file table name"
            Top             =   240
            Width           =   4695
         End
         Begin VB.TextBox RecordFilter 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   2040
            TabIndex        =   43
            Top             =   2880
            Width           =   4095
         End
         Begin VB.CommandButton btnBrwsDocRoot 
            Appearance      =   0  'Flat
            Caption         =   "Browse"
            Height          =   375
            Index           =   1
            Left            =   9120
            TabIndex        =   35
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox DeleteDataFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Delete Data File"
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
            Left            =   1920
            TabIndex        =   39
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Frame frmDefaults 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Defaults"
            ForeColor       =   &H80000008&
            Height          =   2295
            Left            =   6240
            TabIndex        =   103
            Top             =   1920
            Width           =   4095
            Begin VB.TextBox DefaultFolderName 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   50
               Top             =   1200
               Width           =   2175
            End
            Begin VB.CheckBox DefaultLaunchWorkFlow 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Caption         =   "Launch Workflow"
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
               Left            =   2040
               TabIndex        =   49
               Top             =   720
               Width           =   1935
            End
            Begin VB.CheckBox DefaultMakeBarCode 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Caption         =   "Make Bar Code"
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
               Left            =   2040
               TabIndex        =   47
               Top             =   360
               Width           =   1815
            End
            Begin VB.CheckBox DefaultDeleteImage 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Caption         =   "Delete Image"
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
               Left            =   120
               TabIndex        =   48
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox DefaultImageExtension 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1800
               TabIndex        =   51
               Top             =   1800
               Width           =   2175
            End
            Begin VB.CheckBox DefaultReprint 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Caption         =   "RePrint Bar Code"
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
               Left            =   120
               TabIndex        =   46
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label lblDefaultFolderName 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
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
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   114
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label lblDflExt 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Caption         =   "Image Extension:"
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
               Left            =   120
               TabIndex        =   104
               Top             =   1800
               Width           =   1695
            End
         End
         Begin VB.TextBox Margin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   44
            Top             =   3360
            Width           =   4095
         End
         Begin VB.Label lblImageRoot 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Image Root:"
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
            Left            =   3720
            TabIndex        =   115
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblBCHeight 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Height:"
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
            Left            =   120
            TabIndex        =   113
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label lblBCWidth 
            Caption         =   "Width:"
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
            Left            =   2040
            TabIndex        =   112
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label lblColHdrFle 
            Caption         =   "Column Header File:"
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
            Left            =   3720
            TabIndex        =   110
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label lblDatFle 
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
            Index           =   1
            Left            =   3720
            TabIndex        =   108
            Top             =   600
            Width           =   615
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
            Index           =   1
            Left            =   3720
            TabIndex        =   107
            Top             =   240
            Width           =   1575
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
            Index           =   1
            Left            =   120
            TabIndex        =   106
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label lblCap 
            Caption         =   "Bar Code Caption:"
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
            Index           =   1
            Left            =   120
            TabIndex        =   105
            Top             =   3840
            Width           =   1695
         End
         Begin VB.Label lblMargin 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Margin:"
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
            Left            =   120
            TabIndex        =   102
            Top             =   3360
            Width           =   855
         End
      End
      Begin VB.TextBox DMSFolderName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   56
         Top             =   2400
         Width           =   8535
      End
      Begin MSComDlg.CommonDialog CommonDialogDat 
         Left            =   6120
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialogColHd 
         Left            =   5400
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialogRt 
         Left            =   4800
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialogHdr 
         Left            =   4080
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3360
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame frmPrinterGeneral 
         Caption         =   "General"
         Height          =   2055
         Index           =   0
         Left            =   -74880
         TabIndex        =   86
         Top             =   360
         Width           =   3975
         Begin VB.CheckBox PrintBarCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Print Barcode"
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
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox PrinterName 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   2
            Top             =   600
            Width           =   2415
         End
         Begin VB.ComboBox PaperBin 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   3
            Top             =   960
            Width           =   2415
         End
         Begin VB.ComboBox PaperSize 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   4
            Top             =   1320
            Width           =   2415
         End
         Begin VB.ComboBox Orientation 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   5
            Top             =   1680
            Width           =   2415
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
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   88
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
            TabIndex        =   87
            Top             =   1680
            Width           =   1335
         End
      End
      Begin VB.Frame frmPrinterLabel 
         Caption         =   "Label"
         Height          =   2295
         Left            =   -74880
         TabIndex        =   85
         Top             =   2400
         Width           =   3975
         Begin VB.CheckBox LabelSheet 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Label Sheet"
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
            Height          =   210
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox LabelSheetLabelWidth 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   11
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox LabelSheetLabelHeight 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   12
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox LabelSheetVerticalSpacing 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   14
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox LabelSheetHorizontalSpacing 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   13
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox LabelSheetMarginLeft 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   10
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox LabelSheetMarginTop 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   9
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox LabelSheetRows 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3000
            TabIndex        =   8
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox LabelSheetColumns 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1200
            TabIndex        =   7
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblLblShtHght 
            Caption         =   "Height:"
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
            Left            =   1920
            TabIndex        =   131
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lblLblShtWdth 
            Caption         =   "Width:"
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
            TabIndex        =   130
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lblLblShtVrtlSpc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Vertical Spacing:"
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
            Left            =   120
            TabIndex        =   121
            Top             =   1920
            Width           =   1575
         End
         Begin VB.Label lblLblShtHrztlSpc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Horizontal Spacing:"
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
            Left            =   120
            TabIndex        =   120
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label lblLblShtLftMrn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Left Margin:"
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
            Left            =   1920
            TabIndex        =   119
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblLblShtTopMrn 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Top Margin:"
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
            Left            =   120
            TabIndex        =   118
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblLblShtRws 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Rows:"
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
            Left            =   1920
            TabIndex        =   117
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblLblShtCol 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Columns:"
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
            TabIndex        =   116
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.TextBox DMSServerURL 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   52
         Top             =   480
         Width           =   8535
      End
      Begin VB.TextBox DMSSiteName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   53
         Top             =   960
         Width           =   8535
      End
      Begin VB.TextBox DMSUserName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   54
         Top             =   1440
         Width           =   8535
      End
      Begin VB.TextBox DMSPassword 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   55
         Top             =   1920
         Width           =   8535
      End
      Begin VB.CommandButton btnTestConnection 
         Appearance      =   0  'Flat
         Caption         =   "Test Connection"
         Height          =   495
         Left            =   1800
         TabIndex        =   57
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Frame frmBarCode 
         Caption         =   "Bar Code"
         Height          =   2175
         Left            =   -70800
         TabIndex        =   82
         Top             =   360
         Width           =   6495
         Begin VB.TextBox BarcodeHeight 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            TabIndex        =   18
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox BarCodeWidth 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   17
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox BarCodeLeft 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            TabIndex        =   16
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox BarCodeTop 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   15
            Top             =   360
            Width           =   615
         End
         Begin VB.ComboBox BarCodeType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2880
            TabIndex        =   19
            Top             =   1320
            Width           =   3375
         End
         Begin VB.TextBox BarCodeValue 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2880
            TabIndex        =   20
            Top             =   1680
            Width           =   3375
         End
         Begin VB.Label lblBrCdHght 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Barcode Height:"
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
            Left            =   2880
            TabIndex        =   125
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblBrCdWdth 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Barcode Width:"
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
            TabIndex        =   124
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblBrCdLft 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Barcode Left:"
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
            Left            =   2880
            TabIndex        =   123
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblBrCdTop 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Barcode Top:"
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
            TabIndex        =   122
            Top             =   360
            Width           =   1335
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
            Left            =   240
            TabIndex        =   84
            Top             =   1320
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
            Left            =   240
            TabIndex        =   83
            Top             =   1680
            Width           =   1815
         End
      End
      Begin VB.Frame frmCaption 
         Caption         =   "Captions"
         Height          =   2175
         Left            =   -70800
         TabIndex        =   78
         Top             =   2520
         Width           =   6495
         Begin VB.TextBox CaptionText 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   25
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox CaptionHeight 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            TabIndex        =   24
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox CaptionWidth 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   23
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox CaptionLeft 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4560
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox CaptionTop 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox CaptionFontName 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2520
            Sorted          =   -1  'True
            TabIndex        =   26
            Top             =   1320
            Width           =   3855
         End
         Begin VB.ComboBox CaptionFontSize 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2520
            TabIndex        =   27
            Top             =   1680
            Width           =   3855
         End
         Begin VB.Label lblCptnHght 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Caption Height:"
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
            Left            =   2880
            TabIndex        =   129
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblCptnWdth 
            Caption         =   "Caption Width:"
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
            TabIndex        =   128
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lblCptnLft 
            Caption         =   "Caption Left:"
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
            Left            =   2880
            TabIndex        =   127
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblCptnTop 
            Caption         =   "Caption Top:"
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
            TabIndex        =   126
            Top             =   240
            Width           =   1215
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
            Index           =   0
            Left            =   360
            TabIndex        =   81
            Top             =   960
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
            TabIndex        =   80
            Top             =   1320
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
            TabIndex        =   79
            Top             =   1680
            Width           =   1215
         End
      End
      Begin VB.Frame frmFields 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "Fields"
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   -72120
         TabIndex        =   99
         Top             =   2880
         Width           =   7815
         Begin VB.CommandButton btnPlus 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1440
            Width           =   255
         End
         Begin ComctlLib.ListView ListViewData93 
            Height          =   1215
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   7575
            _ExtentX        =   13361
            _ExtentY        =   2143
            View            =   3
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   5
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Type"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Source"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Value"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "DMS Field"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Purpose"
               Object.Width           =   1729
            EndProperty
         End
         Begin VB.CommandButton btnMinus 
            Appearance      =   0  'Flat
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   720
            TabIndex        =   72
            Top             =   1440
            Width           =   255
         End
      End
      Begin VB.Frame frmGeneral 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "General"
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   -74880
         TabIndex        =   96
         Top             =   2880
         Width           =   2655
         Begin VB.CheckBox DeleteDocument 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Delete Documents"
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
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CheckBox LaunchWorkFlow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Launch Workflow"
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
            Index           =   0
            Left            =   120
            TabIndex        =   68
            Top             =   1080
            Width           =   1935
         End
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
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Width           =   1935
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
            Left            =   1320
            TabIndex        =   66
            Text            =   "0"
            Top             =   240
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
            Left            =   1920
            TabIndex        =   98
            Top             =   240
            Width           =   615
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
            TabIndex        =   97
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame frmData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Data Connection Details"
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   -74880
         TabIndex        =   73
         Top             =   360
         Width           =   10575
         Begin VB.TextBox DocumentExtension 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   136
            Top             =   2040
            Width           =   495
         End
         Begin VB.CommandButton btnBrwsColHdrFle 
            Appearance      =   0  'Flat
            Caption         =   "Browse"
            Height          =   375
            Left            =   9240
            TabIndex        =   64
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox ColumnHeaderFile 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2040
            TabIndex        =   62
            Top             =   1680
            Width           =   6975
         End
         Begin VB.CommandButton btnBrwsDocRoot 
            Appearance      =   0  'Flat
            Caption         =   "Browse"
            Height          =   375
            Index           =   0
            Left            =   9240
            TabIndex        =   65
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox DocumentRoot 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   4320
            TabIndex        =   63
            Text            =   "Path to Documents to Import with records"
            Top             =   2040
            Width           =   4695
         End
         Begin VB.TextBox RecordFilter 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2040
            TabIndex        =   60
            Top             =   960
            Width           =   8415
         End
         Begin VB.TextBox TableName 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   2040
            TabIndex        =   59
            Text            =   "Data source or file table name"
            Top             =   600
            Width           =   8415
         End
         Begin VB.TextBox SelectString 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   58
            Text            =   "Where clause for your DMS Record Set"
            Top             =   240
            Width           =   8415
         End
         Begin VB.TextBox ConnectionString 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            TabIndex        =   61
            Text            =   "Connection string to data source"
            Top             =   1320
            Width           =   8415
         End
         Begin VB.Label lblDocumentExtension 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Document Extension:"
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
            Left            =   120
            TabIndex        =   137
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label lblColHdrFle9 
            Caption         =   "Column Header File:"
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
            TabIndex        =   111
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label lblDocumentRoot 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "Document Root:"
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
            Index           =   0
            Left            =   2760
            TabIndex        =   100
            Top             =   2040
            Width           =   1575
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
            Index           =   0
            Left            =   120
            TabIndex        =   77
            Top             =   960
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
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   600
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
            Left            =   120
            TabIndex        =   75
            Top             =   240
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
            Left            =   120
            TabIndex        =   74
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame frmData2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Fields"
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   -74880
         TabIndex        =   132
         Top             =   360
         Width           =   10575
         Begin ComctlLib.ListView ListViewData9 
            Height          =   2055
            Left            =   120
            TabIndex        =   133
            Top             =   240
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   3625
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   5
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Type"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Source"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Value"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "DMS Field"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Purpose"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.CommandButton btnMinus9 
            Appearance      =   0  'Flat
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   135
            Top             =   2400
            Width           =   255
         End
         Begin VB.CommandButton btnPlus9 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   255
            Left            =   480
            TabIndex        =   134
            Top             =   2400
            Width           =   255
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
         Left            =   120
         TabIndex        =   95
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
         Left            =   120
         TabIndex        =   94
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
         Left            =   120
         TabIndex        =   93
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
         Left            =   120
         TabIndex        =   92
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
         Index           =   0
         Left            =   120
         TabIndex        =   91
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
         Caption         =   "Save"
      End
      Begin VB.Menu SaveFileAs 
         Caption         =   "&Save As"
      End
      Begin VB.Menu ExitApplication 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu cfgver 
      Caption         =   "&Config Version"
      Begin VB.Menu SetVersion9 
         Caption         =   "&Version 9.0"
      End
      Begin VB.Menu SetVersion93 
         Caption         =   "&Version 9.3"
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

Public ConnectionStringClicked As Boolean
Public DocumentRootClicked As Boolean
Public TableNameClicked As Boolean
Public SelectStringClicked As Boolean
Public TableNameClicked0 As Boolean
Public TableNameClicked1 As Boolean

Private Sub brwColHedFle_Click()
CommonDialogColHd.CancelError = True
    
    On Error GoTo OpenFileErrHandler
    CommonDialogColHd.ShowOpen
    ColumnHeaderFile(1).Text = CommonDialogColHd.FileName
    
    Exit Sub
OpenFileErrHandler:
    Select Case Err.Number
    Case 32755 'cancel button
        Exit Sub
    Case Else
        MsgBox "Somthing went wrong! Err: " & Err.Number & " " & Err.Description
    End Select

End Sub

Private Sub btnBrseDatFle_Click(index As Integer)
    CommonDialogDat.CancelError = True
    
    On Error GoTo OpenFileErrHandler
    CommonDialogDat.ShowOpen
    DataFile.Text = CommonDialogDat.FileName
    
    Exit Sub
OpenFileErrHandler:
    Select Case Err.Number
    Case 32755 'cancel button
        Exit Sub
    Case Else
        MsgBox "Somthing went wrong! Err: " & Err.Number & " " & Err.Description
    End Select
End Sub

Private Sub btnBrwsColHdrFle_Click()
    CommonDialogHdr.CancelError = True
    
    On Error GoTo OpenFileErrHandler
    CommonDialogHdr.ShowOpen
    ColumnHeaderFile(0).Text = CommonDialogHdr.FileName
    
    Exit Sub
OpenFileErrHandler:
    Select Case Err.Number
    Case 32755 'cancel button
        Exit Sub
    Case Else
        MsgBox "Somthing went wrong! Err: " & Err.Number & " " & Err.Description
    End Select
End Sub

Private Sub btnBrwsDocRoot_Click(index As Integer)
    CommonDialogRt.CancelError = True
    
    On Error GoTo OpenFileErrHandler
    CommonDialogRt.ShowOpen
    If index = 0 Then
        DocumentRoot(0).Text = CommonDialogRt.FileName
    ElseIf index = 1 Then
        ImageRoot.Text = CommonDialogRt.FileName
    End If
    
    Exit Sub
OpenFileErrHandler:
    Select Case Err.Number
    Case 32755 'cancel button
        Exit Sub
    Case Else
        MsgBox "Somthing went wrong! Err: " & Err.Number & " " & Err.Description
    End Select
End Sub

Private Sub DMSPassword_LostFocus()
    setTestConnectionButton
End Sub

Private Sub DMSServerURL_LostFocus()
    setTestConnectionButton
End Sub

Private Sub DMSUserName_LostFocus()
    setTestConnectionButton
End Sub

Private Sub LabelSheet_Click()
    
    If RecordImport.LabelSheet.value = 0 Then
    RecordImport.LabelSheetColumns.Enabled = False
    RecordImport.LabelSheetColumns.BackColor = &H80000000
    RecordImport.LabelSheetRows.Enabled = False
    RecordImport.LabelSheetRows.BackColor = &H80000000
    RecordImport.LabelSheetMarginTop.Enabled = False
    RecordImport.LabelSheetMarginTop.BackColor = &H80000000
    RecordImport.LabelSheetMarginLeft.Enabled = False
    RecordImport.LabelSheetMarginLeft.BackColor = &H80000000
    RecordImport.LabelSheetLabelWidth.Enabled = False
    RecordImport.LabelSheetLabelWidth.BackColor = &H80000000
    RecordImport.LabelSheetLabelHeight.Enabled = False
    RecordImport.LabelSheetLabelHeight.BackColor = &H80000000
    RecordImport.LabelSheetHorizontalSpacing.Enabled = False
    RecordImport.LabelSheetHorizontalSpacing.BackColor = &H80000000
    RecordImport.LabelSheetVerticalSpacing.Enabled = False
    RecordImport.LabelSheetVerticalSpacing.BackColor = &H80000000
    End If
    
    If RecordImport.LabelSheet.value = 1 Then
    RecordImport.LabelSheetColumns.Enabled = True
    RecordImport.LabelSheetColumns.BackColor = &H80000005
    RecordImport.LabelSheetRows.Enabled = True
    RecordImport.LabelSheetRows.BackColor = &H80000005
    RecordImport.LabelSheetMarginTop.Enabled = True
    RecordImport.LabelSheetMarginTop.BackColor = &H80000005
    RecordImport.LabelSheetMarginLeft.Enabled = True
    RecordImport.LabelSheetMarginLeft.BackColor = &H80000005
    RecordImport.LabelSheetLabelWidth.Enabled = True
    RecordImport.LabelSheetLabelWidth.BackColor = &H80000005
    RecordImport.LabelSheetLabelHeight.Enabled = True
    RecordImport.LabelSheetLabelHeight.BackColor = &H80000005
    RecordImport.LabelSheetHorizontalSpacing.Enabled = True
    RecordImport.LabelSheetHorizontalSpacing.BackColor = &H80000005
    RecordImport.LabelSheetVerticalSpacing.Enabled = True
    RecordImport.LabelSheetVerticalSpacing.BackColor = &H80000005
    End If
    
End Sub




Private Sub PrintBarCode_Click()
    If RecordImport.PrintBarCode.value = 1 Then
        RecordImport.LabelSheet.Enabled = True
        RecordImport.PrinterName.Enabled = True
        RecordImport.PaperBin.Enabled = True
        RecordImport.PaperSize.Enabled = True
        RecordImport.Orientation.Enabled = True
        RecordImport.BarcodeHeight.Enabled = True
        RecordImport.BarcodeHeight.BackColor = &H80000005
        RecordImport.BarCodeLeft.Enabled = True
        RecordImport.BarCodeLeft.BackColor = &H80000005
        RecordImport.BarCodeTop.Enabled = True
        RecordImport.BarCodeTop.BackColor = &H80000005
        RecordImport.BarCodeType.Enabled = True
        RecordImport.BarCodeType.BackColor = &H80000005
        RecordImport.BarCodeValue.Enabled = True
        RecordImport.BarCodeValue.BackColor = &H80000005
        RecordImport.BarCodeWidth.Enabled = True
        RecordImport.BarCodeWidth.BackColor = &H80000005
        RecordImport.CaptionLeft.Enabled = True
        RecordImport.CaptionLeft.BackColor = &H80000005
        RecordImport.CaptionTop.Enabled = True
        RecordImport.CaptionTop.BackColor = &H80000005
        RecordImport.CaptionHeight.Enabled = True
        RecordImport.CaptionHeight.BackColor = &H80000005
        RecordImport.CaptionWidth.Enabled = True
        RecordImport.CaptionWidth.BackColor = &H80000005
        RecordImport.CaptionText.Enabled = True
        RecordImport.CaptionText.BackColor = &H80000005
        
    End If
    
    If RecordImport.PrintBarCode.value = 0 Then
        RecordImport.LabelSheet.Enabled = False
        RecordImport.PrinterName.Enabled = False
        RecordImport.PaperBin.Enabled = False
        RecordImport.PaperSize.Enabled = False
        RecordImport.Orientation.Enabled = False
        RecordImport.BarcodeHeight.Enabled = False
        RecordImport.BarcodeHeight.BackColor = &H80000000
        RecordImport.BarCodeLeft.Enabled = False
        RecordImport.BarCodeLeft.BackColor = &H80000000
        RecordImport.BarCodeTop.Enabled = False
        RecordImport.BarCodeTop.BackColor = &H80000000
        RecordImport.BarCodeType.Enabled = False
        RecordImport.BarCodeValue.Enabled = False
        RecordImport.BarCodeValue.BackColor = &H80000000
        RecordImport.BarCodeWidth.Enabled = False
        RecordImport.BarCodeWidth.BackColor = &H80000000
        RecordImport.CaptionLeft.Enabled = False
        RecordImport.CaptionLeft.BackColor = &H80000000
        RecordImport.CaptionTop.Enabled = False
        RecordImport.CaptionTop.BackColor = &H80000000
        RecordImport.CaptionHeight.Enabled = False
        RecordImport.CaptionHeight.BackColor = &H80000000
        RecordImport.CaptionWidth.Enabled = False
        RecordImport.CaptionWidth.BackColor = &H80000000
        RecordImport.CaptionText.Enabled = False
        RecordImport.CaptionText.BackColor = &H80000000
        
        If RecordImport.LabelSheet.value = 1 Then
            RecordImport.LabelSheet.value = 0
        End If
    End If

End Sub

Private Sub PrinterName_Click()
    Dim ep As ESCPrint3.InterFace
    Dim PI As ESCPrint3.PaperInfo
    Dim p As Collection
    Dim x As Long
    Dim bin As Collection
    Dim pSize As Collection

    RecordImport.PaperBin.Clear
    RecordImport.PaperSize.Clear
    
    Set ep = New ESCPrint3.InterFace
    ep.PrinterDevice = RecordImport.PrinterName.Text 'needed for paperbin & papersize else error. setting printer
    Set bin = ep.GetPaperBinList
    Set pSize = ep.GetPaperSizeList
  
    For x = 1 To bin.count
        Set PI = bin(x)
        RecordImport.PaperBin.AddItem PI.Name & " - " & PI.Number
        RecordImport.PaperBin.ItemData(RecordImport.PaperBin.NewIndex) = PI.Number
    Next x
    
    For x = 1 To pSize.count
        Set PI = pSize(x)
        RecordImport.PaperSize.AddItem PI.Name & " - " & PI.Number
        RecordImport.PaperSize.ItemData(RecordImport.PaperSize.NewIndex) = PI.Number
    Next x
    
End Sub

Private Sub LabelPrinter_Click()
    If RecordImport.LabelPrinter.value = 0 Then
        RecordImport.BCCaption.Enabled = False
        RecordImport.BCHeight.Enabled = False
        RecordImport.BCWidth.Enabled = False
    End If
End Sub

Public Sub setTestConnectionButton()
    Dim tempArrayurl()     As String
    Dim temparrayusr() As String
    Dim temparraypw() As String
    Dim lngWordCount    As Long
    Dim lngCharCount    As Long
    Dim lngCharCountS    As Long
    Dim urlcount As Long
    Dim usrcount As Long
    Dim pwcount As Long
    
    tempArrayurl = Split(Trim$(DMSServerURL.Text), " ")
    temparrayusr = Split(Trim$(DMSUserName.Text), " ")
    temparraypw = Split(Trim$(DMSPassword.Text), " ")
    
    
    urlcount = Len(DMSServerURL.Text)
    usrcount = Len(DMSUserName.Text)
    pwcount = Len(DMSPassword.Text)

    If urlcount > 0 And usrcount > 0 And pwcount > 0 Then
       btnTestConnection.Enabled = True
    End If

End Sub

Private Sub ConnectionString_GotFocus()
    If ConnectionStringClicked = False Then
    RecordImport.ConnectionString.Text = ""
    RecordImport.ConnectionString.ForeColor = &H80000008
    ConnectionStringClicked = True
    End If
End Sub

Private Sub DocumentRoot_GotFocus(index As Integer)
    If DocumentRootClicked = False Then
        RecordImport.DocumentRoot(index).Text = ""
        RecordImport.DocumentRoot(index).ForeColor = &H80000008
        DocumentRootClicked = True
    End If
End Sub

Private Sub SelectString_GotFocus()
    If SelectStringClicked = False Then
        RecordImport.SelectString.Text = ""
        RecordImport.SelectString.ForeColor = &H80000008
        SelectStringClicked = True
    End If
End Sub

Private Sub TableName_GotFocus(index As Integer)
    If TableNameClicked0 = False And index = 0 Then
        RecordImport.TableName(index).Text = ""
        RecordImport.TableName(index).ForeColor = &H80000008
        TableNameClicked0 = True
    End If
    If TableNameClicked1 = False And index = 1 Then
        RecordImport.TableName(index).Text = ""
        RecordImport.TableName(index).ForeColor = &H80000008
        TableNameClicked1 = True
    End If
End Sub

