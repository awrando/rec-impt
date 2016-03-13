VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mediator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' CLASS MEDIATOR

Private WithEvents btnTestConnection As CommandButton
Attribute btnTestConnection.VB_VarHelpID = -1
Private WithEvents btnPlus As CommandButton
Attribute btnPlus.VB_VarHelpID = -1
Private WithEvents btnPlus9 As CommandButton
Attribute btnPlus9.VB_VarHelpID = -1
Private WithEvents btnMinus9 As CommandButton
Attribute btnMinus9.VB_VarHelpID = -1
Private WithEvents btnMinus As CommandButton
Attribute btnMinus.VB_VarHelpID = -1
Private WithEvents btnBrws As CommandButton
Attribute btnBrws.VB_VarHelpID = -1
Private WithEvents btnPtrPlus As CommandButton
Attribute btnPtrPlus.VB_VarHelpID = -1
Private WithEvents btnPtrMinus As CommandButton
Attribute btnPtrMinus.VB_VarHelpID = -1
Private WithEvents OpenFile As Menu
Attribute OpenFile.VB_VarHelpID = -1
Private WithEvents SaveFile As Menu
Attribute SaveFile.VB_VarHelpID = -1
Private WithEvents SaveFileAs As Menu
Attribute SaveFileAs.VB_VarHelpID = -1
Private WithEvents ExitApplication As Menu
Attribute ExitApplication.VB_VarHelpID = -1
Private WithEvents btnCreate As CommandButton
Attribute btnCreate.VB_VarHelpID = -1
Private WithEvents mVersion9 As Menu
Attribute mVersion9.VB_VarHelpID = -1
Private WithEvents mVersion93 As Menu
Attribute mVersion93.VB_VarHelpID = -1
Private WithEvents btnPtrOK As CommandButton
Attribute btnPtrOK.VB_VarHelpID = -1

Private btnTestConnectionCommand As CommandTestConnection
Private btnPlusCommand As CommandPlus
Private btnPtrPlusCommand As CommandPtrPlus
Private btnPtrMinusCommand As CommandPtrMinus
Private btnMinusCommand As CommandMinus
Private btnBrwsCmd As CommandBrowse
Private menuOpenFileCommand As CommandOpenFile
Private menuSaveFileCommand As CommandSaveFile
Private menuSaveFileAsCommand As CommandSaveAsFile
Private menuExitApplicationCommand As CMDExtApp
Private loadConfigVersionCommand As CommandConfig
Private btnCreateField As CommandCreateField
Private menuVersion9Command As CommandVersion9
Private menuVersion93Command As CommandVersion93
Private btnPtrOKCommand As CommandPtrAdd

Private ep As ESCPrint3.InterFace
Private configModel As ConfigModelA
Private p As Collection
Private css As Boolean

Private v1foldername As String

Public Sub init()
    OpenState
End Sub

Private Sub OpenState()
    Load RecordImport
    RecordImport.Visible = True
    
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
    RecordImport.BarcodeHeight.Enabled = False
    RecordImport.BarcodeHeight.BackColor = &H80000000
    RecordImport.BarcodeLeft.Enabled = False
    RecordImport.BarcodeLeft.BackColor = &H80000000
    RecordImport.BarcodeTop.Enabled = False
    RecordImport.BarcodeTop.BackColor = &H80000000
    RecordImport.BarCodeType.Enabled = False
    RecordImport.BarCodeValue.Enabled = False
    RecordImport.BarCodeValue.BackColor = &H80000000
    RecordImport.BarcodeWidth.Enabled = False
    RecordImport.BarcodeWidth.BackColor = &H80000000
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
    RecordImport.PrinterName.Enabled = False
    RecordImport.SelectString.ForeColor = &H80000003
    RecordImport.TableName(0).ForeColor = &H80000003
    RecordImport.ConnectionString.ForeColor = &H80000003
    RecordImport.DocumentRoot(0).ForeColor = &H80000003
    
    css = False
    
    Dim p As Collection
    Dim x As Long
    
    Set ep = New ESCPrint3.InterFace
    Set btnTestConnection = RecordImport.btnTestConnection
    Set btnPlus = RecordImport.btnPlus
    Set btnPlus9 = RecordImport.btnPlus9
    Set btnMinus9 = RecordImport.btnMinus9
    Set btnPtrPlus = RecordImport.btnPtrPlus
    Set btnPtrMinus = RecordImport.btnPtrMinus
    Set btnMinus = RecordImport.btnMinus
    Set OpenFile = RecordImport.OpenFile
    Set SaveFile = RecordImport.SaveFile
    Set SaveFileAs = RecordImport.SaveFileAs
    Set ExitApplication = RecordImport.ExitApplication
    Set btnCreate = dlgAdFld.btnCreateField
    Set mVersion9 = RecordImport.SetVersion9
    Set mVersion93 = RecordImport.SetVersion93
    Set btnPtrOK = dlgAddPrinter.btnPtrOK
    
    Set p = ep.GetPrinterList 'Collections, load printers
    
    For x = 1 To p.count
    RecordImport.PrinterName.AddItem p(x)
    Next x
    
    Dim i As Integer
    For i = 0 To Screen.FontCount - 1 ' Determine number of fonts.
        RecordImport.CaptionFontName.AddItem Screen.Fonts(i)  ' Put each font into list box.
    Next i
    
    RecordImport.CaptionFontSize.AddItem "8"
    RecordImport.CaptionFontSize.AddItem "9"
    RecordImport.CaptionFontSize.AddItem "10"
    RecordImport.CaptionFontSize.AddItem "11"
    RecordImport.CaptionFontSize.AddItem "12"
    RecordImport.CaptionFontSize.AddItem "14"
    RecordImport.CaptionFontSize.AddItem "16"
    RecordImport.CaptionFontSize.AddItem "18"
    RecordImport.CaptionFontSize.AddItem "20"
    RecordImport.CaptionFontSize.AddItem "22"
    RecordImport.CaptionFontSize.AddItem "24"
    RecordImport.CaptionFontSize.AddItem "26"
    RecordImport.CaptionFontSize.AddItem "28"
    RecordImport.CaptionFontSize.AddItem "30"
    RecordImport.CaptionFontSize.AddItem "48"
    RecordImport.CaptionFontSize.AddItem "72"
    
    RecordImport.BarCodeType.AddItem "LV"
    RecordImport.BarCodeType.AddItem "CUSTOM"
    
    Set configModel = New ConfigModelA
    
    Dim r As REGADP
    Set r = New REGADP
    Dim rVerCfg As String
    rVerCfg = r.RegGetValue(HKEY_LOCAL_MACHINE, _
                            "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\LVDMSRecordImport", _
                            "DisplayName" _
                            )
    
    If rVerCfg <> "" Then
        Select Case rVerCfg
            Case "LVDMS Record Import 9.000"
                configModel.newType VERSION9
            Case "LVDMS Record Import 9.007"
                configModel.newType VERSION93
            Case Default
                configModel.newType VERSION93
        End Select
    End If

    If configModel.getType = VERSION9 Then
        RecordImport.SetVersion9.Checked = True
        RecordImport.SetVersion93.Checked = False
        setVer9GUI
    ElseIf configModel.getType = VERSION93 Then
        RecordImport.SetVersion9.Checked = False
        RecordImport.SetVersion93.Checked = True
        setVer93GUI
    Else
        RecordImport.SetVersion9.Checked = False
        RecordImport.SetVersion93.Checked = True
        setVer93GUI
    End If
    
    RecordImport.Orientation.AddItem "Portrait"
    RecordImport.Orientation.AddItem "LandScape"
    
    ' Need to turn this into a function. Lets refactor later.
    Dim thisListView As ListView
    Dim selectInt As Integer
    
    If configModel.getType = VERSION9 Then
        Set thisListView = RecordImport.ListViewData9
    End If
    
    If configModel.getType = VERSION93 Then
        Set thisListView = RecordImport.ListViewData93
    End If
    
    If thisListView.ListItems.count = 0 Then
        btnMinus.Enabled = False
        btnMinus9.Enabled = False
    End If
    ' End of function this later.
    
    btnTestConnection.Enabled = False
    btnPtrMinus.Enabled = False
    
End Sub

Public Sub ExitCommand()
    End
End Sub

Public Sub BrowseCommand()
    MsgBox "BrowseCommand() called from mediator"
End Sub

Public Sub MinusCommand()
    Dim thisListView As ListView
    Dim selectInt As Integer
    
    If configModel.getType = VERSION9 Then
        Set thisListView = RecordImport.ListViewData9
    End If
    
    If configModel.getType = VERSION93 Then
        Set thisListView = RecordImport.ListViewData93
    End If
    
    selectInt = thisListView.SelectedItem.index
    
    thisListView.ListItems.Remove (selectInt)
    thisListView.Refresh
    
    If thisListView.ListItems.count = 0 Then
        btnMinus.Enabled = False
        btnMinus9.Enabled = False
    End If

End Sub

Public Sub PlusCommand()
    Load dlgAdFld
    
    dlgAdFld.rMappedField.Item(0).value = True
    dlgAdFld.Label1.Caption = "Mapped Field"
        
    dlgAdFld.cmbFieldPurpose.AddItem "Update"
    dlgAdFld.cmbFieldPurpose.AddItem "Lookup"
    dlgAdFld.cmbFieldPurpose.AddItem "Both"

    dlgAdFld.Visible = True
End Sub

Public Sub PtrPlusCommand()
    Dim x As Integer
    
    Load dlgAddPrinter
    dlgAddPrinter.Visible = True
    
    Set p = ep.GetPrinterList 'Collections, load printers
    
    For x = 1 To p.count
    dlgAddPrinter.cmbPtrNm.AddItem p(x)
    Next x
    
    btnPtrMinus.Enabled = True
End Sub

Private Sub addPtrs(p As String, f As String)
    ' Loads printers from file.
    Dim LI As ListItem
    
    Set LI = RecordImport.ListViewPrint.ListItems.Add()
    
    LI.Text = p
    LI.SubItems(1) = f
    
    Set LI = Nothing
End Sub

Public Sub addPtr()
    ' Loads printers from dialog.
    Dim LI As ListItem
    
    Set LI = RecordImport.ListViewPrint.ListItems.Add()
    
    LI.Text = dlgAddPrinter.cmbPtrNm.Text
    LI.SubItems(1) = dlgAddPrinter.txtPtrFldNm.Text
    
    dlgAddPrinter.Visible = False
    Unload dlgAddPrinter
    
    Set LI = Nothing
End Sub

Public Sub PtrMinusCommand()
    RecordImport.ListViewPrint.ListItems.Remove (RecordImport.ListViewPrint.SelectedItem.index)
    
    If RecordImport.ListViewPrint.ListItems.count = 0 Then
        btnPtrMinus.Enabled = False
    End If
End Sub

Public Sub OpenFileCommand()
    RecordImport.CommonDialog1.CancelError = True
    RecordImport.CommonDialog1.Filter = "Config (*.xml)|*.xml"
    
    On Error GoTo OpenFileErrHandler
    RecordImport.CommonDialog1.ShowOpen
    
    configModel.OpenFile RecordImport.CommonDialog1.FileName
    
    If configModel.getType = VERSION93 Then
        Me.Version93Command
    ElseIf configModel.getType = VERSION9 Then
        Me.Version9Command
    Else
        Me.Version9Command
    End If
    
    configModel.OpenFile RecordImport.CommonDialog1.FileName 'Bug fix. After version 93 command gui clears
    
    RecordImport.btnTestConnection.Enabled = False
    
    loadCGFGUI ' works for 9 & 9.3
    loadGUIFields
    
    If configModel.getType = VERSION93 Then
        hardwareChk
        
        If RecordImport.ListViewData93.ListItems.count <> 0 Then
            btnMinus.Enabled = True
        Else: btnMinus.Enabled = False
        End If
        
    End If
    
    'btnchk?
    
    If configModel.getType = VERSION9 Then
        loadGUIPrinters
        
        If RecordImport.ListViewPrint.ListItems.count <> 0 Then
            btnPtrMinus.Enabled = True
        Else: btnPtrMinus.Enabled = False
        End If
        
        If RecordImport.ListViewData9.ListItems.count <> 0 Then
            btnMinus9.Enabled = True
        Else: btnMinus9.Enabled = False
        End If
        
    End If
    
    
    RecordImport.setTestConnectionButton
 
    Exit Sub
OpenFileErrHandler:
    Select Case Err.Number
    Case 32755 'cancel button
        Exit Sub
    Case Else
        MsgBox "Somthing went wrong! Err: " & Err.Number & " " & Err.Description
    End Select

End Sub

Public Sub SaveFileCommand()
    If Not RecordImport.ConnectionStringClicked Then RecordImport.ConnectionString.Text = ""
    If Not RecordImport.DocumentRootClicked Then RecordImport.DocumentRoot(0) = ""
    If Not RecordImport.SelectStringClicked Then RecordImport.SelectString.Text = ""
    If Not RecordImport.TableNameClicked0 Then RecordImport.TableName(0).Text = ""
    If Not RecordImport.TableNameClicked1 Then RecordImport.TableName(1).Text = ""

    configModel.SaveFile
End Sub

Public Sub SaveFileAsCommand()
    If Not RecordImport.ConnectionStringClicked Then RecordImport.ConnectionString.Text = ""
    If Not RecordImport.DocumentRootClicked Then RecordImport.DocumentRoot(0) = ""
    If Not RecordImport.SelectStringClicked Then RecordImport.SelectString.Text = ""
    If Not RecordImport.TableNameClicked0 Then RecordImport.TableName(0).Text = ""
    If Not RecordImport.TableNameClicked1 Then RecordImport.TableName(1).Text = ""

    configModel.SaveFileAs
End Sub

Public Sub TestConnectionCommand()
    Dim Session As LvDmsCore9.ClientSession
    
    Set Session = New LvDmsCore9.ClientSession
    Session.ServerURL = RecordImport.DMSServerURL
    Session.ClientAppName = "LVDMSRecordImportUtil"
    Session.ClientAppVersion = "001"
    Session.User_Login RecordImport.DMSSiteName, _
                       RecordImport.DMSUserName, _
                       RecordImport.DMSPassword, _
                       False
    ' Capture errors
    If Session.Error.Code Then
      MsgBox Session.Error.Code
      Select Case Session.Error.Code
      Case -2147012890
        MsgBox "URL not recognized."
        RecordImport.DMSServerURL.SetFocus
      Case 5
        MsgBox "Username & password combination not recognized."
        RecordImport.DMSUserName.SetFocus
      End Select
      
      css = False
      Exit Sub
    End If
    ' End errors
    
    MsgBox "Connection successful.", vbOKOnly + vbInformation, "Test Connection"
    css = True
    Session.User_Logout False

End Sub

Public Sub LoadConfigCommand()
    MsgBox "load config command is working."
End Sub

Public Sub createFieldCommand()
    Dim LI As ListItem
    Dim index As Long
    Dim thisListView As ListView
    Const fSource = 1
    Const fValue = 2
    Const fDms = 3
    Const fPurpose = 4

    If configModel.getType = VERSION9 Then
        Set LI = RecordImport.ListViewData9.ListItems.Add()
        Set thisListView = RecordImport.ListViewData9
    End If
    
    If configModel.getType = VERSION93 Then
        Set LI = RecordImport.ListViewData93.ListItems.Add()
        Set thisListView = RecordImport.ListViewData93
    End If
    
    index = thisListView.ListItems.count + 1
    
    If dlgAdFld.rMappedField.Item(0).value Then
        LI.Text = "Mapped Field"
        LI.SubItems(fSource) = dlgAdFld.Text1.Text
        LI.SubItems(fDms) = dlgAdFld.txtDMSFieldName.Text
        LI.SubItems(fPurpose) = dlgAdFld.cmbFieldPurpose.Text
    End If
    
    If dlgAdFld.rStaticField.Item(1).value Then
        LI.Text = "Static Field"
        LI.SubItems(fValue) = dlgAdFld.txtDMSFieldName.Text
        LI.SubItems(fDms) = dlgAdFld.cmbFieldPurpose.Text
        LI.SubItems(fPurpose) = dlgAdFld.Text1.Text
    End If
    
    Unload dlgAdFld
    
    btnMinus.Enabled = True
    btnMinus9.Enabled = True
End Sub

Public Sub Version93Command()
    configModel.newType (VERSION93)
    RecordImport.SetVersion9.Checked = False
    RecordImport.SetVersion93.Checked = True
    setVer93GUI
End Sub

Public Sub Version9Command()
    configModel.newType (VERSION9)
    RecordImport.SetVersion9.Checked = True
    RecordImport.SetVersion93.Checked = False
    setVer9GUI
End Sub

Public Sub loadCGFGUI()
    ' This is broken
    Dim count As Integer
    Dim currentCtrlName As String
    Dim currentCtrlVal As String
    
    On Error Resume Next ' continue on if you don't find the control
    'update to catch error, we need toupdate so DMSFolderName(version93) can get it's value
    For count = 1 To configModel.getConfigNode.Children.count
        currentCtrlName = configModel.getConfigNode.Children.GetNodeByIndex(count).Name
        currentCtrlVal = configModel.getConfigNode.Children.GetNodeByIndex(count).value
        RecordImport.Controls(currentCtrlName).Text = currentCtrlVal
    Next
End Sub

Public Sub hardwareChk()
' switch case version
'ver 9 print box
'ver93 print box
    Dim sysP As String
    Dim x As Long
    Dim currentP As String
    Dim configP As String
    
    Set p = ep.GetPrinterList
    currentP = RecordImport.PrinterName.Text
    
    For x = 1 To p.count
        configP = p(x)
        If (StrComp(currentP, configP, vbTextCompare) = 0) Then Exit Sub
    Next x
    MsgBox "Can't find printer from config file installed on system"
End Sub

Public Sub RegCMDExtApp(c As Command)
    Set menuExitApplicationCommand = c
End Sub

Public Sub RegisterSaveFileCommand(c As Command)
    Set menuSaveFileCommand = c
End Sub

Public Sub RegisterSaveFileAsCommand(c As Command)
    Set menuSaveFileAsCommand = c
End Sub

Public Sub RegisterTestConnectionCommand(c As Command)
    Set btnTestConnectionCommand = c
End Sub

Public Sub RegisterMinusCommand(c As Command)
    Set btnMinusCommand = c
End Sub

Public Sub RegisterPlusCommand(c As Command)
    Set btnPlusCommand = c
End Sub

Public Sub RegisterBrowseCommand(c As Command)
    Set btnBrwsCmd = c
End Sub

Public Sub RegisterOpenFileCommand(c As Command)
    Set menuOpenFileCommand = c
End Sub

Public Sub RegisterLoadConfigCommand(c As Command)
    Set loadConfigVersionCommand = c
End Sub

Public Sub RegisterCreateFieldCommand(c As Command)
    Set btnCreateField = c
End Sub

Public Sub RegisterVersion9Command(c As Command)
    Set menuVersion9Command = c
End Sub

Public Sub RegisterVersion93Command(c As Command)
    Set menuVersion93Command = c
End Sub

Public Sub RegisterPtrPlusCommand(c As Command)
    Set btnPtrPlusCommand = c
End Sub

Public Sub RegisterPtrMinusCommand(c As Command)
    Set btnPtrMinusCommand = c
End Sub

Public Sub RegisterPtrOKCommand(c As Command)
    Set btnPtrOKCommand = c
End Sub

Public Sub PrintBarCodeCommand()

End Sub

Public Sub LabelPrinterCommand()

End Sub

Private Sub btnCreate_Click()
    btnCreateField.Command_Execute
End Sub

Public Sub btnTestConnection_Click()
    btnTestConnectionCommand.Command_Execute
End Sub

Public Sub OpenFile_Click()
    menuOpenFileCommand.Command_Execute
End Sub

Public Sub SaveFile_Click()
    menuSaveFileCommand.Command_Execute
End Sub

Public Sub SaveFileAs_Click()
    menuSaveFileAsCommand.Command_Execute
End Sub

Public Sub ExitApplication_Click()
    menuExitApplicationCommand.Command_Execute
End Sub

Public Sub btnBrws_Click()
    btnBrwsCmd.Command_Execute
End Sub

Public Sub btnPlus_Click()
    btnPlusCommand.Command_Execute
End Sub

Public Sub btnPlus9_click()
    btnPlusCommand.Command_Execute
End Sub

Public Sub btnMinus9_Click()
    btnMinusCommand.Command_Execute
End Sub

Public Sub btnPtrPlus_Click()
    btnPtrPlusCommand.Command_Execute
End Sub

Public Sub btnPtrMinus_Click()
    btnPtrMinusCommand.Command_Execute
End Sub

Public Sub btnMinus_Click()
    btnMinusCommand.Command_Execute
End Sub

Private Sub mVersion9_Click()
    menuVersion9Command.Command_Execute
End Sub

Private Sub mVersion93_Click()
    menuVersion93Command.Command_Execute
End Sub

Private Sub btnPtrOK_click()
    btnPtrOKCommand.Command_Execute
End Sub

Private Sub loadGUIPrinters()
    Dim p As ENode
    Dim tempE As ENode
    Dim count As Long
    
    Set p = configModel.getConfigNode.Children.GetNodeByName("Printers")
    
    RecordImport.ListViewPrint.ListItems.Clear
    
    For count = 1 To p.Children.count - 1
        Set tempE = p.Children.GetNodeByIndex(count)
        addPtrs tempE.Children.GetNodeByName("PrinterName").value, _
                tempE.Children.GetNodeByName("FolderName").value
    Next
    
    RecordImport.ListViewPrint.Refresh
End Sub

Private Sub loadGUIFields() ' Needs to be update to update 9 fields list.
    Dim mf As ENode
    Dim sf As ENode
    Dim tempE As ENode
    Dim count As Long
    Dim LstVw As ListView
    
    If configModel.getType = VERSION9 Then
        Set LstVw = RecordImport.ListViewData9
    Else
        Set LstVw = RecordImport.ListViewData93
    End If

    Set mf = configModel.getConfigNode.Children.GetNodeByName("MappedFields")
    Set sf = configModel.getConfigNode.Children.GetNodeByName("StaticFields")
    
    LstVw.ListItems.Clear
    
    For count = 1 To mf.Children.count
        Set tempE = mf.Children.GetNodeByIndex(count)
        addFields "MappedFields", _
                   tempE.Children.GetNodeByName("ADOFieldName").value, _
                   tempE.Children.GetNodeByName("DMSFieldName").value, _
                   tempE.Children.GetNodeByName("FieldPurpose").value
    Next
    
     For count = 1 To sf.Children.count
        Set tempE = sf.Children.GetNodeByIndex(count)
        addFields "StaticFields", _
                  tempE.Children.GetNodeByName("FieldValue").value, _
                  tempE.Children.GetNodeByName("DMSFieldName").value, _
                  tempE.Children.GetNodeByName("FieldPurpose").value
    Next
    
    LstVw.Refresh
End Sub

Private Sub setVer9GUI()
    RecordImport.frmPrinterGeneral(0).Visible = False
    RecordImport.frmPrinterGeneral(1).Visible = True
    RecordImport.frmBarCode.Visible = False
    RecordImport.frmCaption.Visible = False
    RecordImport.frmPrinterLabel.Visible = False
    RecordImport.SSTab1.TabCaption(1) = "General"
    RecordImport.frmData2.Visible = True
    RecordImport.frmGeneral.Visible = False
    RecordImport.frmFields.Visible = False
    RecordImport.frmData.Visible = False
    RecordImport.DMSFolderName.Visible = False
    RecordImport.lblFolderName(0).Visible = False
End Sub

Private Sub setVer93GUI()
    RecordImport.frmPrinterGeneral(0).Visible = True
    RecordImport.frmPrinterGeneral(1).Visible = False
    RecordImport.frmBarCode.Visible = True
    RecordImport.frmCaption.Visible = True
    RecordImport.frmPrinterLabel.Visible = True
    RecordImport.SSTab1.TabVisible(2) = True
    RecordImport.SSTab1.TabCaption(1) = "Printer"
    RecordImport.frmData2.Visible = False
    RecordImport.frmGeneral.Visible = True
    RecordImport.frmFields.Visible = True
    RecordImport.frmData.Visible = True
    RecordImport.DMSFolderName.Visible = True
    RecordImport.lblFolderName(0).Visible = True
End Sub

Private Sub addFields(fType As String, sourceData As String, DMSFieldName As String, dmsFieldPurpose As String)

    Dim LI As ListItem
    Dim LstVw As ListView
    Const fSource = 1
    Const fValue = 2
    Const fDms = 3
    Const fPurpose = 4
    
    If configModel.getType = VERSION9 Then
        Set LstVw = RecordImport.ListViewData9
    Else
        Set LstVw = RecordImport.ListViewData93
    End If

    Set LI = LstVw.ListItems.Add()
    
    If fType = "MappedFields" Then
        LI.Text = "Mapped Field"
        LI.SubItems(fSource) = sourceData
        LI.SubItems(fDms) = DMSFieldName
        LI.SubItems(fPurpose) = dmsFieldPurpose
    ElseIf fType = "StaticFields" Then
        LI.Text = "Mapped Field"
        LI.SubItems(fValue) = sourceData
        LI.SubItems(fDms) = DMSFieldName
        LI.SubItems(fPurpose) = dmsFieldPurpose
    Else
        MsgBox "Looks like you broke our List Item add function. Opps! We've stopped the process."
        Exit Sub
    End If
        
    Set LI = Nothing

End Sub
