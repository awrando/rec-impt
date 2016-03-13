Attribute VB_Name = "modMain"
Option Explicit

Public Const VERSION9 = 1
Public Const VERSION93 = 0
Public Const hasConnection = False

Private cTestConnection As Command
Private cMinus As Command
Private cOpenFile As Command
Private cSaveFile As Command
Private cSaveFileAs As Command
Private cExitApplication As Command
Private cBrws As Command
Private cPlus As Command
Private cCreateField As Command
Private cVersion9 As Command
Private cVersion93 As Command
Private cPtrPlus As Command
Private cPtrMinus As Command
Private cPtrOk As Command
Private med As Mediator

Sub Main()
    Set med = New Mediator
    Set cTestConnection = New CommandTestConnection
    Set cPlus = New CommandPlus
    Set cMinus = New CommandMinus
    Set cBrws = New CommandBrowse
    Set cOpenFile = New CommandOpenFile
    Set cSaveFile = New CommandSaveFile
    Set cSaveFileAs = New CommandSaveAsFile
    Set cExitApplication = New CMDExtApp
    Set cCreateField = New CommandCreateField
    Set cVersion9 = New CommandVersion9
    Set cVersion93 = New CommandVersion93
    Set cPtrPlus = New CommandPtrPlus
    Set cPtrMinus = New CommandPtrMinus
    Set cPtrOk = New CommandPtrAdd

    cTestConnection.setMediator med
    cMinus.setMediator med
    cBrws.setMediator med
    cOpenFile.setMediator med
    cSaveFile.setMediator med
    cSaveFileAs.setMediator med
    cExitApplication.setMediator med
    cPlus.setMediator med
    cCreateField.setMediator med
    cVersion9.setMediator med
    cVersion93.setMediator med
    cPtrPlus.setMediator med
    cPtrMinus.setMediator med
    cPtrOk.setMediator med

    ' Change button command @ run-time here!
    med.RegisterTestConnectionCommand cTestConnection
    med.RegisterPlusCommand cPlus
    med.RegisterMinusCommand cMinus
    med.RegisterBrowseCommand cBrws
    med.RegisterOpenFileCommand cOpenFile
    med.RegisterSaveFileCommand cSaveFile
    med.RegisterSaveFileAsCommand cSaveFileAs
    med.RegCMDExtApp cExitApplication
    med.RegisterCreateFieldCommand cCreateField
    med.RegisterVersion9Command cVersion9
    med.RegisterVersion93Command cVersion93
    med.RegisterPtrPlusCommand cPtrPlus
    med.RegisterPtrMinusCommand cPtrMinus
    med.RegisterPtrOKCommand cPtrOk

    med.init
End Sub
