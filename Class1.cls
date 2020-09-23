VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintSourceCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

' Module Name:  ClassAddin  (filename: Class1.cls)
' Author:       Santiago A. Méndez  (Guatemala, C.A.)
' email:        Santiago@InternetDeTelgua.com.gt
' Date:         23-Abr-01
' Description:  This module was copied from MSDN Library Visual Studio 6.0
'               and I made some Little changes to the code.
'               The purpose of functions in this module are:
'               1. Add a menu item in Tools Menu to give the user a way to print Code

' VBI is assigned a pointer to the current IDE's
' VBA object which is later passed as a parameter
' to the OnConnection procedure. It's retained
' because you need it later for disconnecting the
' add-in. Other procedures may have a need for it
' as well.

Private Modulo As Office.CommandBarControl
'Private Seleccion As Office.CommandBarControl

' This will be set to the new command bar control.

Private WithEvents PrintModule As CommandBarEvents
Attribute PrintModule.VB_VarHelpID = -1
'Private WithEvents PrintSelection As CommandBarEvents

' This is the event handling procedure for
' the click event of the new command bar control.

Implements IDTExtensibility

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
'must be here in order to the add-in work correctly
End Sub

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, _
    ByVal ConnectMode As vbide.vbext_ConnectMode, ByVal AddInInst As vbide.AddIn, custom() As Variant)

    ' Save the current instance of Visual Basic.
    Set VBI = VBInst
    ' Add a menu command to the Tools menu.
    
    Set Modulo = VBI.CommandBars("Tools").Controls.Add(before:=4)
    Modulo.Caption = "Print Current Module"
    Modulo.BeginGroup = True
    
    ' Connect the event handler to receive the
    ' events for the new command bar control.
    Set PrintModule = VBI.Events.CommandBarEvents(Modulo)
End Sub

Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbide.vbext_DisconnectMode, custom() As Variant)
    ' Delete the new menu command from the Tools
    Modulo.Delete
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
'must be here in order to the add-in work correctly
End Sub

Private Sub PrintModule_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If VBI.ActiveCodePane Is Nothing Then
        MsgBox "You must activate the code window to print"
    Else
        frmPrint.Display
    End If
End Sub
