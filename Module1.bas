Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text

' Module Name:  Module1.bas
' Author:       Santiago A. Méndez  (Guatemala, C.A.)
' email:        Santiago@InternetDeTelgua.com.gt
' Date:         23-Abr-01
' Description:  This module was copied from MSDN Library Visual Studio 6.0
'               and I made some Little changes to the code.
'               The purpose of functions in this module are:
'               1. Add Addin-Project to Add-In Manager List
'               2. Show de Printer Setup Dialog Through API Call


Public VBI As vbide.VBE

Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

' Global constants for Win32 API
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const DM_DUPLEX = &H1000&
Public Const DM_ORIENTATION = &H1&
Public Const DM_PRINTQUALITY = &H400&

Public Const PD_ALLPAGES = &H0
Public Const PD_COLLATE = &H10
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLESETUPTEMPLATE = &H8000
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NONETWORKBUTTON = &H200000
Public Const PD_NOPAGENUMS = &H8
Public Const PD_NOSELECTION = &H4
Public Const PD_NOWARNING = &H80
Public Const PD_PAGENUMS = &H2
Public Const PD_PRINTSETUP = &H40
Public Const PD_PRINTTOFILE = &H20
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_RETURNIC = &H200
Public Const PD_SELECTION = &H1
Public Const PD_SHOWHELP = &H800
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000

Public Const DLG_PRINT = 0
Public Const DLG_PRINTSETUP = 1
    
Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
    End Type


Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
    End Type


Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private m_PrintDlg As PRINTDLG_TYPE, m_DevMode As DEVMODE_TYPE, m_DevName As DEVNAMES_TYPE

'ADD Add-in to Addin Manager
Sub AddToINI()
    Dim rc As Long
    rc = WritePrivateProfileString("Add-Ins32", "PrintSourceCodeAddIn.PrintSourceCode", "0", "VBADDIN.INI")
    MsgBox "Add-in is now entered in VBADDIN.INI file."
End Sub

Public Function ShowPrinter(frmOwner As Form, Optional PrintFlags As Integer) As Boolean
'    Dim m_PrintDlg As m_PrintDlg_TYPE
'    Dim m_devmode As m_devmode_TYPE
'    Dim m_devname As DEVNAMES_TYPE
    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String
    Dim strSetting As String
    
    
    ' Use PrintDialog to get the handle to a memory block with a m_devmode and m_devname structures
    m_PrintDlg.lStructSize = Len(m_PrintDlg)
    m_PrintDlg.hwndOwner = frmOwner.hWnd
    
    m_PrintDlg.Flags = PrintFlags
    
    'Set the current orientation and duplex setting
    m_DevMode.dmDeviceName = Printer.DeviceName
    m_DevMode.dmSize = Len(m_DevMode)
'    m_devmode.dmFields = DM_ORIENTATION Or DM_DUPLEX
'    m_devmode.dmOrientation = Printer.Orientation
'
'
'    On Error Resume Next
'    m_devmode.dmDuplex = Printer.Duplex
'    On Error GoTo 0
    
    'Allocate memory for the initialization hDevMode structure and copy the settings gathered above into this memory
    m_PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(m_DevMode))
    lpDevMode = GlobalLock(m_PrintDlg.hDevMode)

    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, m_DevMode, Len(m_DevMode)
        bReturn = GlobalUnlock(m_PrintDlg.hDevMode)
    End If
    
    'Set the current driver, device, and port name strings
    With m_DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        m_DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With
    
    'Allocate memory for the initial hDevName structure and copy the settings gathered above into this memory
    m_PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(m_DevName))
    lpDevName = GlobalLock(m_PrintDlg.hDevNames)

    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, m_DevName, Len(m_DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If
    
    'Call the print dialog up and let the user make changes

    If PrintDialog(m_PrintDlg) Then
        'First get the m_devname structure.
        lpDevName = GlobalLock(m_PrintDlg.hDevNames)
        CopyMemory m_DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree m_PrintDlg.hDevNames
        
        'Next get the m_devmode structure and set the printer properties appropriately
        
        lpDevMode = GlobalLock(m_PrintDlg.hDevMode)
        CopyMemory m_DevMode, ByVal lpDevMode, Len(m_DevMode)
        bReturn = GlobalUnlock(m_PrintDlg.hDevMode)
        GlobalFree m_PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(m_DevMode.dmDeviceName, InStr(m_DevMode.dmDeviceName, Chr$(0)) - 1))


        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                End If
            Next
        End If
        
        On Error Resume Next
        'Set printer object properties according to selections made by user

        DoEvents

        'SET TO PRINTER OBJECT PROPERTIES MODIFIED BY THE USER
        With Printer
            .ColorMode = m_DevMode.dmColor
            .Copies = m_DevMode.dmCopies
            .PaperBin = m_DevMode.dmDefaultSource
            .Duplex = m_DevMode.dmDuplex
            .Orientation = m_DevMode.dmOrientation
            .PaperSize = m_DevMode.dmPaperSize
            .PrintQuality = m_DevMode.dmPrintQuality
            .Zoom = m_DevMode.dmScale
        End With
        On Error GoTo 0
        ShowPrinter = True
    End If
End Function
