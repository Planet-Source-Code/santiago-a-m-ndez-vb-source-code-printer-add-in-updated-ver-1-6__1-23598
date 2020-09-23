VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Source Code By Santiago Méndez"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text"
      ToolTipText     =   "Margin in cm."
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Text            =   "Text"
      ToolTipText     =   "Margin in cm."
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Text            =   "Text"
      ToolTipText     =   "Margin in cm."
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Text            =   "Text"
      ToolTipText     =   "Margin in cm."
      Top             =   2640
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      ItemData        =   "frmPrint.frx":000C
      Left            =   1560
      List            =   "frmPrint.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Options"
      Height          =   1095
      Left            =   2400
      TabIndex        =   14
      Top             =   720
      Width           =   2295
      Begin VB.CheckBox Check1 
         Caption         =   "Print Index page"
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Range :"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   1935
      Begin VB.OptionButton OptionB 
         Caption         =   "Selection"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptionB 
         Caption         =   "Current Module"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CommandButton CmdSetup 
      Caption         =   "Setup..."
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Note: Margins are set from printer print area. Not from paper borders."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3480
      TabIndex        =   20
      ToolTipText     =   "Note: Margins are set from printer print area. Not from paper borders."
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4680
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   4680
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Image Image1 
      Height          =   1290
      Left            =   1440
      Picture         =   "frmPrint.frx":003E
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1020
   End
   Begin VB.Label Label6 
      Caption         =   "Bottom:"
      Height          =   255
      Left            =   1920
      TabIndex        =   19
      Top             =   4335
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Right :"
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   3255
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Left :"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3735
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Top:"
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   2655
      Width           =   615
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   4680
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4680
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Print Quality :"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2070
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4680
      Y1              =   4815
      Y2              =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   4680
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label LblPrinter 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Printer : "
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Module Name:  frmPrint.frm
' Author:       Santiago A. Méndez  (Guatemala, C.A.)
' email:        Santiago@InternetDeTelgua.com.gt
' Date:         23-Abr-01
' Description:  The purpose of this form is to bring the user the choice to define printer settings,
'               print an index page, print the entire module or to print only the selected text

Public Sub Display()
    Dim StartLine&, EndLine&, StartCol&, EndCol&
    
    'GET RANGE OF LINES SELECTED IN WINDOW
    VBI.ActiveCodePane.GetSelection StartLine, StartCol, EndLine, EndCol
    
    
    LblPrinter.Caption = Printer.DeviceName
    
    'ENABLE OR DISABLE Range "Selection" OPTION BUTTON
    OptionB(0).Enabled = StartCol <> EndCol Or StartLine <> EndLine
    Me.Show vbModal
    
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim PrintCode As New PrintSource

    MousePointer = vbHourglass
    Screen.MousePointer = vbHourglass
    
    'SET PRINT QUALITY & MARGINS
    Printer.PrintQuality = Combo.ItemData(Combo.ListIndex)
    
    PrintCode.PrinterTop = Text(0)
    PrintCode.PrinterLeft = Text(1)
    PrintCode.PrinterRigth = Text(2)
    PrintCode.PrinterBottom = Text(3)
    
    If OptionB(0) Then
        'PRINT SELECTION
        PrintCode.PrintSourceCode True, Check1.Value = vbChecked
    Else
        'PRINT ENTIRE MODULE
        PrintCode.PrintSourceCode False, Check1.Value = vbChecked
    End If
    MousePointer = vbDefault
    
    'MARGINS
    SaveSetting Me.Name, "Settings", "PrintTop", Text(0)
    SaveSetting Me.Name, "Settings", "PrintLeft", Text(1)
    SaveSetting Me.Name, "Settings", "PrintRight", Text(2)
    SaveSetting Me.Name, "Settings", "PrintBottom", Text(3)
    
    'PRINT QUALITY
    SaveSetting Me.Name, "Settings", "PrintQuality", Combo.ItemData(Combo.ListIndex)
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdSetup_Click()
    If ShowPrinter(Me, PD_PRINTSETUP) Then UpdateCombo Printer.PrintQuality
End Sub

Private Sub UpdateCombo(PQuality As Long)
    Dim i%
    
    If PQuality <> Combo.ItemData(Combo.ListIndex) Then
        Select Case Printer.PrintQuality
            Case vbPRPQDraft
                Combo.ListIndex = 0
            Case vbPRPQLow
                Combo.ListIndex = 1
            Case vbPRPQMedium
                Combo.ListIndex = 2
            Case vbPRPQHigh
                Combo.ListIndex = 3
            Case Else
                For i = 0 To Combo.ListCount - 1
                    If Combo.ItemData(i) = PQuality Then
                        Combo.ListIndex = i
                        Exit For
                    End If
                Next
                If i = Combo.ListCount Then
                    Combo.AddItem PQuality
                    Combo.ItemData(Combo.NewIndex) = PQuality
                    Combo.ListIndex = Combo.NewIndex
                End If
        End Select
    End If
End Sub

Private Sub Combo_Click()
    Printer.PrintQuality = Combo.ItemData(Combo.ListIndex)
End Sub

Private Sub Form_Load()
    'RETRIEVE FORM'S LAST POSITION
    With Me
        .Left = GetSetting(Me.Name, "Settings", Me.Name & "MainLeft", .Left)
        .Top = GetSetting(Me.Name, "Settings", Me.Name & "MainTop", .Top)
    End With
    
    'MARGINS
    Text(0) = FormatNumber(GetSetting(Me.Name, "Settings", "PrintTop", 0), 2)
    Text(1) = FormatNumber(GetSetting(Me.Name, "Settings", "PrintLeft", 0), 2)
    Text(2) = FormatNumber(GetSetting(Me.Name, "Settings", "PrintRight", 0), 2)
    Text(3) = FormatNumber(GetSetting(Me.Name, "Settings", "PrintBottom", 0), 2)
    
    'PRINT QUALITY
    Combo.ListIndex = 0
    UpdateCombo FormatNumber(GetSetting(Me.Name, "Settings", "PrintQuality", vbPRPQDraft), 2)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'SAVE FORM'S POSITION
    With Me
        SaveSetting Me.Name, "Settings", Me.Name & "MainLeft", .Left
        SaveSetting Me.Name, "Settings", Me.Name & "MainTop", .Top
    End With
    
    Set frmPrint = Nothing
End Sub

Private Sub Text_GotFocus(Index As Integer)
    With Text(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text_LostFocus(Index As Integer)
    With Text(Index)
        .Text = FormatNumber(.Text, 2)
    End With
End Sub

Public Function FormatNumber(Numero As String, Decimales As Integer) As String
    FormatNumber = Format(Numero, "###,###,##0." & String(Decimales, "0"))
End Function

