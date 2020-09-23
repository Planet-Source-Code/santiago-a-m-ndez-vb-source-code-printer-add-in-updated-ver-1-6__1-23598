VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrintSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

' Module Name:  PrintSource.cls
' Author:       Santiago A. Méndez  (Guatemala, C.A.)
' email:        Santiago@InternetDeTelgua.com.gt
' Date:         23-Abr-01
' Description:  The purpose of this Class module is to print the code in the active code pane window of
'               VB IDE.  The way the code is printed is to determine the amount of lines (selected or
'               total module lines) to print, then sends line by line to printer.

Private Type CodeLine                   'IS USED TO CLASSIFY THE CODE LINE PRINTED
    IsSub As Boolean                    'LINE IS A SUB DECLARATION
    IsFunction As Boolean               'LINE IS A FUNCTION DECLARATION
    IsEnumType As Boolean               'LINE IS A TYPE DEFINITION
    IsEndProc As Boolean                'LINE IS A END SUB/FUNCTION
    IsEndType As Boolean                'LINE IS A END TYPE DEFINITION
    IsComment As Boolean                'LINE IS A COMMENT LINE ('/Rem)
End Type

Private m_PrinterLeft As Single
Private m_PrinterRigth As Single
Private m_PrinterTop As Single
Private m_PrinterBottom As Single
Private m_PrinterHeight As Single
Private m_PrinterWidth As Single


Public Property Let PrinterBottom(ByVal Data As Single)
    m_PrinterBottom = Data
End Property

Public Property Get PrinterBottom() As Single
    PrinterBottom = m_PrinterBottom
End Property

Public Property Let PrinterLeft(ByVal Data As Single)
    m_PrinterLeft = Data
End Property

Public Property Get PrinterLeft() As Single
    PrinterLeft = m_PrinterLeft
End Property

Public Property Let PrinterRigth(ByVal Data As Single)
    m_PrinterRigth = Data
End Property

Public Property Get PrinterRigth() As Single
    PrinterRigth = m_PrinterRigth
End Property

Public Property Let PrinterTop(ByVal Data As Single)
    m_PrinterTop = Data
End Property

Public Property Get PrinterTop() As Single
    PrinterTop = m_PrinterTop
End Property

'*--- Function to Print Entire Module or Selected text from current code pane window
Public Sub PrintSourceCode(PrintSelection As Boolean, PrintIndex As Boolean)
    Dim StartLine&, EndLine&, i&, J&, LineaTexto$, sBuffer$, Indice(), Prop As Property
    Dim DeclarationLines&, a&, b&, c&, d&, sBuffer2$, CodeLineType As CodeLine, PadLeft$
    
    Printer.ScaleMode = vbCentimeters         'CM AS UNIT WHEN SETTING CurrentX, CurrentY IN PRINTER OBJECT
    
    'CALCULATE PRINT AREA
    m_PrinterHeight = Printer.ScaleHeight - m_PrinterTop
    m_PrinterWidth = Printer.ScaleWidth - m_PrinterLeft - m_PrinterRigth
    PadLeft = String(m_PrinterLeft / Printer.TextWidth(" "), " ")
    
    'FONT
    Printer.FontName = "Courier New"
    Printer.FontSize = 9
    
    If PrintSelection Then
        VBI.ActiveCodePane.GetSelection StartLine, J, EndLine, i                'GET RANGE OF LINES SELECTED IN WINDOW
    Else
        StartLine = 1
        EndLine = VBI.ActiveCodePane.CodeModule.CountOfLines                    'COUNT OF LINES IN MODULE
    End If
    DeclarationLines = VBI.ActiveCodePane.CodeModule.CountOfDeclarationLines    'DECLARATIONS AT BEGINNING OF MODULE
    
    PrintHeader True
    
    ReDim Preserve Indice(1, 0)
    For i = StartLine To EndLine
        Printer.FontBold = False
        Printer.FontItalic = False
        
        LineaTexto = VBI.ActiveCodePane.CodeModule.Lines(i, 1)   'GET LINE OF TEXT
        sBuffer = Trim(LineaTexto)
        
        CheckLineWidth LineaTexto
        
        'CHECK IF LINE ENDS WITH LINE-CONTINUATION CHARACTER
        If Right(LineaTexto, 1) = "_" Then
            J = i
            Do
                J = J + 1
                sBuffer = VBI.ActiveCodePane.CodeModule.Lines(J, 1)
                
                CheckLineWidth sBuffer
                LineaTexto = LineaTexto & vbCr & sBuffer
                If Right(sBuffer, 1) <> "_" Then Exit Do
            Loop
            i = J
        End If
        sBuffer = Trim(LineaTexto)
        
        'CHECK IF NEXT LINE FITS ON CURRENT PAGE
        If Printer.CurrentY + Printer.TextHeight(LineaTexto) > m_PrinterHeight Then
            Printer.NewPage
            PrintHeader True
        End If
        
        'PRINT TEXT LINE
        PrintTextLine LineaTexto, CodeLineType
        
        If CodeLineType.IsFunction Or CodeLineType.IsSub Then       'IF SUB/FUNCTION ADD TO INDEX
            ReDim Preserve Indice(1, UBound(Indice, 2) + 1)
            
            J = InStr(sBuffer, "(") - 1
            Indice(0, UBound(Indice, 2)) = Left(sBuffer, J)         'SUB/FUNCTION NAME
            Indice(1, UBound(Indice, 2)) = Printer.Page             'PAGE WHERE IS PRINTED
            
        ElseIf CodeLineType.IsEndProc Then                          'IF END OF PROCEDURE PRINT A LINE
            PrintLine
        End If
        
        If i = DeclarationLines Then PrintLine                      'IF END OF DECLARATIONS SECTION PRINT A LINE
    Next
    
    'NOW PRINT INDEX
    If PrintIndex Then
        
        PrintHeaderIndex EndLine - StartLine + 1
        
        'LOOP ARRAY OF PROCEDURES
        For i = 1 To UBound(Indice, 2)
            'CHECK IF NEXT LINE FITS ON CURRENT PAGE
            If Printer.CurrentY + Printer.TextHeight(Indice(0, i)) > m_PrinterHeight Then
                PrintHeaderIndex EndLine - StartLine + 1
            End If
            
            'Printer.CurrentX = 1
            Printer.CurrentX = m_PrinterLeft + 1
            Printer.Print Indice(0, i);             'SUB/FUNCTION NAME
            ImprimirPuntos
            PrintRight Indice(1, i) & Space(3)      '# PAGE
        Next
    End If
    
    Printer.EndDoc
    
End Sub

'*--- Check if line to print is larger than page width, break text line in many lines as needed inserting carriage returns
Private Sub CheckLineWidth(ByRef LineaTexto As String)
    Dim sBuffer$, a&, b&, c&, d&
    
    If Printer.TextWidth(LineaTexto) > m_PrinterWidth Then
        sBuffer = ""
        
        a = Printer.TextWidth(LineaTexto) \ m_PrinterWidth
        
        c = 1
        For b = 1 To a
            Do
                d = d + 1
            Loop Until Printer.TextWidth(Mid(LineaTexto, c, d - 1)) > m_PrinterWidth
            
            sBuffer = sBuffer & Mid(LineaTexto, c, d - 2) & vbCr
            c = c + d - 2
            d = 0
        Next
        sBuffer = sBuffer & Mid(LineaTexto, c)
        LineaTexto = sBuffer
    End If
End Sub

'*--- Print Header of Index Page
Private Sub PrintHeaderIndex(CountLines As Long)
   Dim Prop As Property, sBuffer$
    
    Printer.FontItalic = False
    Printer.NewPage
    PrintHeader False
    
    Printer.CurrentY = m_PrinterTop + 3     '3CM FROM LEFT MARGIN
    
    Printer.FontBold = True
    Printer.CurrentX = m_PrinterLeft + 1
    Printer.Print "Archivo:";
    Printer.FontBold = False
    Printer.CurrentX = m_PrinterLeft + 5
    Printer.Print VBI.ActiveCodePane.CodeModule.Parent.Name;                'MODULE NAME
    Printer.Print
    Printer.FontBold = True
    Printer.CurrentX = m_PrinterLeft + 1
    Printer.Print "Path:";
    Printer.FontBold = False
    Printer.CurrentX = m_PrinterLeft + 5
    
    sBuffer = VBI.ActiveCodePane.CodeModule.Parent.FileNames(1)             'FILE NAME
    
    CheckLineWidth sBuffer
    PrintText sBuffer, False, True
    
    If CountLines > 0 Then
        Printer.FontBold = True
        Printer.CurrentX = m_PrinterLeft + 1
        Printer.Print "Lines of Code:";
        Printer.FontBold = False
        Printer.CurrentX = m_PrinterLeft + 5
        Printer.Print Format(CountLines, "###,##0")
        Printer.Print
    End If
    Printer.Print
    
    On Error Resume Next
    'PRINT PROPERTIES OF CURRENT MODULE ONLY IF IT IS A CLASS MODULE OR MODULE
    If VBI.ActiveCodePane.CodeModule.Parent.Type <= vbext_ct_ClassModule + vbext_ct_ClassModule Then
        For Each Prop In VBI.ActiveCodePane.CodeModule.Parent.Properties
            Printer.CurrentX = m_PrinterLeft + 1
            Printer.Print Prop.Name;        'PROPERTY NAME
            Printer.CurrentX = m_PrinterLeft + 7
            Printer.Print Prop.Value;       'PROPERTY VALUE
            Printer.Print
        Next
    End If
    On Error GoTo 0
    
    Printer.Print       '2 BLANK LINES
    Printer.Print
    
    Printer.FontBold = True
    Printer.FontUnderline = True
    Printer.CurrentX = m_PrinterLeft + 1
    Printer.Print "Procedure";
    PrintRight "Page #"
    Printer.Print
    Printer.FontBold = False
    Printer.FontUnderline = False
End Sub

'*--- Print a line of dots (index page)
Private Sub ImprimirPuntos()
    Dim Fin!
    Fin = m_PrinterWidth - 1.5
    
    Printer.Print " ";
    Do
        Printer.Print ".";
    Loop Until Printer.CurrentX >= Fin
    
    Printer.Print " ";
End Sub

'*--- Prints Text Right Aligned
Private Sub PrintRight(Texto As String)
    Printer.CurrentX = m_PrinterWidth - Printer.TextWidth(Texto) + m_PrinterLeft
    Printer.Print Texto
End Sub

'*--- Prints Page Header, optional prints page number
Private Sub PrintHeader(PrintPageNumber As Boolean)
    Dim sBuffer$, i%
    
    'PRINT DATE, TIME,MODULE NAME AND OPTIONAL PAGE NUMBER
    
    sBuffer = VBI.ActiveCodePane.CodeModule.Parent.Name
    If PrintPageNumber Then sBuffer = sBuffer & "-" & Printer.Page
    sBuffer = Format(Now, "dd-mmm-yyyy HH:mm AM/PM") & Space(10) & sBuffer
    
    With Printer
        .CurrentX = m_PrinterLeft
        .CurrentY = m_PrinterTop
        .FontBold = True
        .FontItalic = True
        PrintRight sBuffer
        
        i = .FontSize
        .FontSize = 2                   'PRINT DOUBLE LINE
        Printer.Line (m_PrinterLeft, .CurrentY)-(m_PrinterWidth + m_PrinterLeft, .CurrentY)
        Printer.Print
        Printer.Line (m_PrinterLeft, .CurrentY)-(m_PrinterWidth + m_PrinterLeft, .CurrentY)
        .FontSize = i \ 2
        Printer.Print
        .FontSize = i
        .FontBold = False
        .FontItalic = False
    End With
End Sub

'*--- Print Line of text formatted according to the text of line
Private Sub PrintTextLine(ByVal LineaTexto As String, ByRef RetTypeLine As CodeLine)
    Dim sBuffer$, J%, LineCode As CodeLine
    
    RetTypeLine = LineCode          'TO CLEAR RetTypeLine VALUES
    
    sBuffer = Trim(LineaTexto)
    If Left(sBuffer, 11) = "Private Sub" Or _
        Left(sBuffer, 10) = "Friend Sub" Or _
        Left(sBuffer, 10) = "Static Sub" Or _
        Left(sBuffer, 10) = "Public Sub" Or _
        Left(sBuffer, 3) = "Sub" Then                           'SUB
            RetTypeLine.IsSub = True
    
    ElseIf Left(sBuffer, 16) = "Private Function" Or _
        Left(sBuffer, 15) = "Public Function" Or _
        Left(sBuffer, 15) = "Friend Function" Or _
        Left(sBuffer, 15) = "Static Function" Or _
        Left(sBuffer, 8) = "Function" Then                      'FUNCTION
        
            RetTypeLine.IsFunction = True
    
    ElseIf Left(sBuffer, 20) Like "Private Property [LGS]et" Or _
        Left(sBuffer, 19) Like "Public Property [LGS]et" Or _
        Left(sBuffer, 19) Like "Friend Property [LGS]et" Or _
        Left(sBuffer, 16) Like "Property [LGS]et" Then          'PROPERTY GET/LET/SET
        
        RetTypeLine.IsFunction = True
    
    ElseIf Left(sBuffer, 11) = "Public Type" Or _
            Left(sBuffer, 12) = "Private Type" Then             'TYPE DEFINITION
            RetTypeLine.IsEnumType = True
    
    ElseIf Left(sBuffer, 8) = "End Type" Then                   'END TYPE DEFINITION
            RetTypeLine.IsEndType = True
            
    ElseIf Left(sBuffer, 7) = "End Sub" Or _
            Left(sBuffer, 12) = "End Function" Or _
            Left(sBuffer, 12) = "End Property" Then            'END PROCEDURE DEFINITION
            RetTypeLine.IsEndProc = True
    
    ElseIf Left(sBuffer, 1) = "'" Or Left(sBuffer, 3) = "Rem" Then      'COMMENT
            RetTypeLine.IsComment = True
    End If
    
    'PRINT TEXT CODE LINE
    
    Printer.ForeColor = QBColor(0)             'BLACK COLOR
    
    'IF LINE CONTAINS COMMENTS, BUT IS NOT A ENTIRE COMMENT LINE
    If InStr(sBuffer, " '") Or InStr(sBuffer, "Rem ") And Not RetTypeLine.IsComment Then
        J = InStr(LineaTexto, " '")
        If J = 0 Then J = InStr(LineaTexto, "Rem")
        
        sBuffer = Mid(LineaTexto, 1, J - 1)
        
        'PRINT CODE WITHOUT COMMENT
        Printer.FontItalic = False
        Printer.FontBold = RetTypeLine.IsEndProc Or RetTypeLine.IsEndType Or RetTypeLine.IsEnumType Or _
                            RetTypeLine.IsFunction Or RetTypeLine.IsSub
        PrintText sBuffer, True
        sBuffer = Mid(LineaTexto, J)
        
        'NOW PRINT COMMENT
        Printer.FontBold = True
        Printer.FontItalic = True
        Printer.ForeColor = QBColor(8)             'GRAY COLOR
        PrintText sBuffer, False, True
    Else
        'PRINT ENTIRE LINE
        Printer.FontBold = RetTypeLine.IsComment Or RetTypeLine.IsEndProc Or RetTypeLine.IsEndType Or _
                        RetTypeLine.IsEnumType Or RetTypeLine.IsFunction Or RetTypeLine.IsSub
        Printer.FontItalic = RetTypeLine.IsComment
        If RetTypeLine.IsComment Then Printer.ForeColor = QBColor(8)                        'COMMENTS IN GRAY COLOR
        PrintText LineaTexto, False
    End If
End Sub

'*--- Print Lines of Text
Private Sub PrintText(ByVal Texto As String, NoBreakLine As Boolean, Optional ContinueLine As Boolean)
    Dim ArrayLines() As String, i%
    
    If Texto = "" Then
        Printer.Print
    Else
        ArrayLines = Split(Texto, vbCr)
        
        If ContinueLine Then
            ContinueLine = False
        Else
            Printer.CurrentX = m_PrinterLeft
        End If
        
        For i = 0 To UBound(ArrayLines) - 1
            Printer.Print ArrayLines(i)
            Printer.CurrentX = m_PrinterLeft
        Next
        If NoBreakLine Then
            Printer.Print ArrayLines(i);
        Else
            Printer.Print ArrayLines(i)
        End If
    End If
End Sub

'*--- Print a line in current printer line
Private Sub PrintLine()
    Dim X As Single, Y As Single, FSize%
    
    With Printer
        FSize = .FontSize
        .FontSize = FSize \ 2
        .FontSize = FSize
        Printer.Line (m_PrinterLeft, .CurrentY)-(m_PrinterWidth + m_PrinterLeft, .CurrentY)
        .FontSize = FSize \ 2
        Printer.Print
        .FontSize = FSize
    End With
End Sub
