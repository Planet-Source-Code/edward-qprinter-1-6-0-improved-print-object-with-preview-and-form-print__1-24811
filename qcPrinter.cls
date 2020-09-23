VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "qcPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
' ==============================================================
' Project:      Enhance Print object
' Author:       edward moth
' Copyright:    © 2000-2001 qbd software ltd
'
' ==============================================================
' Module:       qcPrinter
' Purpose:      Class holding enhanced Print object
' ==============================================================
' Credits:      Thanks to Jarry of Jacsoft and Herman Liu for
'               their contributions toward improving the
'               functionality of this project.
'               Thanks to Alain for investigating and solving
'               the printer transparency issue.
' ==============================================================

Option Explicit
' ALIGNMENT Enum
Public Enum qePrinterAlign
  eLeft
  eCentre
  eRight
  eJustify
End Enum
' END LINE CHARACTER Enum
Public Enum qePrinterChar
  eNone
  eSpace
  eDash
  eLine
  eOops
End Enum
' PAPERSIZE Enum
Public Enum qePrinterPaperSize
  eA3size = 8
  eA4size = 9
  eA5size = 11
  eB4size = 12
End Enum
' SCALEMODE Enum
Public Enum qePrinterScale
  eTwip = 1
  eInch = 5
  eMillimetre = 6
  eCentimetre = 7
  ePercentage = 11
End Enum
' NEWPAGE Enum
Public Enum qeTextNewPageEnum
  None_np = 0
  Before_np = 1
  After_np = 2
  Both_np = 3
End Enum
' HEADER/FOOTER PAGE Enum
Public Enum qePrinterHeadFootPage
  FirstPage_hf
  LastPage_hf
  OddPage_hf
  EvenPage_hf
  Standard_hf
End Enum
' PRINTTEXT Type Enum
Public Enum qePrinterPrintType
  TextFlow_pt
  Absolute_pt
  Header_pt
  Footer_pt
End Enum

Public Type qtTextPositioning
  Absolute As Boolean
  Left As Single
  Top As Single
  Width As Single
  Height As Single
End Type

Public Type qtTextBorderType
  Top As Single
  Left As Single
  Width As Single
  Height As Single
  Line As Integer
  Color As Long
  Shading As Long
End Type

Private Type qtFontStyle
  Name As String
  Size As Currency
  Color As Long
  Bold As Boolean
  Italic As Boolean
  Underline As Boolean
End Type

Public Enum qePrintOrientation
  qPortrait = 1
  qLandscape = 2
End Enum

Public Enum qePrintOptionFlags
  ShowPrinter_po = 1
  ShowCopies_po = 2
End Enum

Private mvarFoot(4) As Boolean
Private mvarHead(4) As Boolean
Private mvarPageSize As qePrinterPaperSize
Private mvarScaleMode As qePrinterScale
Private mvarPages As Integer
Private mvarMarginLeft As Single
Private mvarMarginRight As Single
Private mvarMarginTop As Single
Private mvarMarginBottom As Single
Private mvarAppName As String
Private mvarOrientation As qePrintOrientation
Private mvarPrintOption As qePrintOptionFlags
Private mvarPrinter As Integer
Private mvarOrientAvail As Boolean

Private qcolText As New qcTextItemCol
Private qcolHF As New qcTextItemCol

' Added: version 1.6
Private sExcludeList As String
Private sFormName As String
Private bAutoHeight As Boolean
Private nPreviewState As Integer
Private sPreviewTop As Single
Private sPreviewLeft As Single
Private sPreviewWidth As Single
Private sPreviewHeight As Single

'**************************
'  This API call is used to overcome a bug in VB6(English) that has been
'  corrected in SP4 (may be earlier).  The bug still affects non-English
'  language versions (French and Spanish for definite) - Thanks Alain
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
'**************************

Public Property Get PrintOptions() As qePrintOptionFlags
PrintOptions = mvarPrintOption
End Property
Public Property Let PrintOptions(ByVal eOptions As qePrintOptionFlags)
mvarPrintOption = eOptions

End Property
Friend Sub PreviewPage(picBox As PictureBox, Page As Integer)
Screen.MousePointer = vbHourglass
GetPageSize
With picBox
  .Width = qPage.Width
  .Height = qPage.Height
End With

PrintText picBox, Page, Page

Screen.MousePointer = vbDefault

End Sub
Private Function FormatText(ByVal TextItem As qcTextItem, _
                            ByRef Target As Object, _
                            ByVal Position As Long, _
                            ByRef ForceSameLine As Boolean) As Long
Dim bTarget As Boolean
Dim bCheck As Boolean
Dim lFPos As Long
Dim sFormat As String
Dim sTemp As String

If TextItem.HeadFoot Then
sFormat = TextItem.HFText
Else
sFormat = TextItem.Text
End If
lFPos = InStr(Position, sFormat, ">")
If lFPos > 0 Then
sFormat = Mid$(sFormat, Position + 1, lFPos - Position - 1)
Else
FormatText = 0
Exit Function
End If


bCheck = True
bTarget = Not IsEmpty(Target)
sFormat = UCase(sFormat)
If Len(sFormat) < 3 Or Left(sFormat, 1) = "/" Then
Select Case sFormat
Case "B"
Printer.Font.Bold = True
If bTarget Then
Target.Font.Bold = Printer.Font.Bold
End If

Case "U"
Printer.Font.Underline = True
If bTarget Then
Target.Font.Underline = Printer.Font.Underline
End If

Case "I"
Printer.Font.Italic = True
If bTarget Then
Target.Font.Italic = Printer.Font.Italic
End If

Case "/B"
Printer.Font.Bold = False
If bTarget Then
Target.Font.Bold = Printer.Font.Bold
End If

Case "/I"
Printer.Font.Italic = False
If bTarget Then
Target.Font.Italic = Printer.Font.Italic
End If

Case "/U"
Printer.Font.Underline = False
If bTarget Then
Target.Font.Underline = Printer.Font.Underline
End If

Case "/FONT"
Printer.Font.Name = TextItem.FontName
If bTarget Then
Target.Font.Name = Printer.Font.Name
End If

Case "/COLOR"
Printer.ForeColor = TextItem.FontColor
If bTarget Then
Target.ForeColor = Printer.ForeColor
End If

Case "/SIZE"
Printer.Font.Size = TextItem.FontSize
If bTarget Then
Target.Font.Size = Printer.Font.Size
End If

Case "/ALIGN"
TextItem.TempAlign = TextItem.Alignment
Case "/LINDENT"
TextItem.TempLIndent = 0
ForceSameLine = True
Case "/RINDENT"
TextItem.TempRIndent = 0
ForceSameLine = True
Case Else
bCheck = False
End Select
ElseIf Left(sFormat, 5) = "FONT=" Then
Printer.Font.Name = Property_Extract(sFormat)
If bTarget Then
Target.Font.Name = Printer.Font.Name
End If

ElseIf Left(sFormat, 6) = "COLOR=" Then
Printer.ForeColor = CLng(Property_Extract(sFormat))
If bTarget Then
Target.ForeColor = Printer.ForeColor
End If

ElseIf Left(sFormat, 5) = "SIZE=" Then
Printer.Font.Size = Val(Property_Extract(sFormat))
If bTarget Then
Target.Font.Size = Printer.Font.Size
End If
ElseIf Left(sFormat, 6) = "ALIGN=" Then
sTemp = UCase(Property_Extract(sFormat))
Select Case sTemp
Case "LEFT"
TextItem.TempAlign = eLeft
Case "RIGHT"
TextItem.TempAlign = eRight
Case "CENTRE", "CENTER"
TextItem.TempAlign = eCentre
Case "JUSTIFY"
TextItem.TempAlign = eJustify
End Select
ElseIf Left(sFormat, 8) = "LINDENT=" Then
TextItem.TempLIndent = ConvertToTwip(TextItem.ScaleMode, Val(Property_Extract(sFormat)))
ForceSameLine = True
ElseIf Left(sFormat, 8) = "RINDENT=" Then
TextItem.TempRIndent = ConvertToTwip(TextItem.ScaleMode, Val(Property_Extract(sFormat)))
ForceSameLine = True
ElseIf Left(sFormat, 5) = "FORCE" Then
ForceSameLine = True
Else
bCheck = False
End If
If Not bCheck Then
FormatText = 0
Else
FormatText = lFPos
End If

End Function




Private Function Property_Extract(ByVal sFormat As String)

Dim lPos As Long
Dim sReturn As String

lPos = InStr(1, sFormat, "=")
If lPos > 0 Then
sReturn = Right(sFormat, Len(sFormat) - lPos)
If Left$(sReturn, 1) = "#" Then
sReturn = ConvertHTMColor(sReturn)
End If
End If
Property_Extract = sReturn

End Function

Private Sub PrintText(Target As Object, _
                      StartPage As Integer, _
                      EndPage As Integer, _
                      Optional bPrint As Boolean = False, _
                      Optional TextType As qePrinterPrintType = TextFlow_pt)

' Procedure prints the text items that fall between StartPage
' and EndPage.  In the case of preview a single page
' The procedure moves through the collection of text items
' and calculates there position on each page

Dim sText As String
Dim lCount As Long
Dim qText As qcTextItem
Dim iCurrentPage As Integer, iItems As Integer, iLine As Integer
Dim sL As Single, sR As Single, sT As Single, sB As Single
Dim sCurY As Single, sCurH As Single

Dim bHF As Boolean

iCurrentPage = StartPage

Select Case TextType
Case qePrinterPrintType.Absolute_pt
If Not bPrint Then
GoTo PrintTextEnd
End If
iItems = 1
Do While iItems <= qcolText.Count
Set qText = qcolText(iItems)


If qText.Absolute And (qText.AbsPage = iCurrentPage Or qText.AbsPage = -1) Then
sL = qText.IndentLeftTwip + qPage.LeftM
sR = qPage.Width - qText.IndentRightTwip - qPage.RightM
sT = qText.TopTwip + qPage.TopM + qText.AdditionalTop
If qText.MaxHeightTwip > 0 Then
If qText.MaxHeightTwip > qPage.Height - sT Then
sB = qPage.Height - qPage.BottomM
Else
sB = qText.MaxHeightTwip + sT
End If
Else
If qText.HeightTwip > qPage.Height - sT Then
sB = qPage.Height - qPage.BottomM
Else
sB = qText.HeightTwip + sT
End If
End If
lCount = PrintItem(qText, Target, 0, iLine, sL, sR, sT, sB, bPrint)
End If
iItems = iItems + 1
Loop
Case qePrinterPrintType.TextFlow_pt
iItems = 1
iCurrentPage = 1
Do While iCurrentPage <= EndPage
If iCurrentPage >= StartPage Then
bPrint = True
End If
GetHeadFoot iCurrentPage, CBool(iCurrentPage = mvarPages)
If bPrint Then
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Absolute_pt
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Header_pt
End If
sCurY = qPage.TopM + qPage.HeaderH
Do While iItems <= qcolText.Count

Set qText = qcolText(iItems)
If qText.Absolute Then
GoTo PrintTextItemFinish
End If
If CBool(qText.NewPage And Before_np) Then
If bPrint Then
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Footer_pt
End If
iCurrentPage = iCurrentPage + 1
If iCurrentPage > EndPage Then
GoTo PrintTextEnd
End If
If iCurrentPage >= StartPage Then
bPrint = True
End If
GetHeadFoot iCurrentPage, CBool(iCurrentPage = mvarPages)
If bPrint Then
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Absolute_pt
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Header_pt
End If
sCurY = qPage.TopM + qPage.HeaderH
End If

sL = qPage.LeftM + qText.IndentLeftTwip
sR = qPage.Width - qText.IndentRightTwip - qPage.RightM
sT = sCurY
If qText.HeightTwip + sCurY > qPage.HFAvailHeight Then
sB = qPage.Height - qPage.BottomM - qPage.FooterH
Else
sB = qText.HeightTwip + sCurY
End If

lCount = 0: iLine = 0
Do
lCount = PrintItem(qText, Target, lCount, iLine, sL, sR, sT, sB, bPrint)
sCurY = Target.CurrentY
If lCount > 0 Then
If bPrint Then
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Footer_pt
End If
iCurrentPage = iCurrentPage + 1
If iCurrentPage > EndPage Then
GoTo PrintTextEnd
End If
If iCurrentPage >= StartPage Then
bPrint = True
End If
GetHeadFoot iCurrentPage, CBool(iCurrentPage = mvarPages)
If bPrint Then
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Absolute_pt
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Header_pt
End If
sCurH = TextItemLineHeight(iItems, iLine)
sCurY = qPage.TopM + qPage.HeaderH
sT = sCurY
If sCurH + sCurY > qPage.HFAvailHeight Then
sB = qPage.Height - qPage.BottomM - qPage.FooterH
Else
sB = sCurH + sCurY
End If
End If
Loop While lCount > 0
If CBool(qText.NewPage And After_np) Then
If bPrint Then
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Footer_pt
End If
iCurrentPage = iCurrentPage + 1
If iCurrentPage > EndPage Then
GoTo PrintTextEnd
End If
If iCurrentPage >= StartPage Then
bPrint = True
End If

GetHeadFoot iCurrentPage, CBool(iCurrentPage = mvarPages)
If bPrint Then
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Absolute_pt
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Header_pt
End If
sCurY = qPage.TopM + qPage.HeaderH

End If

PrintTextItemFinish:
iItems = iItems + 1
Loop
If bPrint Then
PrintText Target, iCurrentPage, iCurrentPage, bPrint, Footer_pt
End If
iCurrentPage = iCurrentPage + 1
PrintTextItemEnd:
Loop

Case qePrinterPrintType.Header_pt
If Not bPrint Then
GoTo PrintTextEnd
End If
If mvarHead(4) Then
Set qText = qcolHF("Head4")
bHF = True
End If

If iCurrentPage Mod 2 = 1 Then
If mvarHead(2) Then
Set qText = qcolHF("Head2")
bHF = True
End If
Else
If mvarHead(3) Then
Set qText = qcolHF("Head3")
bHF = True
End If
End If
If iCurrentPage = 1 And mvarHead(0) = True Then
Set qText = qcolHF("Head0")
bHF = True
End If
If iCurrentPage = mvarPages And mvarHead(1) = True Then
Set qText = qcolHF("Head1")
bHF = True
End If
If bHF Then
qText.SetHFText iCurrentPage, mvarPages, mvarAppName
sL = qPage.LeftM + qText.IndentLeftTwip
sR = qPage.Width - qPage.RightM - qText.IndentRightTwip
sT = qPage.TopM
sB = qText.HeightTwip + qPage.TopM
lCount = PrintItem(qText, Target, 0, iLine, sL, sR, sT, sB, bPrint, Header_pt, iCurrentPage)
End If
Case qePrinterPrintType.Footer_pt
If Not bPrint Then
GoTo PrintTextEnd
End If
If mvarFoot(4) Then
Set qText = qcolHF("Foot4")
bHF = True
End If
If iCurrentPage Mod 2 = 1 Then
If mvarFoot(2) Then
Set qText = qcolHF("Foot2")
bHF = True
End If
Else
If mvarFoot(3) Then
Set qText = qcolHF("Foot3")
bHF = True
End If
End If
If iCurrentPage = 1 And mvarFoot(0) = True Then
Set qText = qcolHF("Foot0")
bHF = True
End If
If iCurrentPage = mvarPages And mvarFoot(1) = True Then
Set qText = qcolHF("Foot1")
bHF = True
End If
If bHF Then
qText.SetHFText iCurrentPage, mvarPages, mvarAppName
sL = qPage.LeftM + qText.IndentLeftTwip
sR = qPage.Width - qPage.RightM - qText.IndentRightTwip
sT = qPage.Height - qText.HeightTwip - qPage.BottomM
sB = qPage.Height - qPage.BottomM
lCount = PrintItem(qText, Target, 0, iLine, sL, sR, sT, sB, bPrint, Footer_pt, iCurrentPage)
End If
If TypeOf Target Is Printer Then
Printer.NewPage
End If
End Select

PrintTextEnd:

Set qText = Nothing
End Sub

Private Function PrintItem(ByVal qText As qcTextItem, _
                           ByRef Target As Object, _
                           ByVal FromPoint As Long, _
                           ByRef LineNumber As Integer, _
                           ByVal LeftMargin As Single, _
                           ByVal RightMargin As Single, _
                           ByVal TopMargin As Single, _
                           ByVal BottomMargin As Single, _
                           ByVal bPrint As Boolean, _
                           Optional TextType As qePrinterPrintType = TextFlow_pt, _
                           Optional PageNum As Integer) As Long

Dim sWidth As Single, sHeight As Single
Dim sSizeX As Single, sSizeY As Single
Dim sText As String
Dim sLine As String, sChar As String, sWord As String
Dim iLines As Integer, lCount As Long, iSpace As Integer
Dim lMarker As Long
Dim sCharWidth As Single, sLineHeight As Single
Dim sCurrentPosY As Single, sCurrentPosX As Single
Dim eCharType As qePrinterChar, eEnd As qePrinterChar
Dim bNewLine As Boolean
Dim bFound As Boolean
Dim iItems As Integer, iCurrentPage As Integer
Dim sCurrentHeight As Single
Dim sIndentL As Single, sIndentR As Single
Dim sTempSize As Currency
Dim bCheck As Boolean, sFormat As String, lFPos As Long
Dim sLineH As Single, sPLineW As Single, sWordH As Single, sWordW As Single
Dim sTLineW As Single, sTWordW As Single
Dim sLStart As Long, sLEnd As Long, sWStart As Single, sWEnd As Single
Dim qFont As qtFontStyle, qAbsFont As qtFontStyle, qAbsPrint As qtFontStyle, sAbsPos As Single, sAbsHeight As Single
Dim qBorder As qtTextBorderType
Dim sTempLIndent As Single, sTempRIndent As Single
Dim sDefaultWidth As Single
Dim bForceSameLine As Boolean

sHeight = BottomMargin - TopMargin
sWidth = RightMargin - LeftMargin
sDefaultWidth = sWidth
  Target.CurrentY = TopMargin
  If FromPoint = 0 And Not qText.Absolute Then
  sCurrentHeight = qText.TopTwip
  Target.CurrentY = Target.CurrentY + qText.TopTwip
  End If
  If FromPoint = 0 Then
  qText.TempAlign = qText.Alignment
  qText.TempLIndent = 0
  qText.TempRIndent = 0
  qText.CurrentLIndent = 0
  qText.CurrentRIndent = 0
  End If
  
  If TextType = Footer_pt Or TextType = Header_pt Then
  sText = qText.HFText
  Else
  sText = qText.Text
  End If
  
  With Target.Font
    .Name = qText.FontName
    .Bold = qText.FontBold
    .Italic = qText.FontItalic
    .Underline = qText.FontUnderline
    If TypeOf Target Is PictureBox Then
    ' Check sizes on printer
    Printer.Font.Name = qText.FontName
    Printer.Font.Size = qText.FontSize
    Printer.Font.Bold = qText.FontBold
    Printer.Font.Italic = qText.FontItalic
    Printer.Font.Underline = qText.FontUnderline
    sTempSize = Printer.Font.Size
    .Size = sTempSize
    Else
    .Size = qText.FontSize
    End If
  End With
  Target.ForeColor = qText.FontColor
  Font_Default Target, qFont, True
  Target.ForeColor = qText.FontColor
  Printer.FontSize = qText.FontSize

  If FromPoint > 0 Then
  lCount = 0
  Do
  lCount = InStr(lCount + 1, sText, "<")
  If lCount > 0 And lCount < FromPoint Then
  lFPos = FormatText(qText, Target, lCount, bForceSameLine)
  If lFPos > 0 Then
  lCount = lFPos
  End If
  End If
  Loop While lCount > 0 And lCount < FromPoint
'  lCount = FromPoint
  Font_Default Target, qFont, True
  bForceSameLine = False
  qText.CurrentLIndent = qText.TempLIndent
  qText.CurrentRIndent = qText.TempRIndent
  sWidth = sDefaultWidth - qText.CurrentLIndent - qText.CurrentRIndent
  End If

  If qText.ShowBorder And bPrint Then
  qBorder.Line = qText.BorderLine
  qBorder.Width = sWidth
  qBorder.Left = LeftMargin
  qBorder.Color = qText.BorderColor
  qBorder.Shading = qText.BorderShading
  If Not qText.Absolute And FromPoint = 0 Then
  qBorder.Height = sHeight - qText.TopTwip
  qBorder.Top = TopMargin + qText.TopTwip
  Else
  qBorder.Height = sHeight
  qBorder.Top = TopMargin
  End If
  Border_Print Target, qBorder
  End If
  
  
  If FromPoint = 0 Then
  FromPoint = 1
  End If
  
  With Target
    sLineHeight = Printer.TextHeight("H")
    sLine = ""
    sWord = ""
    eEnd = eNone
    bNewLine = False
    lCount = FromPoint

    sPLineW = 0: sTLineW = 0: sLStart = FromPoint: sLineH = 0 'sLineHeight
' Loop through text to create lines
  sWordW = 0: sTWordW = 0: sWordH = 0
    Do
' Find suitable line break
      Do
        eCharType = eNone
        sChar = Mid$(sText, lCount, 1)
        bCheck = False
        'Check Formatting
        If sChar = "<" Then
        'lFPos = FormatText(qText, Target, lCount)
        lFPos = FormatText(qText, Target, lCount, bForceSameLine)
        Else
        lFPos = 0
        End If

' CHARACTER CHECK: Look for potential line breaks or where text
'                  width is greater than boundary
        If lFPos = 0 Then
        Select Case sChar
          Case " "
            eCharType = eSpace
            sWEnd = lCount
          Case "-"
            sSizeX = sPLineW + sWordW + Printer.TextWidth(sChar) 'Printer.TextWidth(sLine & sWord & sChar)
            If sSizeX > sWidth Then
              eCharType = eOops
            Else
              eCharType = eDash
            End If
            sWEnd = lCount
          Case vbLf
            sChar = ""
            eCharType = eLine
            sWEnd = lCount
          Case vbCr
            If lCount < Len(sText) Then
              If Mid$(sText, lCount + 1, 1) = vbLf Then
                lCount = lCount + 1
              End If
            End If
            sChar = ""
            eCharType = eLine
            sWEnd = lCount - 1
            sLEnd = lCount - 1
          Case Else
' CHARACTER CHECK: See if addition of character makes line too long
            sSizeX = sPLineW + sWordW + Printer.TextWidth(sChar) ' Printer.TextWidth(sLine & sWord & sChar)
            sWordH = Printer.TextHeight(sChar)
            If sWordH > sLineH Then
            sLineH = sWordH
            End If

            If sSizeX > sWidth Then
              eCharType = eOops
            Else
              sWord = sWord & sChar
              sWordW = sWordW + Printer.TextWidth(sChar)
              sTWordW = sTWordW + Target.TextWidth(sChar)
            End If
        End Select
        Else
        lCount = lFPos
        End If
        lCount = lCount + 1
        
      Loop While lCount <= Len(sText) And eCharType = eNone And Not bForceSameLine
      If bForceSameLine Then
      eCharType = eLine
      sWEnd = lCount
      End If
' LINE SPLIT: Examine potential line break
      If lCount > Len(sText) Then
        eCharType = eLine
        sWEnd = lCount
      End If
      Select Case eCharType
        Case qePrinterChar.eNone
          sLine = sLine & sWord
          eEnd = eLine
      sWStart = lCount
        Case qePrinterChar.eOops
          If eEnd = eNone Then
            sLine = sWord
            sWord = sChar
          Else
            sLine = Trim$(sLine)
            sWord = sWord & sChar
            sWordW = sWordW + Printer.TextWidth(sChar)
            sTWordW = sTWordW + Target.TextWidth(sChar)
          End If
            sLEnd = sWEnd
          bNewLine = True
        Case qePrinterChar.eDash, qePrinterChar.eSpace
          eEnd = eCharType
          sLine = sLine & sWord & sChar
          sPLineW = sPLineW + sWordW + Printer.TextWidth(sChar)
          sTLineW = sTLineW + sTWordW + Target.TextWidth(sChar)
          sWord = ""
          sWordW = 0
          sTWordW = 0
          sLEnd = sWEnd - 1
      sWStart = lCount

        Case qePrinterChar.eLine
          sLine = sLine & sWord
          If sLineH = 0 Then
          sLineH = Printer.TextHeight("H")
          End If
          sPLineW = sPLineW + sWordW
          sTLineW = sTLineW + sTWordW
          sWStart = lCount
          eEnd = eLine
          sWord = ""
          sWordW = 0
          sTWordW = 0
          bNewLine = True
        sLEnd = sWEnd - 1
        sWStart = lCount
      End Select
      If bNewLine Then
      If sCurrentHeight + sLineH > sHeight Then
      lCount = sLStart
      GoTo PrintItemEnd
      End If
      End If

' LINE SPLIT: Add new line if required
      If bNewLine And bPrint Then
      Font_Default Target, qFont, False
        .CurrentX = LeftMargin + qText.CurrentLIndent
' ALIGNMENT: Calculate position of line dependent on alignment setting
        Select Case qText.TempAlign 'qText.Alignment
          Case qePrinterAlign.eLeft
          PrintLine Target, qText, sLStart, sLEnd
'            Target.Print sLine

          Case qePrinterAlign.eCentre
            sSizeX = (sWidth - sTLineW) / 2 + LeftMargin 'qPage.LeftM + sIndentL
            .CurrentX = sSizeX
            PrintLine Target, qText, sLStart, sLEnd
'            Target.Print sLine

          Case qePrinterAlign.eRight
            sSizeX = RightMargin - sTLineW   'qPage.Width - qPage.RightM - sTLineW - sIndentR
            .CurrentX = sSizeX
            PrintLine Target, qText, sLStart, sLEnd

'            Target.Print sLine

          Case qePrinterAlign.eJustify
' ALIGNMENT: Full justification is more complex.  Find spaces
'            and calculate extra spacing required
' NEXT LINE: Check to see if line has an line break
            If eEnd <> eLine Then
              sCharWidth = Target.TextWidth(" ")
              sLine = Trim(sLine)
              iSpace = 0
              lMarker = 0
              Do
                lMarker = InStr(lMarker + 1, sLine, " ")
                If lMarker > 0 Then
                  iSpace = iSpace + 1
                End If
              Loop While lMarker > 0
              sSizeX = sWidth - sTLineW 'Printer.TextWidth(sLine)
              bFound = False
' ALIGNMENT: Check number of spaces and extra size, if too large
'            use character justification as well as word justification
              If iSpace > 0 Then
                If sSizeX \ iSpace > sCharWidth * 3 Then
                  bFound = True
                End If
              Else
                bFound = True
              End If
              If bFound Then
                sSizeY = Len(sLine) - 1 + (iSpace * 2)
                sSizeY = sSizeX / sSizeY
                sSizeX = sSizeY * 3
              Else
                sSizeX = sSizeX / iSpace
                sSizeY = 0
              End If
              PrintLine Target, qText, sLStart, sLEnd, sSizeX, sSizeY
            Else
               PrintLine Target, qText, sLStart, sLEnd
            End If
        End Select
        If qText.UnderlineAll > 0 Then
        sSizeY = .CurrentY
        Target.DrawWidth = qText.UnderlineAll
        Target.Line (LeftMargin + qText.CurrentLIndent, .CurrentY + sLineH)-Step(sWidth, 0), qText.BorderColor
        .CurrentY = sSizeY
        End If
        
      End If
      If bNewLine Then
        If Not bForceSameLine Then
        ' *** Changed 1.6.0
        sCurrentHeight = sCurrentHeight + sLineH + qText.LineSpacingTwip
        LineNumber = LineNumber + 1
        .CurrentY = .CurrentY + sLineH + qText.LineSpacingTwip 'eight ' sHeightDiff
        ' *** End Change 1.6.0
        sLineH = 0 'sLineHeight
        Else
        qText.CurrentLIndent = qText.TempLIndent
        qText.CurrentRIndent = qText.TempRIndent
        sWidth = sDefaultWidth - qText.CurrentLIndent - qText.CurrentRIndent
        bForceSameLine = False
        End If
        sLStart = sWStart
        sLine = ""
        Font_Default Target, qFont, True
        sPLineW = 0
        sTLineW = 0
        eEnd = eNone
        bNewLine = False
        iSpace = 0
      End If
PrintItemEndLine:

    Loop While lCount <= Len(sText)
  End With

lCount = 0
If qText.Underline > 0 Then
Target.DrawWidth = qText.Underline
Target.Line (LeftMargin, BottomMargin)-Step(sDefaultWidth, 0), qText.BorderColor
End If

PrintItemEnd:
PrintItem = lCount

End Function

Private Function TextItemLineHeight(ByVal Item As Integer, ByVal Line As Integer) As Single

Dim iLine As Integer
Dim sHeight As Single
iLine = Line + 1
Do While iLine <= qcolText(Item).Lines
sHeight = sHeight + qcolText(Item).LineHeight(iLine)
iLine = iLine + 1
Loop
TextItemLineHeight = sHeight

End Function


Private Sub PrintLine(ByRef Target As Object, ByVal TextItem As qcTextItem, _
                      ByVal StartPoint As Long, ByVal EndPoint As Long, _
                      Optional SpaceExpand As Single = 0, Optional CharExpand As Single = 0)

Dim iCount As Long, lFPos As Long
Dim bCheck As Boolean
Dim sChar As String
Dim sX As Single, sY As Single
Dim sText As String
Dim bForceTemp As Boolean
  
  iCount = StartPoint
  If TextItem.HeadFoot Then
  sText = TextItem.HFText
  Else
  sText = TextItem.Text
  End If
  If iCount = 0 Then
  iCount = 1
  End If
  Do While iCount <= EndPoint
        bCheck = False
        sChar = Mid$(sText, iCount, 1)
        ' Check for formatting codes
        
        If sChar = "<" Then
        lFPos = FormatText(TextItem, Target, iCount, bForceTemp)
        Else
        lFPos = 0
        End If
        If lFPos = 0 Then
'          If (Asc(sChar) <> 10 And Asc(sChar) <> 13) Then
          sX = Target.CurrentX + Target.TextWidth(sChar)
          sY = Target.CurrentY
          Target.Print sChar
          If sChar = " " Then
          sX = sX + SpaceExpand
          Else
          sX = sX + CharExpand
          End If
          Target.CurrentX = sX
          Target.CurrentY = sY
'          End If
        Else
        iCount = lFPos
        End If
        iCount = iCount + 1
      Loop

End Sub
Private Sub Font_Default(ByRef Target As Object, ByRef qFont As qtFontStyle, ByVal GetFont As Boolean)
With qFont
If GetFont Then
  .Name = Target.Font.Name
  .Size = Target.Font.Size
  .Color = Target.ForeColor
  .Bold = Target.Font.Bold
  .Italic = Target.Font.Italic
  .Underline = Target.Font.Underline
Else
  Target.Font.Name = .Name
  Target.Font.Size = .Size
  Target.ForeColor = .Color
  Target.Font.Bold = .Bold
  Target.Font.Italic = .Italic
  Target.Font.Underline = .Underline
End If
End With
  
End Sub

Private Sub Border_Print(ByRef Target As Object, ByRef Border As qtTextBorderType)
Dim iLine As Integer, sX As Single, sY As Single
With Border

sX = Target.CurrentX
sY = Target.CurrentY
If .Shading <> -1 Then
Target.Line (.Left - 90, .Top)-Step(.Width + 90, .Height), .Shading, BF
End If
If .Line > 0 Then
iLine = Target.DrawWidth
Target.DrawWidth = .Line
Target.Line (.Left - 90, .Top)-Step(.Width + 90, .Height), .Color, B
Target.DrawWidth = iLine
End If
End With
Target.CurrentX = sX
Target.CurrentY = sY
End Sub

Private Sub GetPageSize()
With qPage
  .Width = Printer.ScaleWidth
  .Height = Printer.ScaleHeight
  .LeftM = ConvertToTwip(mvarScaleMode, mvarMarginLeft)
  .RightM = ConvertToTwip(mvarScaleMode, mvarMarginRight)
  .TopM = ConvertToTwip(mvarScaleMode, mvarMarginTop)
  .BottomM = ConvertToTwip(mvarScaleMode, mvarMarginBottom)
  .AvailWidth = .Width - .LeftM - .RightM
  .AvailHeight = .Height - .TopM - .BottomM
End With
End Sub

Public Property Get TextItem() As qcTextItemCol

Set TextItem = qcolText

End Property

Public Property Let TextItem(ByVal vData As qcTextItemCol)
Set qcolText = vData
bPageChange = True
End Property

Public Property Get ItemCount() As Long
ItemCount = qcolText.Count
End Property

Public Property Let MarginBottom(ByVal vData As Single)
If mvarMarginBottom <> vData Then
  bPageChange = True
  mvarMarginBottom = vData
  qPage.BottomM = ConvertToTwip(mvarScaleMode, vData, False)
End If
End Property

Public Property Get MarginBottom() As Single
MarginBottom = mvarMarginBottom
End Property

Public Property Let MarginTop(ByVal vData As Single)
If mvarMarginTop <> vData Then
  bPageChange = True
  mvarMarginTop = vData
  qPage.TopM = ConvertToTwip(mvarScaleMode, vData, False)
End If
End Property

Public Property Get MarginTop() As Single
MarginTop = mvarMarginTop
End Property


Public Property Let MarginRight(ByVal vData As Single)
If mvarMarginRight <> vData Then
  bPageChange = True
  mvarMarginRight = vData
  qPage.RightM = ConvertToTwip(mvarScaleMode, vData)
End If
End Property

Public Property Get MarginRight() As Single
MarginRight = mvarMarginRight
End Property

Public Property Let MarginLeft(ByVal vData As Single)
If mvarMarginLeft <> vData Then
  bPageChange = True
  mvarMarginLeft = vData
  qPage.LeftM = ConvertToTwip(mvarScaleMode, vData)
End If
End Property

Public Property Get MarginLeft() As Single
MarginLeft = mvarMarginLeft
End Property

Public Sub PrintDoc(ByVal StartPage As Integer, _
                    ByVal EndPage As Integer, _
                    Optional Copies As Integer = 1, _
                    Optional Collate As Boolean = True, _
                    Optional PrinterNumber As Integer = 0)

' Procedure sets the Target in PrintText to Printer
' printing the pages defined by StartPage and EndPage
' Check current printer

Dim iStart As Integer
Dim iEnd As Integer
Dim iCopy As Integer
Dim iPage As Integer
Dim lBKMode As Long

' Set printer FillStyle to Transparent
lBKMode = SetBkMode(Printer.hDC, TRANSPARENT)

If PrinterNumber <> mvarPrinter Then
Set Printer = Printers(PrinterNumber)
mvarOrientAvail = Orientation_Check
mvarPrinter = PrinterNumber
bPageChange = True
End If
If Copies < 1 Then
Copies = 1
End If
' # Note about the Printer.Copies property.  Like many of the Printer properties
' and methods, the Copies property is affected by the print driver so may
' have no effect.  I have opted to loop through the number of copies to
' avoid any problems with this.

Screen.MousePointer = vbHourglass
Printer.DrawMode = vbCopyPen
If Collate Or Copies = 1 Then
' This will collate copies: Copy 1: Pages 1, 2, 3 Copy 2: 1, 2, 3 etc.
For iCopy = 1 To Copies
PrintText Printer, StartPage, EndPage
Next iCopy
Else
' This will print multiple page copies: Page 1: Copy 1,2,3 Page 2: Copy 1,2,3 etc.
For iPage = StartPage To EndPage
For iCopy = 1 To Copies
PrintText Printer, iPage, iPage
Next iCopy
Next iPage
End If
Printer.EndDoc
Screen.MousePointer = vbDefault

End Sub

Public Sub Preview()
Dim fShow As Form

' Display the preview form
'Set frmPreview.Document = Me
Set fShow = frmPreview
Set fShow.Document = Me
If nPreviewState = -1 Then
fShow.Move (Screen.Width - fShow.Width) \ 2, (Screen.Height - fShow.Height) \ 2
sPreviewWidth = fShow.Width
sPreviewLeft = fShow.Left
sPreviewTop = fShow.Top
sPreviewHeight = fShow.Height
Else
fShow.Move sPreviewLeft, sPreviewTop, sPreviewWidth, sPreviewHeight
If nPreviewState = 2 Then
fShow.WindowState = vbMaximized
End If
End If
Set fShow.Document = Me
fShow.Show vbModal
nPreviewState = fShow.WindowState
If nPreviewState = 0 Then
sPreviewWidth = fShow.Width
sPreviewLeft = fShow.Left
sPreviewTop = fShow.Top
sPreviewHeight = fShow.Height
End If
Unload fShow
Set fShow = Nothing

End Sub

Public Property Let ScaleMode(ByVal vData As qePrinterScale)
Dim iCount As Integer

If vData <> mvarScaleMode Then
  mvarMarginBottom = ConvertFromTwip(vData, qPage.BottomM)
  mvarMarginTop = ConvertFromTwip(vData, qPage.TopM)
  mvarMarginLeft = ConvertFromTwip(vData, qPage.LeftM)
  mvarMarginRight = ConvertFromTwip(vData, qPage.RightM)
  mvarScaleMode = vData
  iCount = 1
  Do While iCount <= qcolText.Count
    qcolText(iCount).ScaleMode = mvarScaleMode
    iCount = iCount + 1
  Loop
  For iCount = 0 To 4
  qcolHF("Head" & iCount).ScaleMode = mvarScaleMode
  qcolHF("Foot" & iCount).ScaleMode = mvarScaleMode
  Next
  
End If
    
End Property


Public Property Get ScaleMode() As qePrinterScale
ScaleMode = mvarScaleMode
End Property


Public Property Get Pages() As Integer

Dim iCount As Integer
Dim lPages As Integer
Dim sHeight As Single
Dim qText As qcTextItem
Dim iAbsPage As Integer
Dim iLines As Integer
Dim sItemHeight As Single
Dim iLastItem As Single
Dim sStartPos As Single

' Calculate the number of pages only if a change has taken
' place - if not return the last page calculation
If bPageChange Then
  iCount = qcolText.Count
  Do While iCount >= 1 And iLastItem = 0
  If Not qcolText(iCount).Absolute Then
  iLastItem = iCount
  End If
  iCount = iCount - 1
  Loop
  
  lPages = 1
  iCount = 1
  GetPageSize
  GetHeadFoot lPages
  
  Do While iCount <= qcolText.Count
    Set qText = qcolText.Item(iCount)
    If Not qText.Absolute Then
    qText.GetSize 'qPage.LeftM, qPage.RightM
    If CBool(qText.NewPage And Before_np) Then
    lPages = lPages + 1
    GetHeadFoot lPages
    sHeight = 0: sItemHeight = 0
    End If
    If iCount = iLastItem Then
    GetHeadFoot lPages, True
    If qPage.HFAvailHeight < qText.HeightTwip Then
    GetHeadFoot lPages
    End If
    End If
    iLines = 1
    qText.StartPage = 0: qText.EndPage = 0
    sStartPos = sHeight
    sHeight = sHeight + qText.TopTwip
'    qText.PositionStartTwip = sHeight + qPage.HeaderH
    Do While iLines <= qText.Lines
    If sHeight + qText.LineHeight(iLines) > qPage.HFAvailHeight Then
    sHeight = qText.LineHeight(iLines)
    lPages = lPages + 1
    If iLines = 1 Then
    sStartPos = 0
    qText.StartPage = lPages
    End If
    If iCount = iLastItem Then
    GetHeadFoot lPages, True
    If qPage.HFAvailHeight < qText.HeightTwip - sItemHeight Then
    GetHeadFoot lPages
    End If
    End If
    Else
    sHeight = sHeight + qText.LineHeight(iLines)
    If qText.StartPage = 0 Then
    qText.StartPage = lPages
    End If
    End If
    sItemHeight = sItemHeight + qText.LineHeight(iLines)
    iLines = iLines + 1
    Loop
    qText.PositionStartTwip = sStartPos
    qText.EndPage = lPages
    qText.PositionEndTwip = sHeight
    If CBool(qText.NewPage And After_np) Then
    lPages = lPages + 1
    GetHeadFoot lPages
    sHeight = 0
    End If
    Else
    qText.GetSize
    If qText.AbsPage > iAbsPage Then
    iAbsPage = qText.AbsPage
    End If
    End If
    iCount = iCount + 1
  Loop
  If iAbsPage > lPages Then
  lPages = iAbsPage
  End If
  
  mvarPages = lPages
  bPageChange = False
End If

Pages = mvarPages

End Property

Private Sub GetHeadFoot(ByVal Page As Integer, Optional LastPage As Boolean = False)

Dim iHead As Integer
Dim iFoot As Integer

qPage.HeaderH = 0
qPage.FooterH = 0
iHead = -1
iFoot = -1
If mvarHead(4) Then
iHead = 4
End If
If mvarFoot(4) Then
iFoot = 4
End If

If Page Mod 2 = 1 Then
If mvarHead(2) Then
iHead = 2
End If
If mvarFoot(2) Then
iFoot = 2
End If
Else
If mvarHead(3) Then
iHead = 3
End If
If mvarFoot(3) Then
iFoot = 3
End If
End If

If Page = 1 Then
If mvarHead(0) Then
iHead = 0
End If
If mvarFoot(0) Then
iFoot = 0
End If
End If

If LastPage Then
If mvarHead(1) Then
iHead = 1
End If
If mvarFoot(1) Then
iFoot = 1
End If
End If

If iHead >= 0 Then
qcolHF("Head" & iHead).SetHFText Page, mvarPages, mvarAppName
qcolHF("Head" & iHead).GetSize
qPage.HeaderH = qcolHF("Head" & iHead).HeightTwip
End If

If iFoot >= 0 Then
qcolHF("Foot" & iFoot).SetHFText Page, mvarPages, mvarAppName
qcolHF("Foot" & iFoot).GetSize
qPage.FooterH = qcolHF("Foot" & iFoot).HeightTwip
End If

qPage.HFAvailHeight = qPage.AvailHeight - qPage.HeaderH - qPage.FooterH


End Sub

Public Property Let PageSize(ByVal vData As qePrinterPaperSize)
If mvarPageSize <> vData Then
  bPageChange = True
  mvarPageSize = vData
End If
End Property


Public Property Get PageSize() As qePrinterPaperSize
PageSize = mvarPageSize
End Property

Public Function AddText(ByVal Text As String, _
                        ByVal FontName As String, _
                        ByVal FontSize As Single, _
                        Optional FontBold As Boolean = False, _
                        Optional FontItalic As Boolean = False, _
                        Optional FontUnderline As Boolean = False, _
                        Optional FontColor As Long = &H0, _
                        Optional Align As qePrinterAlign = eLeft, _
                        Optional IndentLeft As Single = 0, _
                        Optional IndentRight As Single = 0, _
                        Optional Key As String) As Boolean


Dim qText As qcTextItem

Set qText = New qcTextItem
With qText
  .Text = Text
  .FontName = FontName
  .FontSize = FontSize
  .FontBold = FontBold
  .FontItalic = FontItalic
  .FontUnderline = FontUnderline
  .FontColor = FontColor
  .Alignment = Align
  .ScaleMode = mvarScaleMode
  .IndentLeft = IndentLeft
  .IndentRight = IndentRight
End With

qcolText.Add qText, Key
qText.GetSize
Set qText = Nothing
bPageChange = True

End Function

Private Sub Class_Initialize()

nPreviewState = -1
mvarScaleMode = eTwip
mvarPageSize = Printer.PaperSize
mvarScaleMode = eTwip
mvarMarginLeft = 1440
mvarMarginRight = 1440
mvarMarginTop = 1440
mvarMarginBottom = 1440
mvarPages = 1
mvarPrintOption = ShowPrinter_po + ShowCopies_po
mvarOrientation = Printer.Orientation
mvarOrientAvail = Orientation_Check

bPageChange = False
GetPageSize
HeadFootSetup

End Sub

Private Sub Class_Terminate()
Set qcolText = Nothing
Set qcolHF = Nothing

End Sub



Public Sub ResetItems()

Set qcolText = Nothing
Set qcolText = New qcTextItemCol
Set qcolHF = Nothing
Set qcolHF = New qcTextItemCol
HeadFootSetup
bPageChange = True
End Sub

Public Function RemoveItem(IndexKey As Variant) As Boolean

On Error GoTo RemoveItemError
' Remove Item from TextItem collection.  Returns True if
' successful, otherwise returns false

qcolText.Remove IndexKey
bPageChange = True
RemoveItem = True
Exit Function

RemoveItemError:
RemoveItem = False

End Function


Private Sub HeadFootSetup()

Dim qText As qcTextItem
Dim iCount As Integer

For iCount = 0 To 4
Set qText = New qcTextItem
qText.FontName = "Arial"
qText.FontSize = "10"
qText.HeadFoot = True
qText.ScaleMode = mvarScaleMode

qcolHF.Add qText, "Foot" & iCount
Set qText = Nothing
Next
For iCount = 0 To 4
Set qText = New qcTextItem
qText.FontName = "Arial"
qText.FontSize = "10"
qText.HeadFoot = True
qText.ScaleMode = mvarScaleMode
qcolHF.Add qText, "Head" & iCount
Set qText = Nothing
Next



End Sub

Public Property Get Header(ByVal PageType As qePrinterHeadFootPage) As qcTextItem
Dim sIndex As String

sIndex = "Head" & PageType

Set Header = qcolHF.Item(sIndex)

End Property

Public Property Let Header(ByVal PageType As qePrinterHeadFootPage, ByVal vData As qcTextItem)
Dim sIndex As String

sIndex = "Head" & PageType
bPageChange = True
Set qcolHF.Item(sIndex) = vData

End Property


Public Sub HeaderCopy(ByVal FromItem As qePrinterHeadFootPage, ByVal ToItem As qePrinterHeadFootPage)

HeadFootCopy "Head" & FromItem, "Head" & ToItem

End Sub

Private Sub HeadFootCopy(ByVal sFrom As String, ByVal sTo As String)
Dim qFrom As qcTextItem
Dim qTo As qcTextItem
Dim iCount As Integer

Set qFrom = qcolHF(sFrom)
Set qTo = qcolHF(sTo)
    qTo.Text = qFrom.Text
    qTo.FontUnderline = qFrom.FontUnderline
    qTo.FontItalic = qFrom.FontItalic
    qTo.FontBold = qFrom.FontBold
    qTo.FontColor = qFrom.FontColor
    qTo.FontSize = qFrom.FontSize
    qTo.FontName = qFrom.FontName
    qTo.ScaleMode = qFrom.ScaleMode
    qTo.IndentRight = qFrom.IndentRight
    qTo.IndentLeft = qFrom.IndentLeft
    qTo.Alignment = qFrom.Alignment
    qTo.BorderColor = qFrom.BorderColor
    qTo.BorderShading = qFrom.BorderShading
    qTo.BorderLine = qFrom.BorderLine
    qTo.ShowBorder = qFrom.ShowBorder

Set qFrom = Nothing
Set qTo = Nothing
bPageChange = True

End Sub

Public Property Get Footer(ByVal PageType As qePrinterHeadFootPage) As qcTextItem
Dim sIndex As String

sIndex = "Foot" & PageType

Set Footer = qcolHF.Item(sIndex)

End Property

Public Property Let Footer(ByVal PageType As qePrinterHeadFootPage, ByVal vData As qcTextItem)
Dim sIndex As String

sIndex = "Foot" & PageType
bPageChange = True
Set qcolHF.Item(sIndex) = vData

End Property
Public Sub FooterCopy(ByVal FromItem As qePrinterHeadFootPage, ByVal ToItem As qePrinterHeadFootPage)

HeadFootCopy "Foot" & FromItem, "Foot" & ToItem

End Sub

Public Property Get SetFooter(ByVal PageType As qePrinterHeadFootPage) As Boolean

SetFooter = mvarFoot(PageType)

End Property

Public Property Let SetFooter(ByVal PageType As qePrinterHeadFootPage, ByVal ShowFooter As Boolean)

mvarFoot(PageType) = ShowFooter
If ShowFooter = True Then
qcolHF.Item("Foot" & PageType).GetSize
End If

End Property

Public Property Get SetHeader(ByVal PageType As qePrinterHeadFootPage) As Boolean

SetHeader = mvarHead(PageType)

End Property

Public Property Let SetHeader(ByVal PageType As qePrinterHeadFootPage, ByVal ShowHeader As Boolean)

mvarHead(PageType) = ShowHeader
If ShowHeader = True Then
qcolHF.Item("Head" & PageType).GetSize
End If

End Property


Public Property Get AppName() As String

AppName = mvarAppName
End Property

Public Property Let AppName(ByVal sName As String)

mvarAppName = sName

End Property

Public Property Get Orientation() As qePrintOrientation
Orientation = mvarOrientation

End Property

Public Property Let Orientation(ByVal eOrientation As qePrintOrientation)

If mvarOrientation <> eOrientation And mvarOrientAvail Then
mvarOrientation = eOrientation
Printer.Orientation = mvarOrientation
bPageChange = True
End If

End Property

Public Property Get OrientOkay() As Boolean

OrientOkay = mvarOrientation

End Property

Private Function Orientation_Check() As Boolean

Dim eOrientOld As qePrintOrientation, eOrientNew As qePrintOrientation
Dim sWOld As Single, sWNew As Single
Dim sHOld As Single, sHNew As Single

' This is an exploratory procedure.  I haven't got a printer that does
' not respond to changing the orientation.  The printer orientation is
' affected by the Print Driver so some printers don't respond ... I'm
' hoping this function will determine whether orientation can be changed.

eOrientOld = Printer.Orientation
sWOld = Printer.Width
sHOld = Printer.Height
eOrientNew = eOrientOld Mod 2 + 1
Printer.Orientation = eOrientNew
sWNew = Printer.Width
sHNew = Printer.Height

If sWNew = sHOld And sHNew = sWOld Then
Orientation_Check = True
Else
Orientation_Check = False
If mvarOrientation <> eOrientOld Then
mvarOrientation = eOrientOld
bPageChange = True
End If
End If

Printer.Orientation = eOrientOld
' I'm not sure if these checks are required but knowing vb they
' probably are :-)
If Printer.Width <> sWOld Then
Printer.Width = sWOld
End If
If Printer.Height <> sHOld Then
Printer.Height = sHOld
End If

End Function

Public Sub FormPrint(ByRef frmPrint As Object, _
                     Optional ParentContain As Object, _
                     Optional FormAlign As qePrinterAlign, _
                     Optional InputBorder As Boolean = False, _
                     Optional TopOffset As Single, _
                     Optional AutoHeight As Boolean = True, _
                     Optional ExcludeList As String)

Dim sWidth As Single, sContainer As String
Dim bContainer As Boolean
Dim ctrTemp As Control
Dim sText As String, sKey As String
Dim bPrint As Boolean, bText As Boolean, bAlign As Boolean, bOption As Boolean
Dim eAlign As qePrinterAlign
Dim sRatio As Single, sLeft As Single
Dim nListCount As Integer
Dim oParent As Object
Dim sPosX As Single, sPosY As Single
Dim sParentKey As String, sParentList As String
Dim sOptLeft As Single, sOptRight As Single, sOptTop As Single
Dim nOptionTotal As Integer, nOption As Integer

On Error Resume Next
' Check whether Printing Parent Container and get dimensions
If ParentContain Is Nothing Then
sWidth = frmPrint.ScaleWidth
sContainer = frmPrint.Name
Else
bContainer = True
sWidth = ParentContain.ScaleWidth
sContainer = ParentContain.Name
sContainer = ParentContain.Name & "/" & ParentContain.Index
End If

'Initialise printing the form contents

ResetItems
TopOffset = ConvertToTwip(mvarScaleMode, TopOffset)
If bPageChange Then
GetPageSize
End If
' Check form width is smaller than page width
If qPage.AvailWidth < sWidth Then
FormAlign = eJustify
Else
sRatio = 1
End If
  
Select Case FormAlign
Case qePrinterAlign.eLeft
Case qePrinterAlign.eRight
sLeft = qPage.AvailWidth - sWidth
'mvarMarginLeft = sLeft + 567
Case qePrinterAlign.eCentre
sLeft = (qPage.AvailWidth - sWidth) / 2
'mvarMarginLeft = sLeft + 567
Case qePrinterAlign.eJustify
sRatio = qPage.AvailWidth / sWidth
End Select


For Each ctrTemp In frmPrint
'Check for control type
bText = True
bPrint = True
bAlign = True
bOption = False
If TypeOf ctrTemp Is TextBox Then
sText = ctrTemp.Text
ElseIf TypeOf ctrTemp Is Label Then
sText = ctrTemp.Caption
bText = False
ElseIf TypeOf ctrTemp Is ComboBox Then
bAlign = False
eAlign = eLeft
ElseIf TypeOf ctrTemp Is ListBox Then
nListCount = 0
sText = ""
Do While nListCount < ctrTemp.ListCount
sText = sText & ctrTemp.List(nListCount) & vbCrLf
nListCount = nListCount + 1
Loop
bAlign = False
ElseIf TypeOf ctrTemp Is CheckBox Then
If ctrTemp.Value = vbChecked Then
sText = Chr$(254)
Else
sText = Chr$(111)
End If
' Add Checkbox with tick or empty checkbox
sText = "<FONT=wingdings>" & sText & "</FONT>  " & ctrTemp.Caption
bText = False
bAlign = False
ElseIf TypeOf ctrTemp Is OptionButton Then
bOption = True
Else
bPrint = False
End If

If bPrint Then
' Create Key name for control
sKey = ctrTemp.Name
' If the Control (ctrTemp) is not part of an array an error will occur
' and the statement will not be processed
sKey = ctrTemp.Name & "/" & ctrTemp.Index
' Check if in excluded list
If InStr(1, ExcludeList, "*" & sKey) Then
bPrint = False
Else
' Check to see whether control is in a container(s) and check if
' parent is in exclude list, and calculate additional Left and
' Top positioning.
Set oParent = ctrTemp
sPosY = 0: sPosX = 0
If bContainer Then
bPrint = False
End If
Do
Set oParent = oParent.Container
sParentKey = oParent.Name
sParentKey = oParent.Name & "/" & oParent.Index
If InStr(1, ExcludeList, "*" & sParentKey) > 0 Then
bPrint = False
bOption = False
ElseIf oParent.Name <> sContainer Then
sPosY = sPosY + oParent.Top
sPosX = sPosX + oParent.Left
Else
bPrint = True
End If
Loop While oParent.Name <> frmPrint.Name And sParentKey <> sContainer And bPrint
Set oParent = Nothing
If Not bPrint Then
ExcludeList = ExcludeList & "*" & sKey
End If
End If
End If

If bOption And bPrint Then
sParentKey = ctrTemp.Container.Name
sParentKey = ctrTemp.Container.Name & "/" & ctrTemp.Container.Index
sOptLeft = ctrTemp.Left * sRatio + sLeft + sPosX
sOptRight = qPage.AvailWidth - sLeft - ((sPosX + ctrTemp.Left + ctrTemp.Width) * sRatio)
sOptTop = ctrTemp.Top + TopOffset + sPosY
If InStr(1, sParentList, "*" & sParentKey) = 0 Then
sParentList = sParentList & "*" & sParentKey
AddText "", ctrTemp.FontName, ctrTemp.FontSize, ctrTemp.FontBold, ctrTemp.FontItalic, ctrTemp.FontUnderline, ctrTemp.ForeColor, , , , "Opt:" & sParentKey
With qcolText("Opt:" & sParentKey)
'Set TextItem to Absolute and calculate position on form and page
.Absolute = True
.AbsPage = 1
.IndentLeft = sOptLeft '= ctrTemp.Left * sRatio + sLeft + sPosX
.IndentRight = sOptRight 'qPage.AvailWidth - sLeft - ((sPosX + ctrTemp.Left + ctrTemp.Width) * sRatio)
.Top = sOptTop 'ctrTemp.Top + TopOffset + sPosY
.GetSize
If ctrTemp.Value Then
.Text = ctrTemp.Caption
End If
End With
Else
With qcolText("Opt:" & sParentKey)
If .Top > sOptTop Then
.Top = sOptTop
.IndentLeft = sOptLeft
ElseIf .IndentLeft > sOptLeft Then
.IndentLeft = sOptLeft
.Top = sOptTop
End If
If .IndentLeft + .IndentRight > sOptLeft + sOptRight Then
.IndentRight = qPage.AvailWidth - sLeft - .IndentLeft - (ctrTemp.Width * sRatio)
End If
If ctrTemp.Value = True Then
.Text = ctrTemp.Caption
End If
End With
End If
bPrint = False
End If

If bPrint Then
'Adjust alignment
If bAlign Then
If ctrTemp.Alignment = vbRightJustify Then
eAlign = eRight
ElseIf ctrTemp.Alignment = vbCenter Then
eAlign = eCentre
Else
eAlign = eLeft
End If
End If
AddText sText, ctrTemp.FontName, ctrTemp.FontSize, ctrTemp.FontBold, ctrTemp.FontItalic, ctrTemp.FontUnderline, ctrTemp.ForeColor, eAlign, , , sKey
With qcolText(sKey)
'Set TextItem to Absolute and calculate position on form and page
.Absolute = True
.AbsPage = 1
.IndentLeft = ctrTemp.Left * sRatio + sLeft + sPosX
.IndentRight = qPage.AvailWidth - sLeft - ((sPosX + ctrTemp.Left + ctrTemp.Width) * sRatio)
If bText And InputBorder Then
.BorderLine = 1
.ShowBorder = True
End If
.Top = ctrTemp.Top + TopOffset + sPosY
.MaxHeight = ctrTemp.Height
.OriginalHeight = ctrTemp.Height
.GetSize
End With
End If
Next

sExcludeList = ExcludeList
sFormName = frmPrint.Name
Set frmPrint = Nothing

bAutoHeight = AutoHeight

If AutoHeight Then
Form_CheckHeight
End If

End Sub

Private Sub Form_CheckHeight()
' Check items that require more space and move items below them
' further down to allow expansion

Dim nTallItem As Integer
Dim nMoveItem As Integer
Dim sTopMin As Single
Dim sTopAdd As Single
nTallItem = 1
Do While nTallItem <= qcolText.Count
qcolText(nTallItem).AdditionalTop = 0
qcolText(nTallItem).GetSize
nTallItem = nTallItem + 1
Loop
nTallItem = 1
Do While nTallItem <= qcolText.Count
If qcolText(nTallItem).HeightTwip > qcolText(nTallItem).OriginalHeight And qcolText(nTallItem).OriginalHeight <> 0 Then
sTopMin = qcolText(nTallItem).TopTwip + qcolText(nTallItem).OriginalHeight
sTopAdd = qcolText(nTallItem).HeightTwip - qcolText(nTallItem).OriginalHeight
nMoveItem = 1
Do While nMoveItem <= qcolText.Count
If qcolText(nMoveItem).TopTwip >= sTopMin Then
qcolText(nMoveItem).AdditionalTop = qcolText(nMoveItem).AdditionalTop + sTopAdd
End If
nMoveItem = nMoveItem + 1
Loop

qcolText(nTallItem).MaxHeight = 0
End If
nTallItem = nTallItem + 1
Loop


End Sub
Public Sub FormPrint_Update(ByRef frmPrint As Object)

Dim ctrTemp As Control
Dim sKey As String
Dim sText As String
Dim nListCount As Integer
Dim bInclude As Boolean
If sFormName = "" Then
Err.Raise vbObjectError + 29998, "qPrinter", "FormPrint_Update cannot be called.  The document has not be initialised with FormPrint."
ElseIf sFormName <> frmPrint.Name Then
Err.Raise vbObjectError + 29999, "qPrinter", "FormPrint_Update cannot be called.  The form is different to the document form."
End If

'Update each control with new Text/Caption values
For Each ctrTemp In frmPrint
On Error Resume Next
sKey = ctrTemp.Name
sKey = ctrTemp.Name & "/" & ctrTemp.Index
If InStr(1, sExcludeList, "*" & sKey) = 0 Then
If TypeOf ctrTemp Is TextBox Then
qcolText(sKey).Text = ctrTemp.Text
ElseIf TypeOf ctrTemp Is Label Then
qcolText(sKey).Text = ctrTemp.Caption
ElseIf TypeOf ctrTemp Is ComboBox Then
qcolText(sKey).Text = ctrTemp.Text
ElseIf TypeOf ctrTemp Is ListBox Then
nListCount = 0
sText = ""
Do While nListCount < ctrTemp.ListCount
sText = sText & ctrTemp.List(nListCount) & vbCrLf
nListCount = nListCount + 1
Loop
qcolText(sKey).Text = sText
ElseIf TypeOf ctrTemp Is CheckBox Then
If ctrTemp.Value = vbChecked Then
sText = Chr$(254)
Else
sText = Chr$(111)
End If
qcolText(sKey).Text = "<FONT=wingdings>" & sText & "</FONT>  " & ctrTemp.Caption
ElseIf TypeOf ctrTemp Is OptionButton Then
If ctrTemp.Value Then
sKey = ctrTemp.Container.Name
sKey = ctrTemp.Container.Name & "/" & ctrTemp.Container.Index
qcolText("Opt:" & sKey).Text = ctrTemp.Caption
End If
End If
End If
Next
Set frmPrint = Nothing
If bAutoHeight Then
Form_CheckHeight
End If

End Sub


