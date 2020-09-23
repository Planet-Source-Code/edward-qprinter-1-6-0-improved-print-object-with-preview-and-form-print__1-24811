VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "qcTextItem"
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
' Module:       qcTextItem
' Purpose:      Holds details of each Text Item in Document
' ==============================================================

Option Explicit

Private mvarText As String
Private mvarLineHeight As Single
Private mvarLines As Integer
Private mvarFontName As String
Private mvarFontSize As Single
Private mvarFontBold As Boolean
Private mvarFontItalic As Boolean
Private mvarFontUnderline As Boolean
Private mvarFontColor As Long
Private mvarIndentRight As Single
Private mvarIndentLeft As Single
Private mvarAlignment As qePrinterAlign
Private mvarAbsolute As Boolean
Private mvarTop As Single
Private mvarLeft As Single
Private mvarWidth As Single
Private mvarHeight As Single
Private mvarLineH() As Single
Private mvarShowBorder As Boolean
Private mvarBorderLine As Integer
Private mvarBorderColor As Long
Private mvarBorderShading As Long
Private mvarAbsPage As Integer
Private mbChanged As Boolean
Private mvarNewPage As qeTextNewPageEnum
Private mvarScaleMode As qePrinterScale
Private mvarStartPage As Integer
Private mvarEndPage As Integer
Private mvarHeadFoot As Boolean
Private mvarHFText As String

Private mvarPositionStart As Single
Private mvarPositionEnd As Single
Private mvarTempAlign As qePrinterAlign
Private mvarUnderline As Integer
Private mvarUnderlineAll As Integer

Private mvarTempLIndent As Single
Private mvarTempRIndent As Single
Private mvarCurrentLIndent As Single
Private mvarCurrentRIndent As Single

' *** Added 1.6.0
Private mvarLineSpacing As Single
Private mvarMaxHeight As Single
Private mvarOriginalHeight As Single
Private mvarAdditionalTop As Single
Friend Property Get AdditionalTop() As Single
AdditionalTop = mvarAdditionalTop
End Property
Friend Property Let AdditionalTop(ByVal sValue As Single)
mvarAdditionalTop = sValue
End Property

Friend Property Get OriginalHeight() As Single
OriginalHeight = mvarOriginalHeight
End Property
Friend Property Let OriginalHeight(ByVal sValue As Single)
mvarOriginalHeight = sValue
End Property

Public Property Get MaxHeight() As Single
MaxHeight = ConvertFromTwip(mvarScaleMode, mvarMaxHeight, False)
End Property
Friend Property Get MaxHeightTwip() As Single
MaxHeightTwip = mvarMaxHeight
End Property
Public Property Let MaxHeight(ByVal sValue As Single)
mvarMaxHeight = ConvertToTwip(mvarScaleMode, sValue, False)
bPageChange = True
End Property
Public Property Get LineSpacing() As Single
LineSpacing = ConvertFromTwip(mvarScaleMode, mvarLineSpacing, False)
End Property
Friend Property Get LineSpacingTwip() As Single
LineSpacingTwip = mvarLineSpacing
End Property
Public Property Let LineSpacing(ByVal sValue As Single)
mvarLineSpacing = ConvertToTwip(mvarScaleMode, sValue, False)
bPageChange = True
End Property
' *** End Added 1.6.0

Friend Property Get CurrentLIndent() As Single
CurrentLIndent = mvarCurrentLIndent
End Property
Friend Property Let CurrentLIndent(ByVal sValue As Single)
mvarCurrentLIndent = sValue
End Property
Friend Property Get CurrentRIndent() As Single
CurrentRIndent = mvarCurrentRIndent
End Property
Friend Property Let CurrentRIndent(ByVal sValue As Single)
mvarCurrentRIndent = sValue
End Property
Friend Property Get TempLIndent() As Single
TempLIndent = mvarTempLIndent
End Property
Friend Property Let TempLIndent(ByVal sValue As Single)
mvarTempLIndent = sValue
End Property
Friend Property Get TempRIndent() As Single
TempRIndent = mvarTempRIndent
End Property
Friend Property Let TempRIndent(ByVal sValue As Single)
mvarTempRIndent = sValue
End Property

Public Property Get Underline() As Integer
Underline = mvarUnderline
End Property

Public Property Let Underline(ByVal iLineSize As Integer)
mvarUnderline = iLineSize
End Property

Public Property Get UnderlineAll() As Integer
UnderlineAll = mvarUnderlineAll
End Property

Public Property Let UnderlineAll(ByVal iLineSize As Integer)
mvarUnderlineAll = iLineSize
End Property

Public Property Get PositionStart() As Single
PositionStart = ConvertFromTwip(mvarScaleMode, mvarPositionStart, False)
End Property
Friend Property Get PositionStartTwip() As Single
PositionStartTwip = mvarPositionStart
End Property
Friend Property Let PositionStartTwip(ByVal sValue As Single)
mvarPositionStart = sValue
End Property

Public Property Get PositionEnd() As Single
PositionEnd = ConvertFromTwip(mvarScaleMode, mvarPositionEnd, False)
End Property
Friend Property Get PositionEndTwip() As Single
PositionEndTwip = mvarPositionEnd
End Property
Friend Property Let PositionEndTwip(ByVal sValue As Single)
mvarPositionEnd = sValue
End Property

Friend Property Get TempAlign() As qePrinterAlign
TempAlign = mvarTempAlign
End Property
Friend Property Let TempAlign(ByVal eValue As qePrinterAlign)
mvarTempAlign = eValue
End Property




Friend Property Get HeadFoot() As Boolean
HeadFoot = mvarHeadFoot
End Property
Friend Property Let HeadFoot(ByVal bValue As Boolean)

mvarHeadFoot = bValue

End Property
Friend Sub SetHFText(ByVal PageNum As Integer, ByVal PageTotal As Integer, ByVal AppName As String)
FlyText PageNum, PageTotal, AppName
End Sub
Friend Property Get HFText() As String

HFText = mvarHFText

End Property
Private Sub FlyText(ByVal PageNum As Integer, ByVal PageTotal As Integer, ByVal AppName As String)

mvarHFText = mvarText
mvarHFText = Replace(mvarHFText, "#pagenumber#", PageNum)
mvarHFText = Replace(mvarHFText, "#pagetotal#", PageTotal)
mvarHFText = Replace(mvarHFText, "#longdate#", Format(Date, "Long Date"))
mvarHFText = Replace(mvarHFText, "#shortdate#", Format(Date, "Short Date"))
mvarHFText = Replace(mvarHFText, "#longtime#", Format(Time, "Long Time"))
mvarHFText = Replace(mvarHFText, "#shorttime#", Format(Time, "Short Time"))
mvarHFText = Replace(mvarHFText, "#appname#", AppName)

End Sub

Public Property Get StartPage() As Integer
StartPage = mvarStartPage
End Property

Friend Property Let StartPage(ByVal iValue As Integer)

mvarStartPage = iValue
'bPageChange = True
End Property
Public Property Get EndPage() As Integer
EndPage = mvarEndPage
End Property

Friend Property Let EndPage(ByVal iValue As Integer)

mvarEndPage = iValue

End Property



Public Property Get NewPage() As qeTextNewPageEnum
NewPage = mvarNewPage
End Property

Public Property Let NewPage(ByVal bValue As qeTextNewPageEnum)

mvarNewPage = bValue
bPageChange = True
End Property
Public Property Get ShowBorder() As Boolean
ShowBorder = mvarShowBorder

End Property

Public Property Let ShowBorder(ByVal bValue As Boolean)

mvarShowBorder = bValue

End Property


Public Property Get BorderLine() As Integer
BorderLine = mvarBorderLine
End Property
Public Property Let BorderLine(ByVal lValue As Integer)

mvarBorderLine = lValue

End Property
Public Property Get BorderShading() As Long
BorderShading = mvarBorderShading
End Property
Public Property Let BorderShading(ByVal lValue As Long)

mvarBorderShading = lValue

End Property
Public Property Get BorderColor() As Long
BorderColor = mvarBorderColor
End Property
Public Property Let BorderColor(ByVal lValue As Long)

mvarBorderColor = lValue

End Property


Public Property Get AbsPage() As Integer
AbsPage = mvarAbsPage
End Property
Public Property Let AbsPage(ByVal lValue As Integer)

mvarAbsPage = lValue
End Property
Public Property Get Absolute() As Boolean
Absolute = mvarAbsolute
End Property

Public Property Let Absolute(ByVal bValue As Boolean)

mvarAbsolute = bValue
bPageChange = True
End Property

Public Property Get Left() As Single

Left = ConvertFromTwip(mvarScaleMode, mvarLeft)
End Property

Public Property Let Left(ByVal sValue As Single)
mvarLeft = ConvertToTwip(mvarScaleMode, sValue)
End Property
Public Property Get Top() As Single
Top = ConvertFromTwip(mvarScaleMode, mvarTop, False)
End Property
Friend Property Get TopTwip() As Single
TopTwip = mvarTop
End Property
Public Property Let Top(ByVal sValue As Single)
mvarTop = ConvertToTwip(mvarScaleMode, sValue, False)
bPageChange = True
End Property

Public Property Get Height() As Single

Height = ConvertFromTwip(mvarScaleMode, mvarHeight, False)
End Property
Friend Property Get HeightTwip() As Single

HeightTwip = mvarHeight

End Property

Friend Property Get WidthTwip() As Single
WidthTwip = mvarWidth
End Property
Public Property Get Width() As Single
Width = ConvertFromTwip(mvarScaleMode, mvarWidth)
End Property

Public Property Let ScaleMode(ByVal vData As qePrinterScale)
mvarScaleMode = vData
End Property


Public Property Get ScaleMode() As qePrinterScale
ScaleMode = mvarScaleMode
End Property
'Friend Sub GetSize(MarginLeft As Single, _
'                   MarginRight As Single, _
'                   Optional ScaleMode As qePrinterScale)
Public Sub GetSize(Optional ScaleMode As qePrinterScale)

Dim sHeight As Single
Dim sWidth As Single
Dim sSizeX As Single
Dim sLine As String, sChar As String, sWord As String
Dim sLineW As Single, sLineH As Single, sWordW As Single
Dim sLineHeight As Single
Dim iLines As Integer, lCount As Long
Dim eCharType As qePrinterChar, eEnd As qePrinterChar
Dim bNewLine As Boolean
Dim sfFont As New StdFont
Dim sFormat As String, lFPos As Long
Dim bCheck As Boolean
Dim sText As String
Dim bForceSameLine As Boolean
Dim sTempLIndent As Single
Dim sTempRIndent As Single
Dim sDefaultWidth As Single

With Printer.Font
  .Name = mvarFontName
  .Size = mvarFontSize
  .Bold = mvarFontBold
  .Italic = mvarFontItalic
  .Underline = mvarFontUnderline
End With
' *** Changed 1.6.0
mvarLineHeight = Printer.TextHeight("H") + mvarLineSpacing
' *** End Change 1.6.0

sWidth = Printer.ScaleWidth - qPage.LeftM - qPage.RightM - mvarIndentLeft - mvarIndentRight
sDefaultWidth = sWidth

If mvarHeadFoot Then
sText = mvarHFText
Else
sText = mvarText
End If


ReDim mvarLineH(0)
ReDim mvarPosition(0)
If sWidth <= 0 Then
  Exit Sub
End If

With Printer
  sLine = ""
  sWord = ""
  sLineH = 0: sWordW = 0: sLineW = 0
  eEnd = eNone
  lCount = 1
  If Not mvarAbsolute Then
  sHeight = mvarTop
  End If
  Do
    bNewLine = False
    
    Do
      Do
        eCharType = eNone
        bCheck = True
        sChar = Mid$(sText, lCount, 1)
        ' Check for formatting codes
        If sChar = "<" Then
        lFPos = InStr(lCount, sText, ">")
        If lFPos > 0 Then
        sFormat = Mid$(sText, lCount + 1, lFPos - lCount - 1)
sFormat = UCase(sFormat)
bCheck = False
If Len(sFormat) < 3 Or Left(sFormat, 1) = "/" Then
Select Case sFormat
Case "B"
Printer.Font.Bold = True
Case "U"
Printer.Font.Underline = True
Case "I"
Printer.Font.Italic = True
Case "/B"
Printer.Font.Bold = False
Case "/I"
Printer.Font.Italic = False
Case "/U"
Printer.Font.Underline = False
Case "/FONT"
Printer.Font.Name = mvarFontName
Case "/COLOR"
Printer.ForeColor = mvarFontColor
Case "/SIZE"
Printer.Font.Size = mvarFontSize
Case "/ALIGN"
' Do nothing
Case "/LINDENT"
sWidth = sDefaultWidth - sTempRIndent
sTempLIndent = 0
Case "/RINDENT"
sWidth = sDefaultWidth - sTempLIndent
sTempRIndent = 0

Case Else
lFPos = lCount
bCheck = True
End Select
ElseIf Left(sFormat, 5) = "FONT=" Then
Printer.Font.Name = Property_Extract(sFormat)
ElseIf Left(sFormat, 6) = "COLOR=" Then
Printer.ForeColor = CLng(Property_Extract(sFormat))
ElseIf Left(sFormat, 5) = "SIZE=" Then
Printer.Font.Size = Val(Property_Extract(sFormat))
ElseIf Left(sFormat, 6) = "ALIGN=" Then
' do nothing

ElseIf Left$(sFormat, 8) = "LINDENT=" Then
sTempLIndent = ConvertToTwip(mvarScaleMode, Val(Property_Extract(sFormat)))
sWidth = sDefaultWidth - sTempLIndent - sTempRIndent
bForceSameLine = True
ElseIf Left$(sFormat, 8) = "RINDENT=" Then
sTempRIndent = ConvertToTwip(mvarScaleMode, Val(Property_Extract(sFormat)))
sWidth = sDefaultWidth - sTempLIndent - sTempRIndent
bForceSameLine = True
ElseIf Left(sFormat, 5) = "FORCE" Then
bForceSameLine = True

Else
bCheck = False
End If
If Not bCheck Then
lCount = lFPos
End If

        End If
        End If
' CHARACTER CHECK: Look for potential line breaks or where text
'                  width is greater than boundary
        If bCheck Then
        Select Case sChar
          Case " "
            eCharType = eSpace
          Case "-"
            sSizeX = sLineW + sWordW + .TextWidth(sChar)
            If .TextHeight(sChar) > sLineH Then
            sLineH = .TextHeight(sChar)
            End If
            If sSizeX > sWidth Then
              eCharType = eOops
            Else
              eCharType = eDash
            End If
          Case vbLf
            sChar = ""
            eCharType = eLine
          Case vbCr
            If lCount < Len(sText) Then
              If Mid$(sText, lCount + 1, 1) = vbLf Then
                lCount = lCount + 1
              End If
            End If
            sChar = ""
            eCharType = eLine
          Case Else
' CHARACTER CHECK: See if addition of character makes line too long
            sSizeX = sLineW + sWordW + .TextWidth(sChar)
            If .TextHeight(sChar) > sLineH Then
            sLineH = .TextHeight(sChar)
            End If
            If sSizeX > sWidth Then
              eCharType = eOops
            Else
              sWord = sWord & sChar
              sWordW = sWordW + .TextWidth(sChar)
            End If
        End Select
        End If
        lCount = lCount + 1
      Loop While lCount <= Len(sText) And eCharType = eNone And Not bForceSameLine
      If bForceSameLine Then
      eCharType = eLine
      End If
      
' LINE SPLIT: Examine potential line break
      If lCount > Len(sText) Then
        eCharType = eLine
      End If
      Select Case eCharType
        Case qePrinterChar.eNone
          sLine = sLine & sWord
          sLineW = sLineW + sWordW
          eEnd = eLine
        Case qePrinterChar.eOops
          If eEnd = eNone Then
            sLine = sWord
            sLineW = sLineW + sWordW
            sWord = sChar
            sWordW = .TextWidth(sChar)
          Else
            sLine = Trim$(sLine)
            sWord = sWord & sChar
            sWordW = sWordW + .TextWidth(sChar)
          End If
          bNewLine = True
        Case qePrinterChar.eDash, qePrinterChar.eSpace
          eEnd = eCharType
          sLine = sLine & sWord & sChar
          sLineW = sLineW + sWordW + .TextWidth(sChar)
          sWord = ""
          sWordW = 0
        Case qePrinterChar.eLine
          If sLineH = 0 Then
          sLineH = Printer.TextHeight("H")
          End If
          sLine = sLine & sWord
          sLineW = sLineW + sWordW
          eEnd = eLine
          sWord = ""
          sWordW = 0
          bNewLine = True
      End Select
' LINE SPLIT: Add new line if required
      If bNewLine Then
        If Not bForceSameLine Then
        
        iLines = iLines + 1
        ReDim Preserve mvarLineH(iLines)
        If sLineW > mvarWidth Then
        mvarWidth = sLineW
        End If
        ' *** Changed 1.6.0
        mvarLineH(iLines) = sLineH + mvarLineSpacing
        sHeight = sHeight + sLineH + mvarLineSpacing
        ' *** End Change 1.6.0
        sLineH = 0
        End If
        sLine = ""
        sLineW = 0
        eEnd = eNone
        bForceSameLine = False
      End If
    Loop While Not bNewLine
  Loop While lCount <= Len(sText)
End With

If iLines > 1 Or mvarWidth > sWidth Then
mvarWidth = sWidth
End If
If mvarShowBorder And mvarBorderLine > 0 Then
mvarLineH(iLines) = mvarLineH(iLines) + mvarBorderLine * Printer.TwipsPerPixelY
End If
mvarHeight = sHeight

mvarLines = iLines

End Sub
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
Public Property Let Alignment(ByVal vData As qePrinterAlign)
mvarAlignment = vData
End Property

Public Property Get Alignment() As qePrinterAlign
Alignment = mvarAlignment
End Property

Public Property Let IndentLeft(ByVal vData As Single)
    
mvarIndentLeft = ConvertToTwip(mvarScaleMode, vData)
bPageChange = True
End Property

Public Property Get IndentLeft() As Single
IndentLeft = ConvertFromTwip(mvarScaleMode, mvarIndentLeft)
End Property
Friend Property Get IndentLeftTwip() As Single
IndentLeftTwip = mvarIndentLeft
End Property

Public Property Let IndentRight(ByVal vData As Single)
mvarIndentRight = ConvertToTwip(mvarScaleMode, vData)
bPageChange = True
End Property

Public Property Get IndentRight() As Single
IndentRight = ConvertFromTwip(mvarScaleMode, mvarIndentRight)
End Property
Friend Property Get IndentRightTwip() As Single
IndentRightTwip = mvarIndentRight
End Property

Public Property Let FontName(ByVal vData As String)
mvarFontName = vData
bPageChange = True
End Property

Public Property Get FontName() As String
FontName = mvarFontName
End Property
Public Property Let FontSize(ByVal vData As Single)
mvarFontSize = vData
bPageChange = True
End Property

Public Property Get FontSize() As Single
FontSize = mvarFontSize
End Property

Public Property Let FontColor(ByVal vData As Long)
mvarFontColor = vData
End Property

Public Property Get FontColor() As Long
FontColor = mvarFontColor
End Property

Public Property Let FontBold(ByVal vData As Boolean)
mvarFontBold = vData
bPageChange = True
End Property

Public Property Get FontBold() As Boolean
FontBold = mvarFontBold
End Property

Public Property Let FontItalic(ByVal vData As Boolean)
mvarFontItalic = vData
bPageChange = True
End Property

Public Property Get FontItalic() As Boolean
FontItalic = mvarFontItalic
End Property
Public Property Let FontUnderline(ByVal vData As Boolean)
mvarFontUnderline = vData
bPageChange = True
End Property

Public Property Get FontUnderline() As Boolean
FontUnderline = mvarFontUnderline
End Property

Friend Property Get Lines() As Integer
Lines = mvarLines
End Property

Friend Property Get LineHeight(Index As Integer) As Single
LineHeight = mvarLineH(Index)

End Property


Public Property Let Text(ByVal vData As String)
mvarText = vData
bPageChange = True
End Property

Public Property Get Text() As String
Text = mvarText
End Property

Private Sub Class_Initialize()
mvarBorderShading = -1
End Sub


