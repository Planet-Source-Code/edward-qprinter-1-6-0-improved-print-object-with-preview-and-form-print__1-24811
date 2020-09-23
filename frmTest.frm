VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Enhanced Print with Preview Test"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   ForeColor       =   &H00C00000&
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraUser 
      Caption         =   "User Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      TabIndex        =   21
      Top             =   2520
      Width           =   1695
      Begin VB.CheckBox chkCopies 
         Appearance      =   0  'Flat
         Caption         =   "Choose Copies"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkPrinter 
         Appearance      =   0  'Flat
         Caption         =   "Choose Printer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Value           =   1  'Checked
         Width           =   1455
      End
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "Orientation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5040
      TabIndex        =   16
      Top             =   840
      Width           =   1695
      Begin VB.OptionButton optOrientation 
         Appearance      =   0  'Flat
         Caption         =   "Landscape"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optOrientation 
         Appearance      =   0  'Flat
         Caption         =   "Portrait"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblOStatus 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Available"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblOrientation 
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkLast 
      Appearance      =   0  'Flat
      Caption         =   "Print date and time information on last page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   5760
      Width           =   3735
   End
   Begin VB.CheckBox chkJustify 
      Appearance      =   0  'Flat
      Caption         =   "Justify footer left or right for odd/even pages"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   5520
      Width           =   3735
   End
   Begin VB.TextBox txtFooter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Text            =   "frmTest.frx":27A2
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CheckBox chkAfter 
      Appearance      =   0  'Flat
      Caption         =   "Add new page after last piece of text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   4080
      Width           =   3735
   End
   Begin VB.CheckBox chkBefore 
      Appearance      =   0  'Flat
      Caption         =   "Add new page before next piece of text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CheckBox chkWatermark 
      Appearance      =   0  'Flat
      Caption         =   "Show background watermark text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox txtHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "frmTest.frx":27C5
      Top             =   840
      Width           =   3735
   End
   Begin VB.CheckBox chkHeader 
      Appearance      =   0  'Flat
      Caption         =   "Print header on first page only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox txtTest 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   795
      Index           =   2
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3240
      Width           =   3735
   End
   Begin VB.TextBox txtTest 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   510
      Index           =   1
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmTest.frx":27DF
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox txtTest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   0
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      TabIndex        =   0
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox txtTest 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   510
      Index           =   3
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Label lblInformation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmTest.frx":2DA6
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label lblItem 
      Caption         =   "Footer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   5100
      Width           =   975
   End
   Begin VB.Label lblItem 
      Caption         =   "Header:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================
' Project:      Enhance Print object Test Project
' Author:       edward moth
' Copyright:    © 2000-2001 qbd software ltd
'
' ==============================================================
' Module:       frmTest
' Purpose:      Test qcPrinter Class
' ==============================================================

Option Explicit

Dim qPrint As New qcPrinter

Private Sub chkAfter_Click()
If chkAfter.Value = vbChecked Then
If CBool(qPrint.TextItem(3).NewPage And Before_np) Then
qPrint.TextItem(3).NewPage = Both_np
Else
qPrint.TextItem(3).NewPage = After_np
End If
Else
If qPrint.TextItem(3).NewPage = Both_np Then
qPrint.TextItem(3).NewPage = Before_np
Else
qPrint.TextItem(3).NewPage = None_np
End If
End If

End Sub

Private Sub chkBefore_Click()

If chkBefore.Value = vbChecked Then
If CBool(qPrint.TextItem(3).NewPage And After_np) Then
qPrint.TextItem(3).NewPage = Both_np
Else
qPrint.TextItem(3).NewPage = Before_np
End If
Else
If qPrint.TextItem(3).NewPage = Both_np Then
qPrint.TextItem(3).NewPage = After_np
Else
qPrint.TextItem(3).NewPage = None_np
End If
End If

End Sub

Private Sub chkCopies_Click()
If chkCopies.Value = vbChecked Then
If CBool(qPrint.PrintOptions And ShowPrinter_po) Then
qPrint.PrintOptions = ShowPrinter_po + ShowCopies_po
Else
qPrint.PrintOptions = ShowCopies_po
End If
Else
If CBool(qPrint.PrintOptions And ShowPrinter_po) Then
qPrint.PrintOptions = ShowPrinter_po
Else
qPrint.PrintOptions = 0
End If
End If

End Sub

Private Sub chkHeader_Click()

qPrint.SetHeader(FirstPage_hf) = CBool(chkHeader.Value = vbChecked)
qPrint.SetHeader(OddPage_hf) = CBool(chkHeader.Value = vbUnchecked)
qPrint.SetHeader(EvenPage_hf) = CBool(chkHeader.Value = vbUnchecked)

End Sub

Private Sub chkJustify_Click()

If chkJustify.Value = vbChecked Then
qPrint.Footer(OddPage_hf).Alignment = eRight
qPrint.Footer(EvenPage_hf).Alignment = eLeft
Else
qPrint.Footer(OddPage_hf).Alignment = eCentre
qPrint.Footer(EvenPage_hf).Alignment = eCentre
End If

End Sub

Private Sub chkLast_Click()

qPrint.SetFooter(LastPage_hf) = CBool(chkLast.Value = vbChecked)

End Sub

Private Sub chkPrinter_Click()
If chkPrinter.Value = vbChecked Then
If CBool(qPrint.PrintOptions And ShowCopies_po) Then
qPrint.PrintOptions = ShowPrinter_po + ShowCopies_po
Else
qPrint.PrintOptions = ShowPrinter_po
End If
Else
If CBool(qPrint.PrintOptions And ShowCopies_po) Then
qPrint.PrintOptions = ShowCopies_po
Else
qPrint.PrintOptions = 0
End If
End If


End Sub

Private Sub chkWatermark_Click()

If chkWatermark.Value = vbChecked Then
qPrint.AddText "qbd enhanced print preview", "Arial Black", 36, , , , RGB(240, 240, 240), eCentre, 40, 40, "Watermark"
With qPrint.TextItem("Watermark")
.Absolute = True
' Set page to -1 to print on all pages
.AbsPage = -1
.Top = 80
End With
Else
qPrint.RemoveItem ("Watermark")
End If

End Sub

Private Sub cmdPreview_Click()

qPrint.Preview

End Sub


Private Sub cmdClose_Click()
Set qPrint = Nothing
Unload Me

End Sub



Private Sub Form_Load()

Preview_SetUp
End Sub

Private Sub Preview_SetUp()

Dim iCount As Integer
txtTest(0).Text = "Enhanced Printer Control" & vbCrLf & "with <I>Preview</I>"
txtTest(2).Text = "Right justified text" & vbCrLf & "can be handy for" & vbCrLf & "addresses" & vbCrLf & vbCrLf
txtTest(3).Text = "This <B>paragraph</B> has percentage based indentation.  The left hand side is 20% of the page width, " _
& "the right hand side is 30% of the page width.  The TextItem's alignment has been set for full justification." & vbCrLf & vbCrLf _
   & "<font=Times New Roman><size=15><color=#a1a1a1>This paragraph is part of the same TextItem as the paragraph above, but the " _
   & "font face and size have been altered.  No terminating style tags have been used so any further paragraphs would have the new attributes."

With qPrint
  .ResetItems
  .AppName = "qPrint Test Project"
  .ScaleMode = eMillimetre
  .MarginBottom = 20
  .MarginTop = 10
  .MarginLeft = 10
  .MarginRight = 10
End With
For iCount = 0 To 3
  With txtTest(iCount)
    qPrint.AddText .Text, .FontName, .FontSize, .FontBold, .FontItalic, .FontUnderline, .ForeColor
  End With
Next
With qPrint
' Set up text item values
  .TextItem(1).Alignment = eCentre
  .TextItem(1).ShowBorder = True
  .TextItem(1).BorderShading = txtTest(0).BackColor

  .TextItem(2).Alignment = eJustify
  .TextItem(3).Alignment = eRight
  .TextItem(4).Alignment = eJustify
  .TextItem(4).ScaleMode = ePercentage
  .TextItem(4).indentleft = 20
  .TextItem(4).indentright = 30
  .TextItem(4).ShowBorder = True
  .TextItem(4).BorderColor = 0
  .TextItem(4).BorderShading = RGB(200, 200, 200)
  .TextItem(4).BorderLine = 1
  .AddText "qbd software", "Verdana", "48", , , , RGB(0, 0, 0), eCentre
  .TextItem(5).Absolute = True
  .TextItem(5).AbsPage = 1
  .TextItem(5).Top = 11
  .TextItem(5).indentleft = 20
  .TextItem(5).indentright = 20
  .TextItem(5).ShowBorder = True
  .TextItem(5).BorderColor = 0
  .TextItem(5).BorderShading = RGB(200, 155, 155)
  .TextItem(5).GetSize
  .TextItem(1).Top = .TextItem(5).Height + .TextItem(5).Top + 5
  .TextItem(2).Top = 10
  
  .Header(EvenPage_hf).Text = txtHeader.Text & vbCrLf
  .Header(EvenPage_hf).FontName = txtHeader.FontName
  .Header(EvenPage_hf).FontSize = txtHeader.FontSize
  .Header(EvenPage_hf).Alignment = eCentre
  .Headercopy EvenPage_hf, OddPage_hf
  .Headercopy EvenPage_hf, FirstPage_hf
  .SetHeader(EvenPage_hf) = True
  .SetHeader(OddPage_hf) = True
  .Footer(EvenPage_hf).Text = vbCrLf & txtFooter.Text
  .Footer(EvenPage_hf).FontName = txtFooter.FontName
  .Footer(EvenPage_hf).FontSize = txtFooter.FontSize
  .Footer(EvenPage_hf).Alignment = eCentre
  .FooterCopy EvenPage_hf, OddPage_hf
  .FooterCopy EvenPage_hf, LastPage_hf
  .Footer(LastPage_hf).Alignment = eLeft
  .Footer(LastPage_hf).Text = "Page #pagenumber# of #pagetotal#<FORCE><ALIGN=Right>Printed at #shorttime# on #shortdate#"
  .SetFooter(OddPage_hf) = True
  .SetFooter(EvenPage_hf) = True
  .PrintOptions = ShowPrinter_po + ShowCopies_po
  
 
  optOrientation(1).Enabled = .OrientOkay
  optOrientation(2).Enabled = .OrientOkay
  If .OrientOkay Then
  lblOStatus.Caption = "Available"
  Else
  lblOStatus.Caption = "Not Available"
  End If
  optOrientation(.Orientation).Value = True

End With
End Sub



Private Sub optOrientation_Click(Index As Integer)

qPrint.Orientation = Index

End Sub


