VERSION 5.00
Begin VB.Form frmPreview 
   AutoRedraw      =   -1  'True
   Caption         =   "Print Preview"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7605
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   7605
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.PictureBox picPages 
         AutoRedraw      =   -1  'True
         Height          =   315
         Left            =   4440
         ScaleHeight     =   255
         ScaleWidth      =   1155
         TabIndex        =   14
         Top             =   60
         Width           =   1215
         Begin VB.Label lblStatus 
            Caption         =   "Page:"
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
            Left            =   30
            TabIndex        =   15
            Top             =   15
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Enabled         =   0   'False
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
         Left            =   3480
         TabIndex        =   13
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "P&revious"
         Enabled         =   0   'False
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
         Left            =   2520
         TabIndex        =   12
         Top             =   60
         Width           =   855
      End
      Begin VB.ComboBox cboZoom 
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
         ItemData        =   "frmPreview.frx":27A2
         Left            =   600
         List            =   "frmPreview.frx":27A4
         TabIndex        =   10
         Text            =   "cboZoom"
         Top             =   60
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print ..."
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
         Left            =   5760
         TabIndex        =   2
         Top             =   60
         Width           =   855
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
         Left            =   6720
         TabIndex        =   1
         Top             =   60
         Width           =   855
      End
      Begin VB.Label lblView 
         Caption         =   "View:"
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
         Left            =   60
         TabIndex        =   11
         Top             =   90
         Width           =   495
      End
   End
   Begin VB.PictureBox picScroll 
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5595
      ScaleWidth      =   7515
      TabIndex        =   3
      Top             =   480
      Width           =   7575
      Begin VB.VScrollBar vsPreview 
         Height          =   1215
         Left            =   3000
         TabIndex        =   9
         Top             =   840
         Width           =   255
      End
      Begin VB.HScrollBar hsPreview 
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   4560
         Width           =   1725
      End
      Begin VB.PictureBox picShow 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   240
         ScaleHeight     =   3615
         ScaleWidth      =   3540
         TabIndex        =   4
         Top             =   480
         Width           =   3540
         Begin VB.PictureBox picHold 
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   240
            ScaleHeight     =   1815
            ScaleWidth      =   2175
            TabIndex        =   6
            Top             =   120
            Width           =   2175
            Begin VB.PictureBox picDoc 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               ForeColor       =   &H80000008&
               Height          =   1215
               Left            =   240
               ScaleHeight     =   1215
               ScaleWidth      =   1695
               TabIndex        =   7
               Top             =   240
               Visible         =   0   'False
               Width           =   1695
            End
         End
         Begin VB.PictureBox picNormal 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   480
            ScaleHeight     =   1215
            ScaleWidth      =   1695
            TabIndex        =   5
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================
' Project:      Enhance Print object
' Author:       edward moth
' Copyright:    © 2000-2001 qbd software ltd
'
' ==============================================================
' Module:       frmPreview
' Purpose:      Display Print Preview
' ==============================================================


Option Explicit
Private mDocument As qcPrinter
Private bScrollCode As Boolean
Private sZoom As Single
Private lPage As Integer
Private lPageMax As Integer
Private bDisplayPage As Boolean


Public Property Set Document(ByVal vNewValue As qcPrinter)

Set mDocument = vNewValue

End Property



Private Sub cboZoom_Click()


Dim iEvents As Integer

If Not bScrollCode Then
  If cboZoom.ListIndex >= 0 Then
' Because the Zoom_Check procedure can take some time
' the following line will close the dropdown
    iEvents = DoEvents
    If cboZoom.ItemData(cboZoom.ListIndex) <> sZoom Then
      sZoom = cboZoom.ItemData(cboZoom.ListIndex)
      Zoom_Check
    End If
  End If
End If

End Sub

Private Sub cboZoom_KeyPress(KeyAscii As Integer)

Dim sNewZoom As Single

If KeyAscii = 13 Then
sNewZoom = Val(cboZoom.Text)
If sNewZoom > 0 And sNewZoom <= 200 Then
cboZoom.Text = sNewZoom & " %"
If sNewZoom = sZoom Then
Exit Sub
End If
sZoom = sNewZoom
Zoom_Check
Else
If cboZoom.ListIndex >= 0 Then
cboZoom.Text = cboZoom.List(cboZoom.ListIndex)
Else
cboZoom.Text = sZoom & " %"
End If
End If
End If

End Sub

Private Sub cmdClose_Click()

Set mDocument = Nothing
Me.Hide

End Sub

Private Sub cmdNext_Click()

lPage = lPage + 1
Preview_Display lPage

End Sub

Private Sub cmdPrevious_Click()

lPage = lPage - 1
Preview_Display lPage


End Sub

Private Sub cmdPrint_Click()

Dim lPrintStart As Integer
Dim lPrintEnd As Integer
Dim iCopies As Integer
Dim bCollate As Boolean
Dim iPrinter As Integer

' Display the Print Options form
' Pass current page details

frmPrint.Flags = mDocument.PrintOptions
frmPrint.PageCurrent = lPage
frmPrint.PageMax = lPageMax
frmPrint.Show vbModal
lPrintStart = frmPrint.PageStart
lPrintEnd = frmPrint.PageEnd
iCopies = frmPrint.Copies
bCollate = frmPrint.Collate
iPrinter = frmPrint.PrinterNumber

If frmPrint.PrintDoc Then
  lblStatus.Caption = "Printing..."
  lblStatus.Refresh
  mDocument.PrintDoc lPrintStart, lPrintEnd, iCopies, bCollate, iPrinter
  
  
  lblStatus.Caption = "Page: " & lPage & " / " & lPageMax
End If
Unload frmPrint

End Sub

Private Sub Form_Activate()
Me.Refresh
bDisplayPage = True
Preview_Display 1

End Sub

Private Sub Form_Load()
sZoom = 100
With cboZoom
  .AddItem "100 %"
  .ItemData(.ListCount - 1) = 100
  .AddItem "75 %"
  .ItemData(.ListCount - 1) = 75
  .AddItem "50 %"
  .ItemData(.ListCount - 1) = 50
  .AddItem "Full Page"
  .ItemData(.ListCount - 1) = 0
  .AddItem "Full Width"
  .ItemData(.ListCount - 1) = -1
  bScrollCode = True
  .ListIndex = 0
  bScrollCode = False
End With
sZoom = 100
lPage = 1

lPageMax = mDocument.Pages

End Sub

Public Sub Preview_Display(ByVal iPage As Integer)

Dim iMin As Integer
Dim iMax As Integer
Screen.MousePointer = vbHourglass
picNormal.Cls
picDoc.Visible = False
mDocument.PreviewPage picNormal, iPage
Preview_Status
Zoom_Check
Screen.MousePointer = vbDefault
End Sub
Private Sub Zoom_Check()

Dim sSizeX As Single
Dim sSizeY As Single
Dim sRatio As Single
Dim spImage As StdPicture
Dim sWidth As Single
Dim sHeight As Single
Dim bScroll As Byte
Dim bOldScroll As Byte
Screen.MousePointer = vbHourglass

sWidth = picScroll.ScaleWidth
sHeight = picScroll.ScaleHeight
' Check the height and width to determine whether scroll bars
' are required.  This is in a loop because if a scroll bar is
' required it will affect the opposite dimension of the page
' display
Do
  bOldScroll = bScroll
  If sZoom = 0 Then
    sRatio = (sHeight - 480) / picNormal.Height
  ElseIf sZoom = -1 Then
    sRatio = (sWidth - 480) / picNormal.Width
  Else
    sRatio = sZoom / 100
  End If
  sSizeX = picNormal.Width * sRatio
  sSizeY = picNormal.Height * sRatio
  If sSizeX > sWidth And (bScroll And 1) <> 1 Then
    sHeight = sHeight - hsPreview.Height
    bScroll = bScroll + 1
  End If
  If sSizeY > sHeight And (bScroll And 2) <> 2 Then
    sWidth = sWidth - vsPreview.Width
    bScroll = bScroll + 2
  End If
Loop While bOldScroll <> bScroll

vsPreview.Height = sHeight
hsPreview.Width = sWidth

picShow.Move 0, 0, sWidth, sHeight
picDoc.Move 240, 240, sSizeX, sSizeY
picDoc.Cls
picDoc.PaintPicture picNormal.Image, 0, 0, sSizeX, sSizeY


' Display scroll bars if required
bScrollCode = True
picHold.Move 0, 0, sSizeX + 480, sSizeY + 480
If (bScroll And 2) = 2 Then
  vsPreview.Visible = True
  vsPreview.Max = (picHold.ScaleHeight - picShow.ScaleHeight) / 14.4 + 1
  vsPreview.Min = 0
  vsPreview.SmallChange = 14
  vsPreview.LargeChange = picShow.ScaleHeight / 14.4
  vsPreview.Value = vsPreview.Min
Else
  vsPreview.Visible = False
End If

If (bScroll And 1) = 1 Then
  hsPreview.Visible = True
  hsPreview.Max = (picHold.ScaleWidth - picShow.ScaleWidth) / 14.4 + 1
  hsPreview.Min = 0
  hsPreview.SmallChange = 14
  hsPreview.LargeChange = picShow.ScaleWidth / 14.4
  hsPreview.Value = hsPreview.Min
Else
  hsPreview.Visible = False
End If
bScrollCode = False
Screen.MousePointer = vbDefault
If bDisplayPage Then
picDoc.Visible = True
End If

End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
Exit Sub
End If
If Me.ScaleHeight > 600 Then
picScroll.Move 0, 600, Me.ScaleWidth, Me.ScaleHeight - 600
End If

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode <> vbFormCode Then
Me.Hide
Cancel = 1
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
bDisplayPage = False
End Sub

Private Sub hsPreview_Change()

If Not bScrollCode Then
  picHold.Left = -hsPreview.Value * 14.4
End If

End Sub

Private Sub picScroll_Resize()
vsPreview.Move picScroll.ScaleWidth - vsPreview.Width, 0, 255, picScroll.ScaleHeight
hsPreview.Move 0, picScroll.ScaleHeight - hsPreview.Height, picScroll.ScaleWidth
Zoom_Check

End Sub


Private Sub vsPreview_Change()
If Not bScrollCode Then
  picHold.Top = -vsPreview.Value * 14.4
End If

End Sub

Public Sub Preview_Status()
cmdPrevious.Enabled = CBool(lPage > 1)
cmdNext.Enabled = CBool(lPage < lPageMax)
picPages.Cls
lblStatus.Caption = "Page: " & lPage & " / " & lPageMax
lblStatus.Visible = True

End Sub


Private Sub vsPreview_GotFocus()
picScroll.SetFocus
End Sub

Private Sub hsPreview_GotFocus()
picScroll.SetFocus
End Sub

