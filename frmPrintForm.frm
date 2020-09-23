VERSION 5.00
Begin VB.Form frmPrintForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "qPrinter:FormPrint example"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmPrintForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5640
      TabIndex        =   33
      Top             =   6000
      Width           =   1095
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
      TabIndex        =   32
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame fraBorder 
      Caption         =   "Borders"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2880
      TabIndex        =   5
      Top             =   5400
      Width           =   2655
      Begin VB.CheckBox chkBorder 
         Appearance      =   0  'Flat
         Caption         =   "Show Input borders"
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
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label lblBorder 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Displays borders around Textbox/List/Combo items."
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.Frame fraAlign 
      Caption         =   "Form Alignment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   2655
      Begin VB.OptionButton optAlign 
         Appearance      =   0  'Flat
         Caption         =   "Left"
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAlign 
         Appearance      =   0  'Flat
         Caption         =   "Centre"
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
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optAlign 
         Appearance      =   0  'Flat
         Caption         =   "Right"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optAlign 
         Appearance      =   0  'Flat
         Caption         =   "Full"
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
         Index           =   3
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label lblAlign 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "How the form will be shown horizontally on the page."
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
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.PictureBox picDocument 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4335
      ScaleWidth      =   6615
      TabIndex        =   10
      Top             =   1080
      Width           =   6615
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   23
         Text            =   "FirstName LastName"
         Top             =   0
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   1
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "frmPrintForm.frx":27A2
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   21
         Text            =   "0000 000 0000"
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   5280
         TabIndex        =   20
         Text            =   "4XJ - 5RV"
         Top             =   0
         Width           =   1335
      End
      Begin VB.ComboBox cboInfo 
         Appearance      =   0  'Flat
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
         ItemData        =   "frmPrintForm.frx":27DB
         Left            =   1440
         List            =   "frmPrintForm.frx":27EE
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2400
         Width           =   2175
      End
      Begin VB.ListBox lstInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   3960
         TabIndex        =   18
         Top             =   900
         Width           =   2655
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add Item"
         Height          =   300
         Left            =   5520
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "Full"
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
         Index           =   0
         Left            =   1440
         TabIndex        =   16
         Top             =   2880
         Width           =   615
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "Upgrade from version 3"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   3120
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CheckBox chkInform 
         Appearance      =   0  'Flat
         Caption         =   "Inform me of new products"
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
         Left            =   3960
         TabIndex        =   14
         Top             =   2880
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.OptionButton optType 
         Appearance      =   0  'Flat
         Caption         =   "Competitive Upgrade"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   4
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "frmPrintForm.frx":2863
         Top             =   3720
         Width           =   5175
      End
      Begin VB.CheckBox chkExpress 
         Appearance      =   0  'Flat
         Caption         =   "Express Delivery"
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
         Left            =   3960
         TabIndex        =   11
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label lblInfo 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   31
         Top             =   15
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   30
         Top             =   503
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         Caption         =   "Telephone:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   29
         Top             =   1935
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Caption         =   "Reference:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   28
         Top             =   15
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         Caption         =   "Order List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   27
         Top             =   503
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Caption         =   "Use:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   26
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Caption         =   "License:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   25
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Caption         =   "Comments:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   24
         Top             =   3720
         Width           =   1335
      End
   End
   Begin VB.Label lblInformation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmPrintForm.frx":2935
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
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmPrintForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================
' Project:      Enhance Print object Form_Print Test Form
' Author:       edward moth
' Copyright:    © 2000-2001 qbd software ltd
'
' ==============================================================
' Module:       frmPrintForm
' Purpose:      Test printing Labels and Textboxes on a form
' ==============================================================


Option Explicit
Dim qPrint As qcPrinter
Dim eAlignment As qePrinterAlign

Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub Form_Load()

'Set up the information on the form
lstInfo.AddItem "Item number 1"
lstInfo.AddItem "Item number 2"
lstInfo.AddItem "Item number 3"
cboInfo.ListIndex = 1
eAlignment = eJustify
'Initialise the qPrint object in Form_Load - updates are much quicker
Set qPrint = New qcPrinter
qPrint.MarginTop = 567
qPrint.MarginLeft = 567
qPrint.MarginRight = 567
' FormPrint parameters:
' frmPrint As Object: 'Me'
'     The form holding the controls
' Optional ParentContain As Object: 'picDocument'
'     The container holding the controls to be printed
'     The scalewidth of the container is used for positioning
'     if no container is specified, the Form (Me) is used
' Optional FormAlign As qePrinterAlign = eLeft: 'eAlignment'
'     How to display the form. Left/Right/Centre/Justify
'     'Jusitfy' will stretch the controls to fit the width of the page
' Optional InputBorder As Boolean = False: 'True'
'     Will display a border equivalent to the border of TextBox,
'     ListBox or ComboBox.  Labels, CheckBoxes and OptionButtons are
'     Not given a border.
' Optional TopOffset As Single: 800
'     Distance from Top of page to Offset the Control contents
'     In this instance to allow for the 'docTitle' added afterwards
' Optional AutoHeight As Boolean: True
'     Where the contents of an item are taller than the bounding
'     control, qPrinter will adjust the height and position of
'     subsequent items.
' Optional ExcludeList As String: "*chkExpand"
'     List of controls prefixed by * to identify them
'     to be excluded from the document.  If a container control
'     eg. PictureBox/Frame is included, all the controls in the
'     container will be excluded.  In this instance 'chkExpand'
'     is excluded because it is held in 'picDocument' - the container
'     we want to print.

qPrint.FormPrint Me, picDocument, eAlignment, True, 1000, True, "*chkExpand"
' Add the title
qPrint.AddText "qbd software ltd" & vbCrLf & "qPrinter:FormPrint example", "Verdana", 18, , , , , eCentre, , , "docTitle"


End Sub

Private Sub Form_Unload(Cancel As Integer)
'Destroy qPrint object
Set qPrint = Nothing

End Sub

Private Sub optAlign_Click(Index As Integer)

If Index <> eAlignment Then
eAlignment = Index
' Change the Alignment of the document
qPrint.FormPrint Me, picDocument, eAlignment, CBool(chkBorder.Value = vbChecked), 1000, True, "*chkExpand"
' Add the title
qPrint.AddText "qbd software ltd" & vbCrLf & "qPrinter:FormPrint example", "Verdana", 18, , , , , eCentre, , , "docTitle"
If eAlignment = eLeft Then
qPrint.TextItem("docTitle").indentright = qPrint.TextItem("lstInfo").indentright
ElseIf eAlignment = eRight Then
qPrint.TextItem("docTitle").indentleft = qPrint.TextItem("lblInfo/0").indentleft
End If
End If

End Sub

Private Sub chkBorder_Click()
' Change the borders of the document
qPrint.FormPrint Me, picDocument, eAlignment, CBool(chkBorder.Value = vbChecked), 1000, True, "*chkExpand"
' Add the title
qPrint.AddText "qbd software ltd" & vbCrLf & "qPrinter:FormPrint example", "Verdana", 18, , , , , eCentre, , , "docTitle"
If eAlignment = eLeft Then
qPrint.TextItem("docTitle").indentright = qPrint.TextItem("lstInfo").indentright
ElseIf eAlignment = eRight Then
qPrint.TextItem("docTitle").indentleft = qPrint.TextItem("lblInfo/0").indentleft
End If

End Sub

Private Sub cmdAddItem_Click()
lstInfo.AddItem "Item number " & lstInfo.ListCount + 1

End Sub

Private Sub cmdPreview_Click()
'Update text from the Form contents
qPrint.FormPrint_Update Me
qPrint.Preview

End Sub

