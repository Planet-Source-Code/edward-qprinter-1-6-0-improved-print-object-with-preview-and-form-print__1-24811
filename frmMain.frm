VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "qPrinter: Test Project"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
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
      Left            =   3240
      TabIndex        =   3
      Top             =   120
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
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton optExample 
      Appearance      =   0  'Flat
      Caption         =   "Standard Document Example"
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
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.OptionButton optExample 
      Appearance      =   0  'Flat
      Caption         =   "FormPrint Example"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdShow_Click()

If optExample(0).Value Then
frmPrintForm.Show vbModal
Else
frmTest.Show vbModal
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim nCount As Integer
nCount = Forms.Count
Do While nCount > 0
nCount = nCount - 1
Unload Forms(nCount)
Loop

End Sub


