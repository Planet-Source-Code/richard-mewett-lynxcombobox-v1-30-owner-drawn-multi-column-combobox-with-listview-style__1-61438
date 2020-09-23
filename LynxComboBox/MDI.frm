VERSION 5.00
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "MDI Test"
   ClientHeight    =   3465
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5685
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   0
      Width           =   5685
      Begin LynxComboBoxTest.LynxComboBox cboTarget 
         Height          =   345
         Left            =   1020
         TabIndex        =   1
         Top             =   90
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ThemeStyle      =   0
         Style           =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Target OS:"
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   180
         Width           =   780
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    With cboTarget
        .ColWidth(0) = 2000
        
        .AddItem "Windows 95"
        .AddItem "Windows 98"
        .AddItem "Windows NT4"
        .AddItem "Windows ME"
        .AddItem "Windows 2000", , True
        .AddItem "Windows XP"
        .Refresh
    End With
End Sub


Private Sub mnuOptions_Click()
    frmMDIChild.Show vbModeless
End Sub


