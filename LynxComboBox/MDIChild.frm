VERSION 5.00
Begin VB.Form frmMDIChild 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4215
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   2910
      TabIndex        =   2
      Top             =   1560
      Width           =   1245
   End
   Begin LynxComboBoxTest.LynxComboBox cboTarget 
      Height          =   345
      Left            =   1050
      TabIndex        =   0
      Top             =   120
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
      Left            =   90
      TabIndex        =   1
      Top             =   210
      Width           =   780
   End
End
Attribute VB_Name = "frmMDIChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
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



    
