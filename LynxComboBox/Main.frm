VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "LynxComboBox Tester © 2005 Richard Mewett"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   10815
      TabIndex        =   0
      Top             =   0
      Width           =   10875
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   60
         Width           =   1035
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multi-Column ComboBox UserControl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   2
         Top             =   60
         Width           =   3090
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single File solution to replace a standard VB ComboBox"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   3
         Top             =   300
         Width           =   4740
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sample Properties"
      Height          =   7860
      Left            =   7560
      TabIndex        =   23
      Top             =   660
      Width           =   3285
      Begin VB.ComboBox cboThemeStyle 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2310
         Width           =   1695
      End
      Begin VB.TextBox txtDropDownItems 
         Height          =   315
         Left            =   2070
         TabIndex        =   40
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtTextSelection 
         Height          =   315
         Left            =   1410
         TabIndex        =   48
         Top             =   4800
         Width           =   1155
      End
      Begin VB.TextBox txtTextNone 
         Height          =   315
         Left            =   1410
         TabIndex        =   46
         Top             =   4410
         Width           =   1155
      End
      Begin VB.TextBox txtTextAll 
         Height          =   315
         Left            =   1410
         TabIndex        =   44
         Top             =   4020
         Width           =   1155
      End
      Begin VB.CheckBox chkAutoComplete 
         Caption         =   "AutoComplete"
         Height          =   195
         Left            =   1620
         TabIndex        =   25
         Top             =   270
         Width           =   1395
      End
      Begin VB.ComboBox cboFocusRect 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   2730
         Width           =   1185
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "Apply"
         Height          =   345
         Left            =   1950
         TabIndex        =   63
         Top             =   7410
         Width           =   1245
      End
      Begin VB.ComboBox cboBorderStyle 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1920
         Width           =   1185
      End
      Begin VB.CheckBox chkLocked 
         Caption         =   "Locked"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   750
         Width           =   1035
      End
      Begin VB.TextBox txtRowHeightMin 
         Height          =   315
         Left            =   2070
         TabIndex        =   42
         Top             =   3630
         Width           =   495
      End
      Begin VB.CheckBox chkColumnHeaders 
         Caption         =   "Column Headers"
         Height          =   195
         Left            =   1620
         TabIndex        =   27
         Top             =   510
         Width           =   1575
      End
      Begin VB.CheckBox chkDisplayEllipsis 
         Caption         =   "Display Ellipsis"
         Height          =   195
         Left            =   1620
         TabIndex        =   29
         Top             =   750
         Width           =   1575
      End
      Begin VB.OptionButton optOptionButtons 
         Caption         =   "OptionButtons"
         Height          =   195
         Left            =   90
         TabIndex        =   32
         Top             =   1560
         Width           =   1665
      End
      Begin VB.OptionButton optCheckBoxes 
         Caption         =   "CheckBoxes"
         Height          =   195
         Left            =   90
         TabIndex        =   31
         Top             =   1350
         Width           =   1665
      End
      Begin VB.OptionButton optStandard 
         Caption         =   "Standard"
         Height          =   195
         Left            =   90
         TabIndex        =   30
         Top             =   1140
         Width           =   1665
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   195
         Left            =   90
         TabIndex        =   26
         Top             =   510
         Width           =   1035
      End
      Begin VB.CheckBox chkEditable 
         Caption         =   "Editable"
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label lblThemeStyle 
         AutoSize        =   -1  'True
         Caption         =   "ThemeStyle"
         Height          =   195
         Left            =   90
         TabIndex        =   35
         Top             =   2370
         Width           =   840
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "DropDownItemsVisible"
         Height          =   195
         Left            =   90
         TabIndex        =   39
         Top             =   3270
         Width           =   1590
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "TextSelection"
         Height          =   195
         Left            =   90
         TabIndex        =   47
         Top             =   4830
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "TextNone"
         Height          =   195
         Left            =   90
         TabIndex        =   45
         Top             =   4440
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TextAll"
         Height          =   195
         Left            =   90
         TabIndex        =   43
         Top             =   4050
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FocusRect"
         Height          =   195
         Left            =   90
         TabIndex        =   37
         Top             =   2790
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "BorderStyle"
         Height          =   195
         Left            =   90
         TabIndex        =   33
         Top             =   1980
         Width           =   810
      End
      Begin VB.Label lblHotBorderColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   60
         Top             =   6750
         Width           =   795
      End
      Begin VB.Label lblBackColorText 
         AutoSize        =   -1  'True
         Caption         =   "HotBorderColor"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   59
         Top             =   6780
         Width           =   1080
      End
      Begin VB.Label lblBorderColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   56
         Top             =   6150
         Width           =   795
      End
      Begin VB.Label lblBackColorText 
         AutoSize        =   -1  'True
         Caption         =   "BorderColor"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   55
         Top             =   6165
         Width           =   825
      End
      Begin VB.Label lblFocusRectColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   54
         Top             =   5850
         Width           =   795
      End
      Begin VB.Label lblBackColorText 
         AutoSize        =   -1  'True
         Caption         =   "FocusRectColor"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   53
         Top             =   5865
         Width           =   1140
      End
      Begin VB.Label lblHotButtonbackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   62
         Top             =   7050
         Width           =   795
      End
      Begin VB.Label lblBackColorText 
         AutoSize        =   -1  'True
         Caption         =   "HotButtonBackColor"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   61
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label lblButtonBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   58
         Top             =   6450
         Width           =   795
      End
      Begin VB.Label lblBackColorText 
         AutoSize        =   -1  'True
         Caption         =   "ButtonBackColor"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   57
         Top             =   6465
         Width           =   1200
      End
      Begin VB.Label lblForeColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   52
         Top             =   5550
         Width           =   795
      End
      Begin VB.Label lblBackColorText 
         AutoSize        =   -1  'True
         Caption         =   "ForeColorEdit"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   51
         Top             =   5550
         Width           =   945
      End
      Begin VB.Label lblBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         TabIndex        =   50
         Top             =   5250
         Width           =   795
      End
      Begin VB.Label lblBackColorText 
         AutoSize        =   -1  'True
         Caption         =   "BackColorEdit"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   49
         Top             =   5250
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RowHeightMin"
         Height          =   195
         Left            =   90
         TabIndex        =   41
         Top             =   3660
         Width           =   1050
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   1770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1C02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame2 
      Height          =   7860
      Left            =   30
      TabIndex        =   4
      Top             =   660
      Width           =   7455
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "Remove Item"
         Height          =   345
         Left            =   3720
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   300
         Width           =   1245
      End
      Begin VB.CommandButton cmdMDITest 
         Caption         =   "MDI Test...."
         Height          =   345
         Left            =   6030
         TabIndex        =   64
         Top             =   5670
         Width           =   1245
      End
      Begin VB.ListBox lstHistory 
         Height          =   1230
         Left            =   1950
         TabIndex        =   21
         Top             =   3300
         Width           =   5295
      End
      Begin VB.TextBox txtHelp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   1635
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "Main.frx":219C
         Top             =   6090
         Width           =   7095
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load!"
         Height          =   345
         Left            =   3720
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1305
      End
      Begin LynxComboBoxTest.LynxComboBox cboSimple 
         Height          =   315
         Left            =   1950
         TabIndex        =   6
         Top             =   330
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
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
         BorderCurve     =   15
         BorderStyle     =   3
         AutoComplete    =   -1  'True
         Editable        =   -1  'True
      End
      Begin LynxComboBoxTest.LynxComboBox cboMultiColumn 
         Height          =   315
         Left            =   1950
         TabIndex        =   8
         Top             =   720
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonBackColor =   16761024
         DisplayEllipsis =   -1  'True
         ColumnResize    =   -1  'True
         ColumnSort      =   -1  'True
         ColumnHeaders   =   -1  'True
         DropDownAutoWidth=   -1  'True
         DropDownItemsVisible=   12
         RowHeightMin    =   285
         Style           =   1
      End
      Begin LynxComboBoxTest.LynxComboBox cboAutoComplete 
         Height          =   420
         Left            =   1950
         TabIndex        =   20
         Top             =   2790
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         ForeColor       =   8421631
         BorderStyle     =   2
         AutoComplete    =   -1  'True
         Editable        =   -1  'True
      End
      Begin LynxComboBoxTest.LynxComboBox cboSort 
         Height          =   315
         Left            =   1950
         TabIndex        =   10
         Top             =   1110
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
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
         BorderStyle     =   3
         FocusRectStyle  =   2
         ColumnResize    =   -1  'True
         ColumnSort      =   -1  'True
         ColumnHeaders   =   -1  'True
         DropDownAutoWidth=   -1  'True
         RowHeightMin    =   350
      End
      Begin LynxComboBoxTest.LynxComboBox cbbBorder 
         Height          =   315
         Left            =   1950
         TabIndex        =   15
         Top             =   1950
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
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
         ArrowColor      =   16761024
         BorderColor     =   16761024
         BorderCurve     =   15
         BorderStyle     =   4
         ButtonBackColor =   12632319
         HotArrowColor   =   16744576
         HotBorderColor  =   16744576
         HotButtonBackColor=   255
         AutoComplete    =   -1  'True
      End
      Begin LynxComboBoxTest.LynxComboBox cboUnicode 
         Height          =   360
         Left            =   1950
         TabIndex        =   12
         Top             =   1500
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderCurve     =   15
         BorderStyle     =   3
         AutoComplete    =   -1  'True
         DropDownAutoWidth=   -1  'True
         IntegralHeight  =   -1  'True
      End
      Begin LynxComboBoxTest.LynxComboBox cboLargeList 
         Height          =   315
         Left            =   1950
         TabIndex        =   17
         Top             =   2370
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
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
         BorderColor     =   8454143
         BorderCurve     =   15
         BorderStyle     =   4
         ButtonBackColor =   16744576
         HotBorderColor  =   65535
         HotButtonBackColor=   12582912
         AutoComplete    =   -1  'True
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Auto-Complete"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Top             =   2880
         Width           =   1035
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "50,000 Items"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Using ""@Arial Unicode MS"" Font"
         Height          =   390
         Left            =   5070
         TabIndex        =   13
         Top             =   1500
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Unicode"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   1530
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Custom Border"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   1980
         Width           =   1035
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Custom Sort"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Multi-Column"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Simple"
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   360
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#############################################################################################################################
'Demo of LynxComboBox UserControl
'
'This control can be used as an alternative to a ComboBox. It combines
'ComboBox operation with the ListBox Checkbox (Style=Checked) functionality

'When in Multi Column mode & Displaying Column Headings it looks like a ListView
'hence the name LynxComboBox

'I have attempted to replicate the behaviour/properties of the standard Combo
'where possible but there are some (intentional) differences

'Item:          VB:                     UserControl:
'Property       ListBox.Selected        ItemChecked
'Property       ComboBox.Style          Editable

'Key Features:
'Multiple Columns
'Column Sorting (by mouse-click on list (ColumnSort Property) or Sort Method)
'CheckBox & OptionButton modes (Style Property)
'Item Formatting (ItemForeColor, ItemFontBold)
'Border Styles (BorderStyle - Raised, Sunken, Flat & None)
'Adjustable Dropdown Height (ItemsVisible Property)
'#############################################################################################################################

Private mLynxComboBox As LynxComboBox
Private Sub LoadMultiColumnCombo()
    Dim nCount As Integer
    Dim nIndex As Integer
    
    With cboMultiColumn
        .Clear
        
        'FormatString creates a Column for each item in the string.
        '> Right Justify
        '< Left Justify
        '^ Centre Justify
        .FormatString = "<Code|^G|<Forename|<Surname|<DOB"
        
        'Set the Column Widths
        .ColWidth(0) = 1250
        .ColWidth(1) = 350
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        
        'ColType property allows Sort to process values correctly
        .ColType(4) = lcDate
        
        'Set images
        .ImageList = ImageList1
        
        'Override column alignment from FormatString
        .ColAlignment(0) = lcAlignLeftCenter
        .ColAlignment(1) = lcAlignCenterCenter
        .ColAlignment(2) = lcAlignLeftCenter
        .ColAlignment(3) = lcAlignLeftCenter
        .ColAlignment(4) = lcAlignLeftCenter
        
        For nCount = 1 To 200
            .AddItem Format$("XD" & Format$(nCount, "000"))
            
            If RandomInt(0, 1) = 0 Then
                .ItemImage(.NewIndex) = RandomInt(1, 3)
                .ItemText(.NewIndex, 1) = "M"
                .ItemText(.NewIndex, 2) = GetForename(ntMale)
                .ItemForeColor(.NewIndex) = vbBlue
            Else
                .ItemImage(.NewIndex) = RandomInt(3, 6)
                .ItemText(.NewIndex, 1) = "F"
                .ItemText(.NewIndex, 2) = GetForename(ntFemale)
                .ItemForeColor(.NewIndex) = vbRed
            End If
            .ItemText(.NewIndex, 3) = GetSurname()
            .ItemText(.NewIndex, 4) = DateSerial(RandomInt(1930, 1990), RandomInt(1, 12), RandomInt(1, 28))
        Next nCount
        
        .ListIndex = 0
    End With
End Sub


Private Sub LoadAutoComplete()
    Dim lCount As Long
    Dim bExists As Boolean
    Dim sData As String
    
    With cboAutoComplete
        .Clear
        .ColWidth(0) = .Width - 250
        
        Do Until .ListCount = 99
            If RandomInt(0, 1) = 0 Then
                sData = GetNameOfPerson(ntMale)
            Else
                sData = GetNameOfPerson(ntFemale)
            End If
            
            bExists = False
            For lCount = 0 To .ListCount - 1
                If .ItemText(lCount, 0) = sData Then
                    bExists = True
                    Exit For
                End If
            Next lCount
            
            If Not bExists Then
                .AddItem sData
            End If
        Loop
        
        'Sort the Items by name
        .Sort 0
        .ListIndex = 0
    End With
End Sub



Private Sub SetProperties(cvCombo As LynxComboBox)
    Set mLynxComboBox = cvCombo

    With mLynxComboBox
        chkEditable.Value = Abs(.Editable)
        chkEnabled.Value = Abs(.Enabled)
        chkLocked.Value = Abs(.Locked)
        
        If .Style = lcCheckboxes Then
            optCheckBoxes.Value = True
        ElseIf .Style = lcOptionButtons Then
            optOptionButtons.Value = True
        Else
            optStandard.Value = True
        End If
        
        chkAutoComplete.Value = Abs(.AutoComplete)
        chkColumnHeaders.Value = Abs(.ColumnHeaders)
        chkDisplayEllipsis.Value = Abs(.DisplayEllipsis)
        
        cboBorderStyle.ListIndex = .BorderStyle
        cboThemeStyle.ListIndex = .ThemeStyle
        cboFocusRect.ListIndex = .FocusRectStyle
        
        txtDropDownItems.Text = .DropDownItemsVisible
        txtRowHeightMin.Text = .RowHeightMin
        
        txtTextAll.Text = .TextAll
        txtTextNone.Text = .TextNone
        txtTextSelection.Text = .TextSelection
        
        lblBackColor.BackColor = .BackColorEdit
        lblForeColor.BackColor = .ForeColorEdit
        lblFocusRectColor.BackColor = .FocusRectColor
        lblBorderColor.BackColor = .BorderColor
        lblButtonBackColor.BackColor = .ButtonBackColor
        lblHotBorderColor.BackColor = .HotBorderColor
        lblHotButtonbackColor.BackColor = .HotButtonBackColor
    End With
End Sub

Private Sub cbbBorder_GotFocus()
    SetProperties cbbBorder
End Sub


Private Sub cboAutoComplete_AutoCompleteSearch(ListIndex As Long)
    If ListIndex >= 0 Then
        lstHistory.AddItem "AutoComplete Success"
    Else
        lstHistory.AddItem "AutoComplete Failed"
    End If
End Sub

Private Sub cboLargeList_GotFocus()
    SetProperties cboLargeList
End Sub


Private Sub cboMultiColumn_GotFocus()
    SetProperties cboMultiColumn
End Sub

Private Sub cboAutoComplete_GotFocus()
    SetProperties cboAutoComplete
End Sub


Private Sub cboSimple_GotFocus()
    SetProperties cboSimple
End Sub


Private Sub cboSort_CustomSort(Ascending As Boolean, Col As Long, Value1 As String, Value2 As String, bSwap As Boolean)
    'Simple Demo of Custom Sort. This Event is fired for each sort comparison &
    'the Swap value determines whether we change the Sort Sequence
    
    'In this example I am comparing only numeric part of the data
    
    If Ascending Then
        bSwap = (Mid$(Value1, 3) > Mid$(Value2, 3))
    Else
        bSwap = (Mid$(Value1, 3) < Mid$(Value2, 3))
    End If
End Sub


Private Sub cboSort_GotFocus()
    SetProperties cboSort
End Sub


Private Sub cboUnicode_GotFocus()
    SetProperties cboUnicode
End Sub


Private Sub cmdApply_Click()
   With mLynxComboBox
        .Editable = chkEditable.Value
        .Enabled = chkEnabled.Value
        .Locked = chkLocked.Value
        
        If optCheckBoxes.Value Then
            .Style = lcCheckboxes
        ElseIf optOptionButtons.Value Then
            .Style = lcOptionButtons
        Else
            .Style = lcStandard
        End If
        
        .AutoComplete = chkAutoComplete.Value
        .ColumnHeaders = chkColumnHeaders.Value
        .DisplayEllipsis = chkDisplayEllipsis.Value
        
        .BorderStyle = cboBorderStyle.ListIndex
        .ThemeStyle = cboThemeStyle.ListIndex
        .FocusRectStyle = cboFocusRect.ListIndex
        
        .DropDownItemsVisible = Val(txtDropDownItems.Text)
        .RowHeightMin = Val(txtRowHeightMin.Text)
        
        .TextAll = txtTextAll.Text
        .TextNone = txtTextNone.Text
        .TextSelection = txtTextSelection.Text
        
        .BackColorEdit = lblBackColor.BackColor
        .ForeColorEdit = lblForeColor.BackColor
        .FocusRectColor = lblFocusRectColor.BackColor
        .BorderColor = lblBorderColor.BackColor
        .ButtonBackColor = lblButtonBackColor.BackColor
        .HotBorderColor = lblHotBorderColor.BackColor
        .HotButtonBackColor = lblHotButtonbackColor.BackColor
        
        .Refresh
    End With
End Sub

Private Function GetColor(NewValue As Long) As Long
    On Local Error GoTo SetBCError

    With CommonDialog1
        .Flags = cdlCCRGBInit
        .Color = NewValue
        .ShowColor
        
        GetColor = .Color
    End With
    Exit Function
    
SetBCError:
    GetColor = NewValue
    Exit Function
End Function


Private Sub cmdLoad_Click()
    Dim lRow As Long
    
    Screen.MousePointer = vbHourglass
    
    With cboLargeList
        'Set a larger pre-allocation buffer for faster loading
        .CacheIncrement = 10000
    
        For lRow = 1 To 50000
            .AddItem CStr(lRow)
        Next lRow
        .ListIndex = 0
        .Refresh
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdMDITest_Click()
    frmMDI.Show vbModeless
End Sub

Private Sub cmdRemoveItem_Click()
    With cboSimple
        If .ListCount > 0 Then
            .RemoveItem .ListIndex
        End If
    End With
End Sub


Private Sub Form_Activate()
    cboSimple.SetFocus
End Sub

Private Sub Form_Load()
    Dim lRow As Long
    Dim nCount As Integer
    Dim nIndex As Integer
    Dim sText As String
    
    Randomize Timer
    
    With cboBorderStyle
        .AddItem "None"
        .AddItem "Sunken"
        .AddItem "Raised"
        .AddItem "Flat"
        .AddItem "Custom"
    End With
    
    With cboThemeStyle
        .AddItem "Windows3D"
        .AddItem "WindowsFlat"
        .AddItem "WindowsXP"
        .AddItem "OfficeXP"
    End With
    
    With cboFocusRect
        .AddItem "None"
        .AddItem "Light"
        .AddItem "Heavy"
    End With
    
    '####################################################################################
    With cboSimple
        For nCount = 1 To 8
            .AddItem "PART" & Format$(nCount, "000")
        Next nCount
        
        .ListIndex = 0
    End With
    
    '####################################################################################
    'Custom Sorting
    With cboSort
        .FormatString = "<Text Sort|<Custom Sort"
        .ColWidth(0) = 2000
        .ColWidth(1) = 2000
        .ColAlignment(0) = lcAlignLeftCenter
        .ColAlignment(1) = lcAlignLeftCenter
        .ColType(1) = lcCustom
        
        'The Tab character is used to seperate the columns
        For nCount = 1 To 50
            sText = Chr$(RandomInt(65, 90)) & Chr$(RandomInt(65, 90)) & Format$(nCount, "000")
            .AddItem sText & vbTab & sText
        Next nCount
        
        .ListIndex = 0
    End With
    
    '####################################################################################
    LoadMultiColumnCombo
    LoadAutoComplete
    
    '####################################################################################
    'Unicode
    With cboUnicode
        .ColWidth(0) = .Width
        For nCount = 1 To 16
            .AddItem LoadResString(101 + nCount)
        Next nCount
        
        .ListIndex = 0
    End With
    
    '####################################################################################
    'Custom Border
    With cbbBorder
        .AddItem "One"
        .AddItem "Two"
        .AddItem "Three"
        .AddItem "Four"
        .AddItem "Five"
        .AddItem "Six"
        .ListIndex = 4
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mLynxComboBox = Nothing
End Sub

Private Sub lblBackColor_Click()
    lblBackColor.BackColor = GetColor(lblBackColor.BackColor)
End Sub

Private Sub lblBorderColor_Click()
    lblBorderColor.BackColor = GetColor(lblBorderColor.BackColor)
End Sub

Private Sub lblButtonBackColor_Click()
    lblButtonBackColor.BackColor = GetColor(lblButtonBackColor.BackColor)
End Sub


Private Sub lblFocusRectColor_Click()
    lblFocusRectColor.BackColor = GetColor(lblFocusRectColor.BackColor)
End Sub

Private Sub lblForeColor_Click()
    lblForeColor.BackColor = GetColor(lblForeColor.BackColor)
End Sub


Private Sub lblHotBorderColor_Click()
    lblHotBorderColor.BackColor = GetColor(lblHotBorderColor.BackColor)
End Sub

Private Sub lblHotButtonbackColor_Click()
    lblHotButtonbackColor.BackColor = GetColor(lblHotButtonbackColor.BackColor)
End Sub


Private Sub txtDropDownItems_KeyPress(KeyAscii As Integer)
    KeyAscii = NumericFilter(KeyAscii)
End Sub


Private Sub txtRowHeightMin_KeyPress(KeyAscii As Integer)
    KeyAscii = NumericFilter(KeyAscii)
End Sub


Public Function NumericFilter(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyDelete
        NumericFilter = KeyAscii
    
    Case 45 '-
        NumericFilter = KeyAscii
    
    Case 46 '.
        NumericFilter = KeyAscii
    
    Case 48 To 57 '0-9
        NumericFilter = KeyAscii
        
    Case 58 ':
        NumericFilter = KeyAscii
    
    End Select
End Function





