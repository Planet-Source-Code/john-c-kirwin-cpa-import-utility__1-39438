VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmImport_Utility 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   " Import Utility"
   ClientHeight    =   9615
   ClientLeft      =   5970
   ClientTop       =   2415
   ClientWidth     =   9615
   DrawStyle       =   5  'Transparent
   Icon            =   "frmImport_Utility.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   9615
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "x"
      CausesValidation=   0   'False
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
      Left            =   9255
      MaskColor       =   &H8000000F&
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Exit"
      Top             =   -75
      UseMaskColor    =   -1  'True
      Width           =   195
   End
   Begin VB.CommandButton cmdDelete 
      DisabledPicture =   "frmImport_Utility.frx":0442
      Enabled         =   0   'False
      Height          =   300
      Left            =   3240
      MaskColor       =   &H8000000F&
      Picture         =   "frmImport_Utility.frx":1B14
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Clear All"
      Top             =   8710
      Width           =   1550
   End
   Begin VB.CommandButton cmdInsert 
      DisabledPicture =   "frmImport_Utility.frx":31E6
      Enabled         =   0   'False
      Height          =   300
      Left            =   1680
      MaskColor       =   &H8000000F&
      Picture         =   "frmImport_Utility.frx":48B8
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Clear All"
      Top             =   8710
      Width           =   1550
   End
   Begin VB.CommandButton cmdReplace 
      DisabledPicture =   "frmImport_Utility.frx":5F8A
      Enabled         =   0   'False
      Height          =   300
      Left            =   125
      MaskColor       =   &H8000000F&
      Picture         =   "frmImport_Utility.frx":765C
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Clear All"
      Top             =   8710
      Width           =   1550
   End
   Begin VB.CommandButton cmdReset 
      DisabledPicture =   "frmImport_Utility.frx":8D2E
      Height          =   300
      Left            =   6375
      MaskColor       =   &H8000000F&
      Picture         =   "frmImport_Utility.frx":A400
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Clear All"
      Top             =   8710
      Width           =   1550
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      DisabledPicture =   "frmImport_Utility.frx":BAD2
      Height          =   300
      Left            =   7930
      MaskColor       =   &H8000000F&
      Picture         =   "frmImport_Utility.frx":D1A4
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Exit"
      Top             =   8710
      Width           =   1550
   End
   Begin VB.CommandButton cmdLayoutMaint 
      DisabledPicture =   "frmImport_Utility.frx":E876
      Height          =   300
      Left            =   4820
      MaskColor       =   &H8000000F&
      Picture         =   "frmImport_Utility.frx":FF48
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8710
      Width           =   1550
   End
   Begin VB.Frame fraStatusBar 
      Height          =   495
      Left            =   125
      TabIndex        =   19
      Top             =   8940
      Width           =   9375
      Begin VB.TextBox txtRecordType 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1900
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   150
         Width           =   2400
      End
      Begin VB.TextBox txtRecords 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7475
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   150
         Width           =   1850
      End
      Begin VB.TextBox txtCompany 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4320
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   150
         Width           =   765
      End
      Begin VB.TextBox txtImportType 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   50
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   150
         Width           =   1850
      End
      Begin VB.TextBox txtStatus 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5100
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   150
         Width           =   2350
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H8000000F&
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   100
      Picture         =   "frmImport_Utility.frx":1161A
      ScaleHeight     =   735
      ScaleWidth      =   9405
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   9400
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   9360
      TabIndex        =   10
      Top             =   1920
      Width           =   255
      Begin MSComctlLib.Slider sldFileMove 
         Height          =   7035
         Left            =   -25
         TabIndex        =   6
         Top             =   -100
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   12409
         _Version        =   393216
         Orientation     =   1
         Max             =   100000
      End
   End
   Begin VB.CommandButton cmdPrevious 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   3400
      Left            =   9000
      MaskColor       =   &H00000000&
      Picture         =   "frmImport_Utility.frx":27D70
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Scroll fields Backward"
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Default         =   -1  'True
      Height          =   3400
      Left            =   9000
      MaskColor       =   &H00000000&
      Picture         =   "frmImport_Utility.frx":295C2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Scroll fields Forward"
      Top             =   5310
      Width           =   375
   End
   Begin VB.ListBox lstLayout 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4830
      ItemData        =   "frmImport_Utility.frx":2ACAC
      Left            =   100
      List            =   "frmImport_Utility.frx":2ACB3
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   9
      ToolTipText     =   "Detail List of Import File Fields"
      Top             =   1950
      Width           =   8900
   End
   Begin VB.TextBox txtPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   100
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   7
      ToolTipText     =   "Original text of the Line(s) parsed into the fields listed below "
      Top             =   8200
      Width           =   8900
   End
   Begin VB.TextBox txtFieldInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1350
      Left            =   100
      MultiLine       =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Field Information"
      Top             =   6815
      Width           =   8900
   End
   Begin VB.ComboBox cboDelimeter 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      ItemData        =   "frmImport_Utility.frx":2ACC2
      Left            =   8760
      List            =   "frmImport_Utility.frx":2ACC4
      TabIndex        =   3
      ToolTipText     =   "Select a valid delimeter"
      Top             =   1100
      Width           =   735
   End
   Begin VB.TextBox txtImportFile 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Indicate File Name"
      Top             =   1100
      Width           =   5415
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   325
      Left            =   8280
      Picture         =   "frmImport_Utility.frx":2ACC6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1100
      Width           =   375
   End
   Begin VB.ComboBox cboImportType 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmImport_Utility.frx":2B1F8
      Left            =   120
      List            =   "frmImport_Utility.frx":2B1FA
      TabIndex        =   0
      ToolTipText     =   "Indicate the type of import file layout"
      Top             =   1100
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9360
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H8000000F&
      BorderWidth     =   2
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   9615
   End
   Begin VB.Label lblImportType 
      Caption         =   "Import Type"
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   900
      Width           =   2055
   End
   Begin VB.Label lblImportFile 
      Caption         =   "File Name"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   900
      Width           =   2415
   End
   Begin VB.Label lblDelimeter 
      Caption         =   "Delimeter"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   900
      Width           =   735
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H8000000F&
      BorderColor     =   &H8000000F&
      BorderWidth     =   3
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   6825
      Left            =   0
      Top             =   1920
      Width           =   9585
   End
   Begin VB.Label lblTotalFields 
      Caption         =   "Field"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   780
      TabIndex        =   14
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1500
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblValue 
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblRequired 
      Caption         =   "Req'd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   75
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000F&
      BorderColor     =   &H8000000F&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   -15
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   9630
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmImport_Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//**************************************************************************
'// Imports - Preview, validate, and edit delimited import files
'//
'// Version:       6.20.0
'// Created:       12/26/2002 John C. Kirwin (JCK)
'// Modified:      03/03/2002 JCK - Formated & Commented
'//**************************************************************************
'// Opens and validates an import file for preview and editing by displaying
'// the original file and each field vertically in a ListBox for previewing
'// and in a horizontal text box for editing.
'//
'// Funtionality:
'// -- Beta Design phase development & testing with the following import layouts:
'//    -- AR Posted Invoice
'//    -- AR Pending Invoice
'//    -- AP Posted Vouchers
'//    -- AP Pending Vouchers
'//    -- GL Natural Accounts
'//    -- GL Account Segments
'//    -- GL Accounts
'//    -- GL Transactions
'//    -- IM Items
'// -- Common Dialog File Lookup
'// -- Parse fields of a delimited import file into an array
'// -- Display current import file line fields to lstLayout ListBox
'// -- Display current import file line to txtPreview editable text box
'// -- Retrieve field information from validation definition Layouts.ini file
'//    imported into a dynamic array based on online help import layouts
'// -- Single record scroll buttons to assist record navigation
'// -- Multi record navigation by slider control
'// -- Required fields are indicated with checked boxes
'// -- Track lstLayout ListBox clicks to change record/field information displayed
'//    in txtFieldInfo
'// -- Track Error Count
'// -- Displays Import Company
'//
'// Enhancements:
'// -- Add arrow up/down,  +/-, and Page up/down keyboard navigation control
'// -- What about fixed length?  "They are the hardest to proofread!"
'// -- Validate and highlight errors based on validation definition file
'// -- AP Layouts (Vendor, Vend Address, Vend Contact, Vendor Purchase History,
'//                Vendor Transaction History, Posted Voucher, Voucher Batch)
'// -- AR Layouts (Customer, Cust Address, Cust Contact, Customer Sales History,
'//                Customer Transaction History)
'// -- GL Layouts (Budget Data)
'//
'// Bug Report:
'// -- Checking and unchecking does not seem to do anything
'// -- Required fields should not be removable in the lstLayout ListBox
'//
'// Issues and Challenges:
'// -- The index value of the List property for a ListBox must be from 0 to 32,766.
'// -- One size fits all at the moment (i.e. no screen resolution alternatives)
'// -- Does not connect to sources other than delimeted text file (i.e. no database
'//    or fixed length files)
'// -- Default Option Base of 0. The Option Base statement makes code reuse more
'//    difficult, especially if you like to cut and paste, because you have to
'//    be aware of what the code base was originally. The Option Base 1 sets the
'//    beginning indices at 1 instead of 0 the default.
'//
'//****************************************************************************

'//****************************************************************************
'// Windows API/Global Declarations
'//****************************************************************************

'//**** SetWindowRgn
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, _
                ByVal hRgn As Long, ByVal bRedraw As Long) As Long
'//**** CreateEllipticRgn
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, _
                ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'//**** GetPrivateProfileString
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                ByVal lpDefault As String, ByVal lpReturnedString As String, _
                ByVal nSize As Long, ByVal lpFileName As String) As Long

'//**** WritePrivateProfileString
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                ByVal lpString As Any, ByVal lpFileName As String) As Long

'//****************************************************************************
'// Constants
'//****************************************************************************
Const RGN_DIFF = 4

'//****************************************************************************
'// Global Variables
'//****************************************************************************
Dim M_Outer_Ring As Long
Dim M_Inner_Ring As Long
Dim M_Combined_Ring As Long
Dim M_Width As Single
Dim M_Height As Single
Dim M_Border_Width As Single
Dim M_Title_Height As Single

Dim sDelimiter As String                                                       ' Delimiter
Dim sFileName As String                                                        ' Import File Name
Dim sMessage As String
Dim sSection As String                                                         ' Import Type Section of Layout.ini file
Dim lBatchCount As Long                                                        ' Count the batches displayed
Dim lCurrentLine As Long                                                       ' Current line displayed in lstLayout
Dim lErrorCount As Long                                                        ' Count errors
Dim lFieldDisplay As Long                                                      ' Number of Fields to display in listbox
Dim lListMax As Long                                                           ' Max field number to display
Dim lTotalLines As Long                                                        ' Total lines in file
Dim bEdit As Boolean                                                           ' Import record changed? true/false

'//**** Status Bar Variables
Dim lFieldCount As Long                                                        ' Field Count
Dim lDataLine As Long                                                          ' Record Count
Dim sImportType As String                                                      ' Import Layout Type
Dim sRecordType As String                                                      ' Record Type
Dim sCompany As String                                                         ' Company extracted from Import File based on Layout
Dim sStatus As String                                                          ' Status
Dim iRows As Integer                                                           ' Total Rows in the Array
Dim iColumns As Integer                                                        ' Total Columns in the Array

'//**** Object Declarations
Private CPAFiles As cCPA_Files                                                ' CPAFiles Object
Private CPAINI As cCPA_INI                                                    ' CPAINI Object
Private CPAStrings As cCPA_Strings                                            ' CPAStrings Object
Private CPATracker As cCPA_Tracker                                            ' CPATracker Object
Private CPAErrorHandler As cCPA_ErrorHandler                                  ' CPAErrorHandler Object
Private CPAParseDelimited As cCPA_ParseDelimited                              ' CPAParseDelimited Object
Private CPALayouts As cCPA_Layouts                                            ' CPALayouts Object

'//**** General Ledger Arrays
Private aGLNRL() As String                                                     ' GL Natural Account Record Layout array
Private aGLSRL() As String                                                     ' GL Account Segment Record Layout array
Private aGLARL() As String                                                     ' GL Account Record Layout array
Private aGLBRL() As String                                                     ' GL Transactions Batch Header Record Layout array
Private aGLTRL() As String                                                     ' GL Transactions Record Layout array
Private aGLFRL() As String                                                     ' GL Transactions Batch Footer Record Layout array

'//**** Pending Voucher Arrays
Private aPVBRL() As String                                                     ' AP Pending Voucher Batch Record Layout array
Private aPVHRL() As String                                                     ' AP Pending Voucher Header Record Layout array
Private aPVDRL() As String                                                     ' AP Pending Voucher Detail Record Layout array

'//**** Posted Voucher Arrays
Private aVPBRL() As String                                                     ' AP Posted Voucher Batch Record Layout array
Private aVPHRL() As String                                                     ' AP Posted Voucher Header Record Layout array
Private aVPDRL() As String                                                     ' AP Posted Voucher Detail Record Layout array
Private aVPARL() As String                                                     ' AP Posted Voucher Application Record Layout array
Private aVPXRL() As String                                                     ' AP Posted Voucher Tax Header Record Layout array
Private aVPTRL() As String                                                     ' AP Posted Voucher Tax Detail Record Layout array

'//**** Pending Invoice Arrays
Private aPIBRL() As String                                                     ' AR Pending Invoice Batch Record Layout array
Private aPIHRL() As String                                                     ' AR Pending Invoice Header Record Layout array
Private aPIDRL() As String                                                     ' AR Pending Invoice Detail Record Layout array

'//**** Posted Invoice Arrays
Private aRIBRL() As String                                                     ' AR Posted Invoice Batch Record Layout array
Private aRIRRL() As String                                                     ' AR Posted Invoice Header Record Layout array
Private aRIDRL() As String                                                     ' AR Posted Invoice Detail Record Layout array
Private aRIARL() As String                                                     ' AR Posted Invoice Application Record Layout array
Private aRIERL() As String                                                     ' AR Posted Invoice Tax Header Record Layout array
Private aRITRL() As String                                                     ' AR Posted Invoice Tax Detail Record Layout array

'//**** Inventory Management Arrays
Private aIMIRL() As String                                                     ' IM Item Record Layout array
Private aIMURL() As String                                                     ' IM Item Unit Of Measure Record Layout
Private aIMLRL() As String                                                     ' IM Landed Cost Factor Record Layout array

Private Sub StatusBarClear()

  '//**************************************************************************
  '// StatusBarClear - Clear Status Bar text
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Clear Status Bar text

    txtRecordType.Text = ""
    txtRecords.Text = ""
    txtCompany = ""
    txtImportType = ""
    txtStatus = ""

    '//****

End Sub

Private Sub StatusBarClearVars()

  '//**************************************************************************
  '// StatusBarClear - Clear Status Bar Variables
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Clear Status Bar Variables

    lBatchCount = 0
    lFieldCount = 0
    lDataLine = 0
    lErrorCount = 0
    sImportType = ""
    sRecordType = ""
    sCompany = ""
    sStatus = ""

End Sub

Private Sub cmdNext_Click()

  '//**************************************************************************
  '// cmdNext_Click - Event to move forward 1 line
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '//**************************************************************************

    On Error GoTo EH

    '//**** Check if the currently displayed import record has changed
    If bEdit Then
        '//****
        If MsgBox(" The import record display has been changed. " & vbCr & vbCr & _
           " Save changes? ", vbYesNo + vbInformation) = vbYes Then
            cmdReplace_Click
            '//****
          Else
            '//****
            MsgBox "Import record display changed, but not saved"
            '//****
        End If
        '//****
    End If

    '//**** Get the total lines in the file
    lTotalLines = GetFileLines(sFileName)

    '//**** Check if there are any lines in the file
    If lTotalLines < 1 Then

        '//**** Since there are no lines the current line is 0
        lCurrentLine = 0

        '//**** Exit Sub/Function before error handler
        Exit Sub

        '//**** End if check if there are any lines in the file
    End If

    '//**** Increment current line variable lCurrentLine to indicate
    '//     next line gets displayed in lstLayout
    lCurrentLine = lCurrentLine + 1

    '//**** Check if out of range
    If lTotalLines < lCurrentLine Then
        lCurrentLine = lTotalLines
    End If

    '//****
    sldFileMove.Value = lCurrentLine

    '//****
    'ListLoad

    '//**** Exit Sub/Function before error handler

Exit Sub

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the cmdNext_Click"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Error Handling

Exit Sub

    '//****

End Sub

Private Sub cmdPrevious_Click()

  '//**************************************************************************
  '// cmdPrevious_Click - Event to move back 1 line
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

    On Error GoTo EH

    '//**** Check if the currently displayed import record has changed
    If bEdit Then
        If MsgBox(" The import record display has been changed. " & vbCr & vbCr & _
           " Save changes? ", vbYesNo + vbInformation) = vbYes Then
            cmdReplace_Click
          Else
            MsgBox "Import record display changed, but not saved"
        End If
    End If

    '//**** Get the total lines in the file
    lTotalLines = GetFileLines(sFileName)

    '//**** Check if there are any lines in the file
    If lTotalLines < 1 Then

        '//**** Since there are no lines the current line is 0
        lCurrentLine = 0

        '//**** Exit Sub/Function before error handler
        Exit Sub

        '//**** End if check if there are any lines in the file
    End If

    '//**** Decrement current line variable lCurrentLine to indicate
    '//     previous line gets displayed in lstLayout
    lCurrentLine = lCurrentLine - 1

    If lTotalLines < lCurrentLine Then
        lCurrentLine = lTotalLines
    End If

    '//****
    sldFileMove.Value = lCurrentLine

    '//****
    'ListLoad

    '//**** Exit Sub/Function before error handler

Exit Sub

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the cmdPrevious_Click"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Error Handling

Exit Sub

    '//****

End Sub

Private Sub cmdLayoutMaint_Click()

  '//**************************************************************************
  '// Open Layout Maintenance task
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Show the Import Utility Main from

    frmLayout_Maint.Show

End Sub

Private Sub cmdExit_Click()

  '//**************************************************************************
  '// cmdClose_Click - Click Event of Cancel Button to end program
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//****

    Unload frmImport_Utility

    '//****
    Set frmImport_Utility = Nothing

    '//****
    End

End Sub

Private Sub sldFileMove_Change()

  '//**************************************************************************
  '// sldFileMove_Change - Event to set current line equal to new slider control value
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

    On Error GoTo EH

    '//**** Get the total lines in the file
    lTotalLines = GetFileLines(sFileName)

    '//**** Check if there are any lines in the file
    If lTotalLines < 1 Then

        '//**** Since there are no lines the current line is 0
        lCurrentLine = 0

        '//**** Exit Sub/Function before error handler
        Exit Sub

        '//**** End if check if there are any lines in the file
    End If

    '//****
    lCurrentLine = sldFileMove.Value

    '//****
    ListLoad

    '//**** Exit Sub/Function before error handler

Exit Sub

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the sldFileMove_Change"

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Error Handling

Exit Sub

    '//****

End Sub

Private Sub cmdSearch_Click()

  '//**************************************************************************
  '// cmdSearch_Click - Common Dialog for searching for import file
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

    On Error GoTo EH

    '//****
    CD.ShowOpen

    '//****
    sFileName = Trim$(CD.FileName)

    '//****
    txtImportFile.Text = sFileName

    '//****
    lTotalLines = GetFileLines(sFileName)

    '//****
    sldFileMove.Max = lTotalLines

    '//**** Exit Sub/Function before error handler

Exit Sub

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the CommonDialog File Search"

    '//**** Error Handling
    Select Case Err.Number
      Case 6
        '//**** Handle VB error 6 Overflow
        '//**** Resume Next

      Case 13

        '//**** Type mismatch occurred
        '//**** Resume Next

      Case 52

        '//**** Handle File error 52 Bad file name or number
        MsgBox sMessage & vbLf & "Verify file name: " & sFileName _
               , vbInformation, "Verify file name"

        '//**** Error Tracking
        CPATracker.Tracker sMessage & "Verify file: " & sFileName, "LogFile.log", False, True

        '//**** Exit Function
        Exit Sub

      Case 53

        '//**** Handle File error 52 Bad file name or number
        MsgBox sMessage & vbLf & "Verify file name: " & sFileName _
               , vbInformation, "Verify file name"

        '//**** Error Tracking
        CPATracker.Tracker sMessage & "Verify file: " & sFileName, "LogFile.log", False, True

        '//**** Exit Function
        Exit Sub

      Case 91

        '//**** Handle VB error 91 Object variable or With block variable not set
        '//**** Resume Next

      Case 440

        '//**** Handle VB error 440 Automation
        '//**** Resume Next

      Case Else

        '//**** Unhandled errors
        '//**** Resume Next

    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function/procedure
    Resume Next

End Sub

Public Function GetFileLines(ByVal strFilePath As String) As Long

  '//**************************************************************************
  '// GetFileLines - function that returns the number of lines from a text file as Integer.
  '//
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '// strFilePath As String is the path\file for which to count total lines
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

    On Error GoTo EH

  '//**** Function Variable Declarations
  Dim iFileNum As Integer
  Dim lLineCount As Long

    '//**** Initialize line count to zero
    lLineCount = 0

    '//**** FreeFile returns the next file number available for use by Open
    iFileNum = FreeFile

    If (CPAFiles.FileExists(strFilePath)) Then

        '//**** Open file
        Open strFilePath For Input As iFileNum

      Else

        '//**** Handle does not exits name or number
        MsgBox "Error occured accessing file " & vbCr & sFileName & vbCr & vbCr & _
               "Verify file name and location  ", vbInformation, " File Error"

        '//**** Maintain slider position
        sldFileMove.Value = lCurrentLine

        txtImportFile.SetFocus

        '//**** Error Tracking
        sMessage = "File doesn't exist error occurred during the GetFileLines function"
        CPATracker.Tracker sMessage & " Verify file: " & sFileName, "LogFile.log", False, True

        '//**** File doesn't exist
        GetFileLines = lLineCount

        Exit Function

    End If

    '//**** loop through file
  Dim strBuffer As String

    Do While Not EOF(iFileNum)

        '//**** read line
        Input #iFileNum, strBuffer

        '//**** update count
        lLineCount = lLineCount + 1

    Loop

    '//**** close file
    Close iFileNum

    '//**** return value
    GetFileLines = lLineCount

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the GetFileLines function"

    '//****
    GetFileLines = lLineCount

    '//**** Error Handling
    Select Case Err.Number

        '//**** Handle VB error 6 Overflow
      Case 6

        '//**** Resume Next

        '//**** Type mismatch occurred
      Case 13

        '//**** Resume Next

        '//**** Handle VB error 91 Object variable or With block variable not set
      Case 91

        '//**** Resume Next

        '//**** Handle File error 52 Bad file name or number
      Case 52

        '//****
        MsgBox "Error occured accessing file " & vbCr & sFileName _
               , vbInformation, "Verify file name"

        '//****
        sldFileMove.Value = lCurrentLine

        '//**** Error Tracking
        CPATracker.Tracker sMessage & " Verify file: " & sFileName, "LogFile.log", False, True

        '//**** Exit Function
        Exit Function

        '//**** Handle File error 52 Bad file name or number
      Case 53

        '//****
        MsgBox "Error occured accessing file " & vbCr & sFileName _
               , vbInformation, "Verify file name"

        '//****
        sldFileMove.Value = lCurrentLine

        '//**** Error Tracking
        CPATracker.Tracker sMessage & " Verify file: " & sFileName, "LogFile.log", False, True

        '//**** Exit Function
        Exit Function

        '//**** Handle File error 75 Path/File access error
      Case 75

        '//****
        MsgBox "Error occured accessing file " & vbCr & sFileName _
               , vbInformation, "Verify file name"

        '//****
        sldFileMove.Value = lCurrentLine

        '//**** Error Tracking
        CPATracker.Tracker sMessage & " Verify file: " & sFileName, "LogFile.log", False, True

        '//****
        txtImportFile.SetFocus

        '//****
        Exit Function

        '//**** Handle VB error 440 Automation
      Case 440

        '//**** Resume Next

        '//**** Unhandled errors
      Case Else

        '//**** Resume Next

    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True
    '//**** Continue function/procedure
    Resume Next

    '//****

End Function

Private Function TextFile(ByRef pStrFileName As String) As String

  '//**********************************************************************
  '// TextFile - Reads and returns a text file as String
  '// Note: If the file does not exist, an error will occur and the function exits
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**********************************************************************

    On Error GoTo EH

  '//**** Function Variable Declarations
  Dim llngFileNum As Long
  Dim llngFileLen As Long

    '//**** Initialize variable for size of file
    llngFileLen = 0

    '//**** Get the size of the file
    llngFileLen = FileLen(pStrFileName)

    '//**** Check if the file is no greater than 0
    If llngFileLen = 0 Then Exit Function

    '//**** Read the file into memory
    llngFileNum = FreeFile

    '//**** Open the file
    Open pStrFileName For Input As #llngFileNum

    '//**** Return TextFile string
    TextFile = Input$(llngFileLen, #llngFileNum)

    '//****
    Close #llngFileNum

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the TextFile function"
    '//**** Error Handling
    Select Case Err.Number

        '//**** Handle VB error 6 Overflow
      Case 6

        '//**** Resume Next

        '//**** Type mismatch occurred
      Case 13

        '//**** Resume Next

        '//**** Handle File error 52 Bad file name or number
      Case 52

        MsgBox sMessage & vbLf & "Verify file name: " & sFileName _
               , vbInformation, "Verify file name"

        '//**** Error Tracking
        CPATracker.Tracker sMessage & "Verify file: " & sFileName, "LogFile.log", False, True

        '//**** Exit Function
        Exit Function

        '//**** Handle File error 52 Bad file name or number
      Case 53

        MsgBox sMessage & vbLf & "Verify file name: " & sFileName _
               , vbInformation, "Verify file name"

        '//**** Error Tracking
        CPATracker.Tracker sMessage & "Verify file: " & sFileName, "LogFile.log", False, True

        '//**** Exit Function
        Exit Function

        '//**** Handle VB error 91 Object variable or With block variable not set
      Case 91

        '//**** Resume Next

        '//**** Handle VB error 440 Automation
      Case 440

        '//**** Resume Next

        '//**** Unhandled errors
      Case Else

        '//**** Resume Next

    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True
    '//**** Continue function/procedure
    Resume Next

End Function

Private Sub lstLayout_Click()

  '//**************************************************************************
  '// lstLayout_Click - Preview and edit delimited import files
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

    On Error GoTo EH

  '//**** Sub Variable Declarations
  Dim sFieldDefinition As String

    '//****
    sFieldDefinition = ""

    '//****
    If lstLayout.ListIndex > -1 Then

        '//****
        sFieldDefinition = Trim$(Mid$(lstLayout.Text, 5, 4))

        '//****
        sFieldDefinition = GetFieldDefinition(sFieldDefinition)

        '//**** Display extended information
        txtFieldInfo.Text = Trim$(sFieldDefinition)

    End If

    '//**** Exit Sub/Function before error handler

Exit Sub

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the lstLayout_Click"
    '//**** Error Handling
    Select Case Err.Number
      Case 6
        '//**** Handle VB error 6 Overflow
        '//**** Resume Next
      Case 13
        '//**** Type mismatch occurred
        '//**** Resume Next
      Case 91
        '//**** Handle VB error 91 Object variable or With block variable not set
        '//**** Resume Next
      Case 440
        '//**** Handle VB error 440 Automation
        '//**** Resume Next
      Case Else
        '//**** Unhandled errors
        '//**** Resume Next
    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True
    '//**** Continue function/procedure
    Resume Next

End Sub

Private Sub DelimeterLoad()

  '//**************************************************************************
  '// DelimeterLoad - Load Delimeters to cboDelimeter Combobox
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//****
  '//**** Load cboDelimter Combobox with Delimeters

    With cboDelimeter
        .AddItem " ; "
        .AddItem " , "
        .AddItem " | "
    End With

    '//**** Set Combobox to first item in list
    If cboDelimeter.ListCount Then
        cboDelimeter.ListIndex = 0
    End If

    '//**** Initialize sDelimeter
    sDelimiter = Trim$(cboDelimeter.Text)

End Sub

Private Sub cboDelimeter_Change()

  '//**************************************************************************
  '// cboDelimeter_Change - Combobox of delimeters available for import file
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Delimeter sDelimiter set by selection in cboDelimeter

    sDelimiter = Trim$(cboDelimeter.Text)

End Sub

Private Sub cboDelimeter_Click()

  '//**************************************************************************
  '// cboDelimeter_Click - Combobox of delimeters available for import file
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Delimeter sDelimiter set by selection in cboDelimeter

    sDelimiter = Trim$(cboDelimeter.Text)

End Sub

Private Sub ImportTypeLoad()

  '//**************************************************************************
  '// ImportTypeLoad - Load Import Layout Types to cboImportType Combobox
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Load cboImportType Combobox with Import Layout Types

    With cboImportType
        .AddItem "AR Posted Invoices"
        .AddItem "AR Pending Invoices"
        .AddItem "AP Posted Vouchers"
        .AddItem "AP Pending Vouchers"
        .AddItem "GL Natural Accounts"
        .AddItem "GL Account Segments"
        .AddItem "GL Accounts"
        .AddItem "GL Transactions"
        .AddItem "IM Items"
        .AddItem "None"
    End With

    '//**** Set Combobox to first item in list
    If cboImportType.ListCount Then
        cboImportType.ListIndex = 0
    End If

    '//**** Initialize sDelimeter
    sImportType = cboImportType.Text

End Sub

Private Sub cboImportType_Click()

  '//**************************************************************************
  '// cboImportType_Click - Click Event of Combobox to select Import Type layout
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Import Layout Type sImportType set by selection in cboImportType

    sImportType = cboImportType.Text

End Sub

Private Sub cboImportType_Change()

  '//**************************************************************************
  '// cboImportType_Change - Change Event of Combobox to select Import Type layout
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Import Layout Type sImportType set by selection in cboImportType

    sImportType = cboImportType.Text

End Sub

Private Sub txtImportFile_Change()

  '//**************************************************************************
  '// txtImportFile_Change -
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//****

    sFileName = txtImportFile.Text

End Sub

Private Sub cmdReset_Click()

  '//**************************************************************************
  '// cmdReset_Click - Click Event of Reset Button to reset form controls
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//****

    Set frmImport_Utility = Nothing

    '//****
    Form_Load

End Sub

Private Sub cmdClose_Click()

  '//**************************************************************************
  '// cmdClose_Click - Click Event of Cancel Button to end program
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//****

    Unload frmImport_Utility

    '//****
    Set frmImport_Utility = Nothing

    '//****
    End

End Sub

Private Sub Form_Unload(Cancel As Integer)

  '//**************************************************************************
  '// Form_UnLoad Event
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//****

    Set frmImport_Utility = Nothing

End Sub

Private Sub cmdReplace_Click()

  '//**************************************************************************
  '// cmdReplace_Click -
  '//
  '// Created:  03/30/2002 John C. Kirwin (JCK)
  '// Modified: 03/30/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//****

    With CPAFiles

        '//****
        .OpenFile .GetFilePath(sFileName), .GetFileNameFromPath(sFileName)

        '//****
        .WriteData ReplaceData, txtPreview.Text, lCurrentLine

        '//****
    End With

End Sub

Private Sub cmdInsert_Click()

  '//**************************************************************************
  '// cmdInsert_Click -
  '//
  '// Created:  03/30/2002 John C. Kirwin (JCK)
  '// Modified: 03/30/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//****

    With CPAFiles

        '//****
        .OpenFile .GetFilePath(sFileName), .GetFileNameFromPath(sFileName)

        '//****
        .WriteData InsertData, txtPreview.Text, lCurrentLine

        '//****
    End With

End Sub

Private Sub cmdDelete_Click()

  '//**************************************************************************
  '// cmdDelete_Click -
  '//
  '// Created:  03/30/2002 John C. Kirwin (JCK)
  '// Modified: 03/30/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  
    '//****
    With CPAFiles

        '//****
        .OpenFile App.Path, .GetFileNameFromPath(sFileName)

        '//****
        .DeleteRow (lCurrentLine)

    '//****
    End With

    

End Sub
Private Sub txtPreview_Change()

  '//**************************************************************************
  '// txtPreview_Change - if changes the enable buttons to manipulate file
  '// Created:  03/30/2002 John C. Kirwin (JCK)
  '// Modified: 03/30/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

    '//**** Check if the line is valid and contains the delimiter
    If (InStr(txtPreview.Text, sDelimiter)) > 0 Then

        '//**** Allow modification by enabling button(s)
        cmdReplace.Enabled = True
        'cmdReplace.Insert = True
        'cmdReplace.Delete = True
    
    '//****
    End If

End Sub

Private Sub txtPreview_GotFocus()
  
  '//**************************************************************************
  '// txtPreview_GotFocus -
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     Cancel As Boolean
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************
  
    '//****
    bEdit = True

End Sub

Private Sub txtPreview_Validate(Cancel As Boolean)

  '//**************************************************************************
  '// txtPreview_Validate - To save or not to save that is the question?
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     Cancel As Boolean
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

    '//****
    Debug.Print "txtPreview_Validate"

End Sub

Private Sub Form_Load()

  '//**************************************************************************
  '// Form_Load - prepares frmImport_Utility by initializing form objects and controls...
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

    '//****
    On Error GoTo EH
     
     '//**** Center Form
     Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2


    '//****
    Set CPAFiles = New cCPA_Files
    Set CPAINI = New cCPA_INI
    Set CPAStrings = New cCPA_Strings
    Set CPATracker = New cCPA_Tracker
    Set CPAErrorHandler = New cCPA_ErrorHandler
    Set CPAParseDelimited = New cCPA_ParseDelimited
    Set CPALayouts = New cCPA_Layouts

    '//**** Display instructions
    txtFieldInfo.Text = "Indicate import type, file name, delimeter, number of fields to display, and click the forward button(green arrow)."

    '//**** Clear Field Display List box
    lstLayout.Clear
    cboImportType.Clear
    cboDelimeter.Clear

    '//**** Clear Import File Display Text box
    txtPreview.Text = ""

    '//**** Clear File
    txtImportFile.Text = ""

    '//**** Turn off controls until data is available
    cmdReplace.Enabled = False
    cmdInsert.Enabled = False
    cmdDelete.Enabled = False

    '//**** Initialize the counters to zero
    lListMax = 0

    '//**** Clear Status Bar Variables
    StatusBarClearVars

    '//**** Clear the contents of the status bar
    StatusBarClear

    '//**** Load the Import Type Combobox into the cbo Combobox
    ImportTypeLoad

    '//**** Load the Import Type Combobox into the cboDelimeter Combobox
    DelimeterLoad

    '//**** Initialize current line variable to the 1st line of the file
    lCurrentLine = 0

    '//****
    sMessage = "Note:  The current resolution is: " & _
               Screen.Width / Screen.TwipsPerPixelX & _
               "x" & Screen.Height / Screen.TwipsPerPixelY

    '//**** Send Note to CPAErrorHandler
    CPAErrorHandler.ErrorHandler "frmImport_Utility", _
                                  "Form_Load", 0, sMessage, _
                                  "LogFile.log", False, False

    'SetWindowRgn hWnd, CreateEllipticRgn(-121, -130, 765, 765), True
    SetWindowRgn hWnd, CreateEllipticRgn(-123, -131, 765, 770), True

    '//**** Exit Sub/Function before error handler

Exit Sub

EH:

    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the Form_Load"
    '//**** Error Handling
    Select Case Err.Number

        '//**** Handle VB error 6 Overflow
      Case 6

        '//**** Resume Next

        '//**** Type mismatch occurred
      Case 13

        '//**** Resume Next

        '//**** Handle VB error 91 Object variable or With block variable not set
      Case 91

        '//**** Resume Next

        '//**** Handle VB error 440 Automation
      Case 440

        '//**** Resume Next

        '//**** Unhandled errors
      Case Else

        '//**** Resume Next

        '//****
    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function/procedure
    Resume Next

    '//****

End Sub

Private Function AddList(sRecord As String, iField As Integer, sFieldName As String, _
                         bRequired As Boolean, sFieldDesc As String, sField As String) As Boolean

  '//**************************************************************************
  '// AddList - Line added to list box display Boolean function
  '//
  '// sRecord As String is the type of record
  '// iField As Integer is the location or position of the field in the line
  '// sFieldName As String as the name of field/column of coresponding layout
  '// sField As String
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

    On Error GoTo EH

    '//**** Increment Field counter variable lFieldCount
    lFieldCount = CLng(lFieldCount + 1)

    '//**** Verify iField formatted as double digit
    If iField < 10 Then

        '//**** Single digit iFields need to be Formatted (ie. 9 should be 09)
        sRecord = sRecord & "00" & iField

      ElseIf iField > 9 And iField < 100 Then

        '//**** Format double digits i.e.
        sRecord = sRecord & "0" & iField

      Else
        '//**** Format of triple digits i.e.
        sRecord = sRecord & iField
    End If

    '//**** Do While...Loop through string field description sFieldName
    '//     to repeatedly add a space to the description while it is
    '//     shorter than 25 characters
    Do While Len(sFieldName) < 25

        '//**** Add a space to the end of string sFieldName
        sFieldName = sFieldName & " "

        '//**** Continue Do While...Loop adding spaces to the string field description sFieldName
    Loop

    '//**** Fill List Box
    With lstLayout

        '//**** Add formated item to list box that contains the
        '//     Checkbox, Field, Record, Field Description and Value
        .AddItem CPAStrings.FillAlign(sRecord, 8, " ", True) & "   " _
                 & sFieldName & vbTab & " = " & sField

        '//**** Mark the Checkbox to indicate required fields
        '//     as indicated by the boolean variable bRequired
        If bRequired Then

            '//**** Checkbox gets selected
            lstLayout.Selected(.NewIndex) = True                             ' 0 = Default 1 = True

        End If

    End With

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the AddList Function"

    '//**** Error Handling
    Select Case Err.Number

        '//**** Overflow occurred
      Case 6

        '//**** Resume Next

        '//**** Type mismatch occurred
      Case 13

        '//**** Resume Next

        '//**** Handle VB error 91 Object variable or With block variable not set
      Case 91

        '//**** Resume Next

        '//**** Handle VB error 381 Invalid property array index
      Case 381
        '//**** The index value of the List property for a ListBox
        '//     must be from 0 to 32,766.  Change the index value
        '//     of the property array to a valid setting.

        '//**** Error Tracking
        CPATracker.Tracker sMessage & " (lFieldCount: " & lFieldCount & "), (iField: " _
                            & iField & ")", "LogFile.log", True, True

        '//****
        Unload frmImport_Utility

        '//****
        End

        '//**** Handle VB error 440 Automation
      Case 440

        '//**** Resume Next

        '//**** Unhandled errors
      Case Else

        '//**** Resume Next

        '//****
    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True
    '//**** Continue function/procedure
    Resume Next

    '//****

End Function

Private Function GetFieldDefinition(sField As String) As String

  '//**************************************************************************
  '// GetFieldDefinition - Function to build field definition String
  '// sField As String
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

  '//**** Function Variable Declarations
  Dim sDescription As String

    '//****
    sDescription = ""

    '//****
    Select Case cboImportType.Text

        
      '//****
      Case "GL Natural Accounts"
        
        '//**** Evaluate if sSection = "GLNRL" GL Natural Accounts Record Layout
        If sSection = "GLNRL" Then

            '//****
            sDescription = aGLNRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aGLNRL(lstLayout.ListIndex + 1, 2)
          '//****
          Else

            '//****
            sDescription = "No Import file definition information available at this time"

        '//****
        End If

      Case "GL Account Segments"
        '//**** Evaluate if sSection = "GLSRL" GL Account Segments Record Layout
        If sSection = "GLSRL" Then

            '//****
            sDescription = aGLSRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aGLSRL(lstLayout.ListIndex + 1, 2)

          '//****
          Else

            '//****
            sDescription = "No Import file definition information available at this time"

        '//****
        End If

      Case "GL Accounts"
        '//**** Evaluate if sSection = "GLARL" GL Account Record Layout
        If sSection = "GLARL" Then

            '//****
            sDescription = aGLARL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aGLARL(lstLayout.ListIndex + 1, 2)

          '//****
          Else

            '//****
            sDescription = "No Import file definition information available at this time"

        '//****
        End If

      Case "GL Transactions"
        '//**** Evaluate if sSection = "GLBRL" GL Transaction Batch Record Layout
        If sSection = "GLBRL" Then

            '//****
            sDescription = aGLBRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aGLBRL(lstLayout.ListIndex + 1, 2)

          '//****
          ElseIf sSection = "GLTRL" Then

            '//****
            sDescription = aGLTRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aGLTRL(lstLayout.ListIndex + 1, 2)

          '//****
          ElseIf sSection = "GLFRL" Then

            '//****
            sDescription = aGLFRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aGLFRL(lstLayout.ListIndex + 1, 2)

          '//****
          Else

            '//****
            sDescription = "No Import file definition information available at this time"

        End If

        '//****
      Case "AP Posted Vouchers"

        '//****
        'sDescription = GetPostedVoucher(sSection)

        '//**** Evaluate if sSection = "VPBRL" Posted Voucher Batch Record Layout array
        If sSection = "VPBRL" Then

            '//****
            sDescription = aVPBRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aVPBRL(lstLayout.ListIndex + 1, 2)

            '//**** Evaluate if sSection = "VPHRL" Posted Voucher Header Record Layout array
          ElseIf sSection = "VPHRL" Then

            '//****
            sDescription = aVPHRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aVPHRL(lstLayout.ListIndex + 1, 2)

            '//**** Evaluate if sSection = "VPDRL" Posted Voucher Detail Record Layout array
          ElseIf sSection = "VPDRL" Then

            '//****
            sDescription = aVPDRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aVPDRL(lstLayout.ListIndex + 1, 2)

            '//**** Evaluate if sSection = "VPARL" Posted Voucher Application Record Layout array
          ElseIf sSection = "VPARL" Then

            '//****
            sDescription = aVPARL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aVPARL(lstLayout.ListIndex + 1, 2)

            '//**** Evaluate if sSection = "VPXRL" Posted Voucher Tax Header Record Layout array
          ElseIf sSection = "VPXRL" Then

            '//****
            sDescription = aVPXRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aVPXRL(lstLayout.ListIndex + 1, 2)

            '//**** Evaluate if sSection = "VPTRL" Posted Voucher Tax Detail Record Layout array
          ElseIf sSection = "VPTRL" Then

            '//****
            sDescription = aVPTRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aVPTRL(lstLayout.ListIndex + 1, 2)

          Else

            '//****
            sDescription = "No Import file definition information available at this time"

        '//****
        End If

      '//****
      Case "AP Pending Vouchers"

        '//**** Evaluate if sSection = "PVBRL" Pending Voucher Batch Record Layout
        If sSection = "PVBRL" Then

            '//****
            sDescription = aPVBRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aPVBRL(lstLayout.ListIndex + 1, 2)

          '//**** Evaluate if sSection = "PVHRL" Pending Voucher Header Record Layout
          ElseIf sSection = "PVHRL" Then

            '//****
            sDescription = aPVHRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aPVHRL(lstLayout.ListIndex + 1, 2)

          '//**** Evaluate if sSection = "PVDRL" Pending Voucher Detail Record Layout
          ElseIf sSection = "PVDRL" Then

            '//****
            sDescription = aPVDRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aPVDRL(lstLayout.ListIndex + 1, 2)

          '//****
          Else

            '//****
            sDescription = "No Import file definition information available at this time"

        '//****
        End If

      '//****
      Case "AR Posted Invoices"

        '//****
        'sDescription = GetPostedInvoice(sSection)

        '//**** Evaluate if sSection = "RIBRL" Posted Invoice Batch Record Layout array (17 Fields)
        If sSection = "RIBRL" Then

            '//****
            sDescription = aRIBRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aRIBRL(lstLayout.ListIndex + 1, 2)

            '//**** Evaluate if sSection = "RIRRL" Posted Invoice Header Record Layout array (84 Fields)
          ElseIf sSection = "RIRRL" Then

            '//****
            sDescription = aRIRRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aRIRRL(lstLayout.ListIndex + 1, 2)

            '//**** Evaluate if sSection = "RIDRL" Posted Invoice Detail Record Layout array (44 Fields)
          ElseIf sSection = "RIDRL" Then

            '//****
            sDescription = aRIDRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aRIDRL(lstLayout.ListIndex + 1, 2)

            '//**** Evaluate if sSection = "RIARL" Posted Invoice Application Record Layout array
          ElseIf sSection = "RIARL" Then

            '//****
            sDescription = aRIARL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aRIARL(lstLayout.ListIndex + 1, 2)

          '//**** Evaluate if sSection = "RIERL" Posted Invoice Tax Header Record Layout array
          ElseIf sSection = "RIERL" Then

            '//****
            sDescription = aRIERL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aRIERL(lstLayout.ListIndex + 1, 2)

          '//**** Evaluate if sSection = "RITRL" Posted Invoice Tax Detail Record Layout array
          ElseIf sSection = "RITRL" Then

            '//****
            sDescription = aRITRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aRITRL(lstLayout.ListIndex + 1, 2)

          '//****
          Else

            '//****
            sDescription = "No Import file definition information available at this time"

        '//****
        End If

      '//****
      Case "AR Pending Invoices"

        '//****
        'sDescription = GetPendingInvoice(sSection)

        '//**** Evaluate if sSection = "PIBRL" Pending Invoice Batch Record Layout array (17 Fields)
        If sSection = "PIBRL" Then

            '//****
            sDescription = aPIBRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aPIBRL(lstLayout.ListIndex + 1, 2)

          '//**** Evaluate if sSection = "PIHRL" Pending Invoice Header Record Layout array (84 Fields)
          ElseIf sSection = "PIHRL" Then

            '//****
            sDescription = aPIHRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aPIHRL(lstLayout.ListIndex + 1, 2)

          '//**** Evaluate if sSection = "PIDRL" Pending Invoice Detail Record Layout array (44 Fields)
          ElseIf sSection = "PIDRL" Then

            '//****
            sDescription = aPIDRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aPIDRL(lstLayout.ListIndex + 1, 2)

          '//****
          Else

            '//****
            sDescription = "No Import file definition information available at this time"

        '//****
        End If

      '//****
      Case "IM Items"

        '//**** Evaluate if sSection = "IMIRL" Inventory Items Record Layout array (61 Fields)
        If sSection = "IMIRL" Then

            '//****
            sDescription = aIMIRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aIMIRL(lstLayout.ListIndex + 1, 2)

          '//**** Evaluate if sSection = "IMURL" Inventory Management Item Unit of Measure Record Layout array (8 Fields)
          ElseIf sSection = "IMURL" Then

            '//****
            sDescription = aIMURL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aIMURL(lstLayout.ListIndex + 1, 2)

          '//**** Evaluate if sSection = "IMLRL" Landed Cost Factor Record Layout array (3 Fields)
          ElseIf sSection = "IMLRL" Then

            '//****
            sDescription = aIMLRL(lstLayout.ListIndex + 1, 1) & " - " & _
                           aIMLRL(lstLayout.ListIndex + 1, 2)

          '//****
          Else

            '//****
            sDescription = "No Import file definition information available at this time"

        '//****
        End If

        '//****
      Case Else

        '//****
        sDescription = "No Import file definition information available at this time"

    '//****
    End Select


    '//****
    GetFieldDefinition = sDescription

    '//**** Exit Sub/Function before error handler

Exit Function

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the GetFieldDefinition Function"

    '//**** Error Handling
    Select Case Err.Number
      Case 6

        '//**** Overflow occurred
        '//**** Resume Next

      Case 13

        '//**** Type mismatch occurred
        '//**** Resume Next

      Case 91

        '//**** Handle VB error 91 Object variable or With block variable not set
        '//**** Resume Next

      Case 440

        '//**** Handle VB error 440 Automation
        '//**** Resume Next

      Case Else

        '//**** Unhandled errors
        '//**** Resume Next

    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True
    '//**** Continue function/procedure
    Resume Next

    '//****

End Function

Private Sub ListLoad()

  '//**************************************************************************
  '// ListLoad - Loads delimited file into lstLayout list box
  '//
  '// Created:  12/26/2002 John C. Kirwin (JCK)
  '// Modified: 03/03/2002 JCK - Formated & Commented
  '//
  '// Parameters:
  '//     None
  '//
  '// Returns:
  '//     None
  '//
  '//**************************************************************************

    On Error GoTo EH

  '//**** Sub Variable Declarations
  Dim sDataLine As String
  Dim iField As Integer
  Dim sField As String
  Dim sFieldName As String
  Dim sFieldDescription As String
  Dim sRecord As String
  Dim lDelimiter As Long
  Dim lArrayMemberCount As Long
  Dim lngI As Long
  Dim lngFile As Long
  Dim bReqField As Boolean
  Dim bAddList As Boolean
  Dim bGetDefinition As Boolean

    '//**** Redimension array completely erasing the array's contents
    ReDim vData(1) As Variant

    '//**** Clear Display and Status Bar
    lstLayout.Clear
    txtPreview.Text = ""
    StatusBarClear

    '//**** Reset Data Line and Field Count variables
    lFieldCount = 0
    lDataLine = 0

    '//**** Define variables
    lDelimiter = Asc(sDelimiter)
    sFileName = txtImportFile.Text
    sImportType = cboImportType.Text

    '//**** Determine Import Type and Prepare Layout array
    If sImportType = "GL Natural Accounts" Then
        
        '//**** GL Natural Accounts
        bGetDefinition = CPALayouts.GLNaturalAccts(aGLNRL(), iRows, iColumns)
      
      '//**** GL Account Segments
      ElseIf sImportType = "GL Account Segments" Then
        
        '//**** GL Account Segments
        bGetDefinition = CPALayouts.GLSegments(aGLSRL(), iRows, iColumns)
      
      '//**** GL Accounts
      ElseIf sImportType = "GL Accounts" Then
        
        '//**** GL Accounts
        bGetDefinition = CPALayouts.GLAccounts(aGLARL(), iRows, iColumns)
      
      '//**** GL Transactions
      ElseIf sImportType = "GL Transactions" Then
        
        '//**** GL Transactions
        bGetDefinition = CPALayouts.GLTransactions(aGLBRL(), aGLTRL(), aGLFRL(), iRows, iColumns)
          
      '//**** AP Pending Vouchers
      ElseIf sImportType = "AP Pending Vouchers" Then
        
        '//**** AP Pending Vouchers
        bGetDefinition = CPALayouts.PendingVouchers(aPVBRL(), aPVHRL(), aPVDRL(), iRows, iColumns)
          
      '//**** AP Posted Vouchers
      ElseIf sImportType = "AP Posted Vouchers" Then
        
        '//**** AP Posted Vouchers
        bGetDefinition = CPALayouts.PostedVouchers(aVPBRL(), aVPHRL(), aVPDRL(), aVPARL(), _
                                                   aVPXRL(), aVPTRL(), iRows, iColumns)
      
      '//**** AR Posted Invoices
      ElseIf sImportType = "AR Posted Invoices" Then
        
        '//**** AR Posted Invoices
        bGetDefinition = CPALayouts.PostedInvoices(aRIBRL(), aRIRRL(), aRIDRL(), aRIARL(), _
                                                    aRIERL(), aRITRL(), iRows, iColumns)
      
      '//**** AR Pending Invoices
      ElseIf sImportType = "AR Pending Invoices" Then
        
        '//**** AR Pending Invoices
        bGetDefinition = CPALayouts.PendingInvoices(aPIBRL(), aPIHRL(), aPIDRL(), iRows, iColumns)
      
      '//**** IM Items
      ElseIf sImportType = "IM Items" Then
        
        '//**** IM Items
        bGetDefinition = CPALayouts.IMItems(aIMIRL(), aIMURL(), aIMLRL(), iRows, iColumns)
      
      '//**** Unknown
      Else
        
        '//**** Unknown
        sImportType = "Unknown"
    
    '//****
    End If

    '//****
    lngFile = FreeFile

    '//**** Open file
    Open sFileName For Input As #lngFile

    '//**** Loop through each line of the file
    Do While Not EOF(lngFile) And lDataLine <= lCurrentLine

        '//**** Line Input
        Line Input #lngFile, sDataLine

        '//**** Increment Record Count lDataLine
        lDataLine = lDataLine + 1

        '//**** Determine if the field is to be displayed
        If lDataLine = lCurrentLine Then

            '//**** Add line to existing line(s) in txtPreview
            txtPreview.Text = sDataLine

            '//**** Parse the data as members of the array
            '//     and return the total number of array members
            lArrayMemberCount = CPAParseDelimited.lStringToArray(vData(), sDataLine, lDelimiter)

            '//**** Check if any
            If lArrayMemberCount = 1 Then

                '//**** Validate delimeter
                MsgBox "Verify character selected as field delimiter is valid", vbInformation, " Warning"

                '//**** Clear Display and Status Bar
                lstLayout.Clear
                StatusBarClear

                '//**** Initialize the counters to zero
                lFieldCount = 0
                lListMax = 0

                '//**** Set form focus to cboDelimeter ComboBox
                cboDelimeter.SetFocus
                Exit Sub

            End If

            '//**** Display the array members and/or pass the data
            For lngI = LBound(vData) + 1 To UBound(vData) - 1

                '//**** Reset required field boolean bRequired
                bReqField = False

                '//**** Parse out 1st field as sRecord located at
                '//     the beginning of the data line sDataLine
                sRecord = CPAParseDelimited.ParseString(sDataLine, 1, sDelimiter)

                '//**** Set array members to field integer variable iField
                iField = Int(lngI)

                '//**** Parse out field located at field iField
                '//     from the data line sDataLine
                sField = CPAParseDelimited.ParseString(sDataLine, iField, sDelimiter)

                '//**** Determine Import Layout Type
                If sImportType = "GL Natural Accounts" Then

                    '//****
                    If iField <= 10 Then
                        '//**** Hard code Import Type since its not in import
                        sRecord = "N"

                        '//**** Lookup Field Name
                        sFieldName = aGLNRL(iField, 1)

                        '//**** Lookup Field Description
                        sFieldDescription = aGLNRL(iField, 2)

                        '//**** Set required field boolean bRequired to True
                        '//     Use CBool to convert zero to False is returned
                        '//     otherwise, True is returned
                        bReqField = CBool(aGLNRL(iField, 3))

                        '//**** Increment BatchCounter
                        lBatchCount = lBatchCount + 1

                        '//****
                        sSection = "GLNRL"

                        '//****
                        sRecordType = "Natural Accounts Record"

                        '//**** Parse out 8th field of the sDataLine as sCompany
                        sCompany = CPAParseDelimited.ParseString(sDataLine, 8, sDelimiter)

                      '//****
                      Else

                        '//**** Unknown Field Description
                        sFieldName = "                   "

                        '//**** Lookup Field Description
                        sFieldDescription = "                   "

                        '//**** Assume undefined field is not required
                        bReqField = False

                        '//**** Increment Error Count
                        lErrorCount = lErrorCount + 1

                    '//****
                    End If

                    '//**** Add Record to list
                    bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                  '//****
                  ElseIf sImportType = "GL Account Segments" Then

                    '//****
                    If iField <= 7 Then
                        '//**** Hard code Import Type since its not in import
                        sRecord = "S"

                        '//**** Lookup Field Name
                        sFieldName = aGLSRL(iField, 1)

                        '//**** Lookup Field Description
                        sFieldDescription = aGLSRL(iField, 2)

                        '//**** Set required field boolean bRequired to True
                        '//     Use CBool to convert zero to False is returned
                        '//     otherwise, True is returned
                        bReqField = CBool(aGLSRL(iField, 3))

                        '//**** Increment BatchCounter
                        lBatchCount = lBatchCount + 1

                        '//****
                        sSection = "GLSRL"

                        '//****
                        sRecordType = "Account Segments Record"

                        '//**** Parse out 8th field of the sDataLine as sCompany
                        sCompany = CPAParseDelimited.ParseString(sDataLine, 8, sDelimiter)

                      '//****
                      Else

                        '//**** Unknown Field Description
                        sFieldName = "                   "

                        '//**** Lookup Field Description
                        sFieldDescription = "                   "

                        '//**** Assume undefined field is not required
                        bReqField = False

                        '//**** Increment Error Count
                        lErrorCount = lErrorCount + 1

                    '//****
                    End If

                    '//**** Add Record to list
                    bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                  '//****
                  ElseIf sImportType = "GL Accounts" Then

                    '//****
                    If iField <= 8 Then
                        '//**** Hard code Import Type since its not in import
                        sRecord = "A"

                        '//**** Lookup Field Name
                        sFieldName = aGLARL(iField, 1)

                        '//**** Lookup Field Description
                        sFieldDescription = aGLARL(iField, 2)

                        '//**** Set required field boolean bRequired to True
                        '//     Use CBool to convert zero to False is returned
                        '//     otherwise, True is returned
                        bReqField = CBool(aGLARL(iField, 3))

                        '//**** Increment BatchCounter
                        lBatchCount = lBatchCount + 1

                        '//****
                        sSection = "GLARL"

                        '//****
                        sRecordType = "Accounts Record"

                        '//**** Parse out 8th field of the sDataLine as sCompany
                        sCompany = CPAParseDelimited.ParseString(sDataLine, 8, sDelimiter)

                      '//****
                      Else

                        '//**** Unknown Field Description
                        sFieldName = "                   "

                        '//**** Lookup Field Description
                        sFieldDescription = "                   "

                        '//**** Assume undefined field is not required
                        bReqField = False

                        '//**** Increment Error Count
                        lErrorCount = lErrorCount + 1

                    '//****
                    End If

                    '//**** Add Record to list
                    bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                  '//**** Determine Import Layout Type
                  ElseIf sImportType = "GL Transactions" Then
                    '//****
                    If sRecord = """@HDR""" Then
                        '//****
                        If CPAParseDelimited.ParseString(sDataLine, 2, sDelimiter) = """BATCH""" Then
                            '//****
                            sRecord = "B"
                          '//****
                          ElseIf CPAParseDelimited.ParseString(sDataLine, 2, sDelimiter) = """JOURNAL""" Then
                            '//****
                            sRecord = "D"
                          '//****
                          Else
                            '//**** Nothing for now
                        End If

                      '//****
                      ElseIf sRecord = """@END""" Then
                        '//****
                        If CPAParseDelimited.ParseString(sDataLine, 2, sDelimiter) = """JOURNAL""" Then
                            '//****
                            sRecord = "E"
                          '//****
                          ElseIf CPAParseDelimited.ParseString(sDataLine, 2, sDelimiter) = """BATCH""" Then
                            '//****
                            sRecord = "F"
                          '//****
                          Else
                        '//**** Nothing for now
                        End If

                      Else
                        sRecord = "T"
                    End If

                    '//**** Compare the Record Type ID to "B"
                    If sRecord = "B" Then

                        '//**** Pending Invoices Record Layout
                        If iField <= 4 Then

                            '//**** Lookup Field Name
                            sFieldName = aGLBRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aGLBRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aGLBRL(iField, 3))

                            '//****
                            sSection = "GLBRL"

                            '//****
                            sRecordType = "Batch Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Lookup Field Description
                            sFieldDescription = "                   "

                            '//**** Assume undefined field is not required
                            bReqField = False

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "V"
                      ElseIf sRecord = "T" Then

                        '//**** Pending Voucher Header Record Layouts
                        If iField <= 19 Then

                            '//**** Lookup Field Name
                            sFieldName = aGLTRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aGLTRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aGLTRL(iField, 3))

                            '//****
                            sSection = "GLTRL"

                            '//****
                            sRecordType = "Transactions Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID
                      ElseIf sRecord = "F" Then

                        '//**** Pending Invoice Detail Record Layout
                        If iField <= 2 Then

                            '//**** Lookup Field Name
                            sFieldName = aGLFRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aGLFRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aGLFRL(iField, 3))

                            '//****
                            sSection = "GLFRL"

                            '//****
                            sRecordType = "Footer Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "I"
                        '//     This means Pending Voucher Import Type indicated
                        '//     but Pending Invoice import file indicated
                      Else
                        If sRecord = "I" Then

                            '//**** Error Tracking
                            CPATracker.Tracker "Verify " & sImportType & " is the " & vbCr _
                                                & "correct Import Type for file: " & vbCr _
                                                & sFileName, "LogFile.log", True, True

                        End If

                        '//**** Unknown Field Description
                        sFieldName = "                   "

                        '//**** Increment Error Count
                        lErrorCount = lErrorCount + 1

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                    End If

                    '//**** Compare the Import Type string
                  ElseIf sImportType = "AP Posted Vouchers" Then

                    '//**** Compare the Record Type ID to "B"
                    If sRecord = "B" Then

                        '//**** Posted Vouchers Record Layout
                        If iField <= 8 Then

                            '//**** Lookup Field Name
                            sFieldName = aVPBRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aVPBRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aVPBRL(iField, 3))

                            '//**** Check if this is a new batch
                            If iField = 1 Then

                                '//**** Increment BatchCounter
                                lBatchCount = lBatchCount + 1

                            End If

                            '//****
                            sSection = "VPBRL"

                            '//****
                            sRecordType = "Batch Record"

                            '//**** Parse out 8th field of the sDataLine as sCompany
                            sCompany = CPAParseDelimited.ParseString(sDataLine, 8, sDelimiter)

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Lookup Field Description
                            sFieldDescription = "                   "

                            '//**** Assume undefined field is not required
                            bReqField = False

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "V"
                      ElseIf sRecord = "V" Then

                        '//**** Posted Voucher Header Record Layouts
                        If iField <= 39 Then

                            '//**** Lookup Field Name
                            sFieldName = aVPHRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aVPHRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aVPHRL(iField, 3))

                            '//****
                            sSection = "VPHRL"

                            '//****
                            sRecordType = "Header Record"

                            '//**** Parse out 8th field of the sDataLine as sCompany
                            sCompany = CPAParseDelimited.ParseString(sDataLine, 8, sDelimiter)

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "D"
                      ElseIf sRecord = "D" Then

                        '//**** Posted Voucher Detail Record Layout
                        If iField <= 25 Then

                            '//**** Lookup Field Name
                            sFieldName = aVPDRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aVPDRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aVPDRL(iField, 3))

                            '//****
                            sSection = "VPDRL"

                            '//****
                            sRecordType = "Detail Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "A" Posted Voucher Applicaton Record Layout
                      ElseIf sRecord = "A" Then

                        '//**** Posted Voucher Applicaton Record Layout
                        If iField <= 10 Then

                            '//**** Lookup Field Name
                            sFieldName = aVPARL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aVPARL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aVPARL(iField, 3))

                            '//****
                            sSection = "VPARL"

                            '//****
                            sRecordType = "Application Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "X" Posted Voucher Tax Header Record Layout
                      ElseIf sRecord = "X" Then

                        '//**** Posted Voucher Tax Header Record Layout
                        If iField <= 15 Then

                            '//**** Lookup Field Name
                            sFieldName = aVPXRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aVPXRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aVPXRL(iField, 3))

                            '//****
                            sSection = "VPXRL"

                            '//****
                            sRecordType = "Tax Header Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "T" Posted Voucher Tax Detail Record Layout
                      ElseIf sRecord = "T" Then

                        '//**** Posted Voucher Tax Detail Record Layout
                        If iField <= 16 Then

                            '//**** Lookup Field Name
                            sFieldName = aVPTRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aVPTRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aVPTRL(iField, 3))

                            '//****
                            sSection = "VPTRL"

                            '//****
                            sRecordType = "Tax Details Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Else somthing is wrong like
                        '//     Posted Voucher Import Type
                        '//     indicated but Posted Voucher
                        '//     import file indicated
                      Else

                        If sRecord = "I" And iField = 1 Then
                            '//**** Error Tracking
                            CPATracker.Tracker "Verify " & sImportType & " is the " & vbCr _
                                                & "correct Import Type for file: " & vbCr _
                                                & sFileName, "LogFile.log", True, True

                        End If

                        '//**** Unknown Field Description
                        sFieldName = "                   "

                        '//**** Increment Error Count
                        lErrorCount = lErrorCount + 1

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                    End If

                    '//**** Determine Import Layout Type
                  ElseIf sImportType = "AP Pending Vouchers" Then

                    '//**** Compare the Record Type ID to "B"
                    If sRecord = "B" Then

                        '//**** Pending Invoices Record Layout
                        If iField <= 17 Then

                            '//**** Lookup Field Name
                            sFieldName = aPVBRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aPVBRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aPVBRL(iField, 3))

                            '//**** Check if this is a new batch
                            If iField = 1 Then

                                '//**** Increment BatchCounter
                                lBatchCount = lBatchCount + 1

                            End If

                            '//****
                            sSection = "PVBRL"

                            '//****
                            sRecordType = "Batch Record"

                            '//**** Parse out 11th field of the sDataLine as sCompany
                            sCompany = CPAParseDelimited.ParseString(sDataLine, 11, sDelimiter)

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Lookup Field Description
                            sFieldDescription = "                   "

                            '//**** Assume undefined field is not required
                            bReqField = False

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "V"
                      ElseIf sRecord = "V" Then

                        '//**** Pending Voucher Header Record Layouts
                        If iField <= 75 Then

                            '//**** Lookup Field Name
                            sFieldName = aPVHRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aPVHRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aPVHRL(iField, 3))

                            '//****
                            sSection = "PVHRL"

                            '//****
                            sRecordType = "Header Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "D"
                      ElseIf sRecord = "D" Then

                        '//**** Pending Invoice Detail Record Layout
                        If iField <= 38 Then

                            '//**** Lookup Field Name
                            sFieldName = aPVDRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aPVDRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aPVDRL(iField, 3))

                            '//****
                            sSection = "PVDRL"

                            '//****
                            sRecordType = "Detail Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "I"
                        '//     This means Pending Voucher Import Type indicated
                        '//     but Pending Invoice import file indicated
                      Else
                        If sRecord = "I" Then

                            '//**** Error Tracking
                            CPATracker.Tracker "Verify " & sImportType & " is the " & vbCr _
                                                & "correct Import Type for file: " & vbCr _
                                                & sFileName, "LogFile.log", True, True

                        End If

                        '//**** Unknown Field Description
                        sFieldName = "                   "

                        '//**** Increment Error Count
                        lErrorCount = lErrorCount + 1

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                    End If

                    '//**** Compare the Import Type string sImportType
                  ElseIf sImportType = "AR Posted Invoices" Then

                    '//**** Compare the Record Type ID to "B"
                    If sRecord = "B" Then

                        '//**** Posted Invoices Record Layout
                        If iField <= 9 Then

                            '//**** Lookup Field Name
                            sFieldName = aRIBRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aRIBRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aRIBRL(iField, 3))

                            '//**** Check if this is a new batch
                            If iField = 1 Then

                                '//**** Increment BatchCounter
                                lBatchCount = lBatchCount + 1

                            End If

                            '//****
                            sSection = "RIBRL"

                            '//****
                            sRecordType = "Batch Record"

                            '//**** Parse out 8th field of the sDataLine as sCompany
                            sCompany = CPAParseDelimited.ParseString(sDataLine, 8, sDelimiter)

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Lookup Field Description
                            sFieldDescription = "                   "

                            '//**** Assume undefined field is not required
                            bReqField = False

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "I"
                      ElseIf sRecord = "I" Then

                        '//**** Posted Voucher Header Record Layouts
                        If iField <= 40 Then

                            '//**** Lookup Field Name
                            sFieldName = aRIRRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aRIRRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aRIRRL(iField, 3))

                            '//****
                            sSection = "RIRRL"

                            '//****
                            sRecordType = "Header Record"

                            '//**** Parse out 8th field of the sDataLine as sCompany
                            sCompany = CPAParseDelimited.ParseString(sDataLine, 8, sDelimiter)

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "D"
                      ElseIf sRecord = "D" Then

                        '//**** Posted Invoice Detail Record Layout
                        If iField <= 28 Then

                            '//**** Lookup Field Name
                            sFieldName = aRIDRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aRIDRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aRIDRL(iField, 3))

                            '//****
                            sSection = "RIDRL"

                            '//****
                            sRecordType = "Detail"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "A" Posted Invoice Applicaton Record Layout
                      ElseIf sRecord = "A" Then

                        '//**** Posted Invoice Applicaton Record Layout
                        If iField <= 12 Then

                            '//**** Lookup Field Name
                            sFieldName = aRIARL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aRIARL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aRIARL(iField, 3))

                            '//****
                            sSection = "RIARL"

                            '//****
                            sRecordType = "Application Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "E" Posted Invoice Tax Header Record Layout
                      ElseIf sRecord = "E" Then

                        '//**** Posted Invoice Tax Header Record Layout
                        If iField <= 10 Then

                            '//**** Lookup Field Name
                            sFieldName = aRIERL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aRIERL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aRIERL(iField, 3))

                            '//****
                            sSection = "RIERL"

                            '//****
                            sRecordType = "Tax Header Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "T" Posted Invoice Tax Detail Record Layout
                      ElseIf sRecord = "T" Then

                        '//**** Posted Invoice Tax Detail Record Layout
                        If iField <= 12 Then

                            '//**** Lookup Field Name
                            sFieldName = aRITRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aRITRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aRITRL(iField, 3))

                            '//****
                            sSection = "RITRL"

                            '//****
                            sRecordType = "Tax Details Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Else somthing is wrong like
                        '//     Posted Invoice Import Type
                        '//     indicated but Posted Voucher
                        '//     import file indicated
                      Else

                        If sRecord = "V" And iField = 1 Then
                            '//**** Error Tracking
                            CPATracker.Tracker "Verify " & sImportType & " is the " & vbCr _
                                                & "correct Import Type for file: " & vbCr _
                                                & sFileName, "LogFile.log", True, True

                        End If

                        '//**** Unknown Field Description
                        sFieldName = "                   "

                        '//**** Increment Error Count
                        lErrorCount = lErrorCount + 1

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                    End If

                    '//**** Compare the Import Type string sImportType to "AR Pending Invoices"
                  ElseIf sImportType = "AR Pending Invoices" Then

                    '//**** Compare the Record Type ID to "B"
                    If sRecord = "B" Then

                        '//**** Pending Invoices Record Layout
                        If iField <= 17 Then

                            '//**** Lookup Field Name
                            sFieldName = aPIBRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aPIBRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aPIBRL(iField, 3))

                            '//**** Check if this is a new batch
                            If iField = 1 Then

                                '//**** Increment BatchCounter
                                lBatchCount = lBatchCount + 1

                            End If

                            '//****
                            sSection = "PIBRL"

                            '//****
                            sRecordType = "Batch Record"

                            '//**** Parse out 11th field of the sDataLine as sCompany
                            sCompany = CPAParseDelimited.ParseString(sDataLine, 11, sDelimiter)

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Lookup Field Description
                            sFieldDescription = "                   "

                            '//**** Assume undefined field is not required
                            bReqField = False

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "I"
                      ElseIf sRecord = "I" Then

                        '//**** Pending Voucher Header Record Layouts
                        If iField <= 84 Then

                            '//**** Lookup Field Name
                            sFieldName = aPIHRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aPIHRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aPIHRL(iField, 3))

                            '//****
                            sSection = "PIHRL"

                            '//****
                            sRecordType = "Header"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Compare the Record Type ID to "D"
                      ElseIf sRecord = "D" Then

                        '//**** Pending Invoice Detail Record Layout
                        If iField <= 44 Then

                            '//**** Lookup Field Name
                            sFieldName = aPIDRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aPIDRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aPIDRL(iField, 3))

                            '//****
                            sSection = "PIDRL"

                            '//****
                            sRecordType = "Detail Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                        '//**** Else somthing is wrong like
                        '//     Pending Invoice Import Type
                        '//     indicated but Pending Voucher
                        '//     import file indicated
                      Else

                        If sRecord = "V" And iField = 1 Then
                            '//**** Error Tracking
                            CPATracker.Tracker "Verify " & sImportType & " is the " & vbCr _
                                                & "correct Import Type for file: " & vbCr _
                                                & sFileName, "LogFile.log", True, True

                        End If

                        '//**** Unknown Field Description
                        sFieldName = "                   "

                        '//**** Increment Error Count
                        lErrorCount = lErrorCount + 1

                        '//**** Add Record to list
                        bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                    End If

                    '//**** Compare the Import Type string sImportType to "IM Items"
                  ElseIf sImportType = "IM Items" Then

                    '//**** Compare the Record Type ID to "I"
                    If sRecord = "I" Then

                        '//**** Inventory Management Item Record Layout
                        If iField <= 61 Then

                            '//**** Lookup Field Name
                            sFieldName = aIMIRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aIMIRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '//     otherwise, True is returned
                            bReqField = CBool(aIMIRL(iField, 3))

                            '//****
                            sSection = "IMIRL"

                            '//****
                            sRecordType = "Item Record"

                          '//****
                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        '//****
                        End If

                      '//**** Compare the Record Type ID to "U"
                      ElseIf sRecord = "U" Then

                        '//**** Inventory Management Item Unit Of Measure Record Layout
                        If iField <= 8 Then

                            '//**** Lookup Field Name
                            sFieldName = aIMURL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aIMURL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aIMURL(iField, 3))

                            '//****
                            sSection = "IMURL"

                            '//****
                            sRecordType = "Unit of Measure Record"

                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        End If

                        '//**** Compare the Record Type ID to "L"
                      ElseIf sRecord = "L" Then

                        '//**** Inventory Management Landed Cost Factor Record Layout
                        If iField <= 3 Then

                            '//**** Lookup Field Name
                            sFieldName = aIMLRL(iField, 1)

                            '//**** Lookup Field Description
                            sFieldDescription = aIMLRL(iField, 2)

                            '//**** Set required field boolean bRequired to True
                            '//     Use CBool to convert zero to False is returned
                            '/      otherwise, True is returned
                            bReqField = CBool(aIMLRL(iField, 3))

                            '//****
                            sSection = "IMLRL"

                            '//****
                            sRecordType = "Landed Cost Factor Record"

                          '//****
                          Else

                            '//**** Unknown Field Description
                            sFieldName = "                   "

                            '//**** Increment Error Count
                            lErrorCount = lErrorCount + 1

                        '//****
                        End If

                      Else

                        '//**** Unknown Field Description
                        sFieldName = "                   "

                        '//**** Lookup Field Description
                        sFieldDescription = "                   "

                        '//**** Assume undefined field is not required
                        bReqField = False

                        '//**** Increment Error Count
                        lErrorCount = lErrorCount + 1

                    End If

                    '//**** Increment BatchCounter
                    lBatchCount = 1

                    '//**** Add Record to list
                    bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                    '//****
                    sRecordType = ""

                  Else
                    '//**** Increment BatchCounter
                    lBatchCount = 1

                    '//**** Unknown Import Type and Record Layout
                    sFieldName = "                   "

                    '//****
                    sRecordType = "Unknown"

                    '//**** Add Record to list
                    bAddList = AddList(sRecord, iField, sFieldName, bReqField, sFieldDescription, sField)

                '//****
                End If

            '//****
            Next

        '//**** End Determination if the field is to be displayed
        End If

    '//**** Continue Do While...Loop through each line of import file lines
    Loop

    '//****
    Close #lngFile

    '//**** Indicate no changes to the import record displayed
    bEdit = False

    '//**** Fill in Status Bar text
    txtImportType = sImportType
    txtRecordType.Text = sRecordType
    txtCompany = sCompany
    txtRecords.Text = "Current Line: " & lCurrentLine & " "

    '//**** Check for errors as indicated by lErrorCount long variable
    If lErrorCount > 0 Then
        txtStatus.ForeColor = &HFF&
        '//**** Indicate error count in status bar
        txtStatus.Text = "Error(s): " & lErrorCount

        '//****
      Else
        txtStatus.ForeColor = &HFFFFFF
        '//**** Indicate no errors in status bar
        txtStatus.Text = "No Errors"

        '//****
    End If

    '//**** Exit Sub/Function before error handler

Exit Sub

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the ListLoad procedure"

    '//**** Error Handling
    Select Case Err.Number

        '//**** Handle VB error 6 Overflow
      Case 6

        '//**** Resume Next

        '//**** Subscript out of range occurred
      Case 9

        '//**** Resume Next

        '//**** Type mismatch occurred
      Case 13

        '//**** Resume Next

        '//**** Handle File error 52 Bad file name or number
      Case 52

        '//****
        MsgBox sMessage & vbLf & "Verify file name: " & sFileName _
               , vbInformation, "Verify file name"

        '//**** Error Tracking
        CPATracker.Tracker sMessage & "Verify file: " & sFileName, "LogFile.log", False, True

        '//****
        Exit Sub

        '//**** Handle File error 52 Bad file name or number
      Case 53

        MsgBox sMessage & vbLf & "Verify file name: " & sFileName _
               , vbInformation, "Verify file name"

        '//**** Error Tracking
        CPATracker.Tracker sMessage & "Verify file: " & sFileName, "LogFile.log", False, True

        '//****
        Exit Sub

        '//**** Input past end of file occurred
      Case 62

        '//**** Resume Next

        '//**** Handle File error 75 Path/File access error
      Case 75

        MsgBox sMessage & vbLf & "Verify file name: " & sFileName _
               , vbInformation, "Verify file name"

        '//**** Error Tracking
        CPATracker.Tracker sMessage & "Verify file: " & sFileName, "LogFile.log", False, True

        '//****
        Exit Sub

        '//**** Handle VB error 91 Object variable or With block variable not set
      Case 91

        '//**** Resume Next

        '//**** Handle VB error 440 Automation
      Case 440

        '//**** Resume Next

        '//**** Unhandled errors
      Case Else

        '//**** Resume Next

    End Select

    '//**** Error Tracking
    CPATracker.Tracker sMessage, "LogFile.log", False, True

    '//**** Continue function/procedure
    Resume Next

    '//****

End Sub


