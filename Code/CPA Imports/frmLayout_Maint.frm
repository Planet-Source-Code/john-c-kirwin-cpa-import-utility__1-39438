VERSION 5.00
Begin VB.Form frmLayout_Maint 
   Caption         =   " Layout Maintenance"
   ClientHeight    =   5010
   ClientLeft      =   7230
   ClientTop       =   6555
   ClientWidth     =   6885
   Icon            =   "frmLayout_Maint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRequired 
      Caption         =   "Required"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox lstFields 
      Height          =   3765
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   250
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   240
      Width           =   750
   End
   Begin VB.ComboBox cboSection 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   250
      Left            =   6000
      TabIndex        =   1
      ToolTipText     =   "Save"
      Top             =   600
      Width           =   750
   End
   Begin VB.TextBox txtValue 
      Height          =   3765
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label lblWarning 
      Caption         =   "View only mode - this task is under contruction. "
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Label lblKey 
      Caption         =   "Keys"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblSection 
      Caption         =   "Section"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblValue 
      Caption         =   "Value"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "frmLayout_Maint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//**************************************************************************
'// Layout Maintenance - User interface for maintaining CPAINI.ini file, which
'// contains the import layout information.
'//
'// Version:       6.20.0
'// Created:       12/26/2002 John C. Kirwin (JCK)
'// Modified:      03/03/2002 JCK - formatting
'//
'//**************************************************************************

'//**************************************************************************
'// Global Variables
'//**************************************************************************
Dim sMessage As String                                                         ' sMessage
Dim sINIKey As String

'//**** Object Declarations
Dim CPAFiles As cCPA_Files                                                    ' CPAFiles Object
Dim CPAINI As cCPA_INI                                                        ' CPAINI Object
Dim CPAStrings As cCPA_Strings                                                ' CPAStrings Object
Dim CPATracker As cCPA_Tracker                                                ' CPATracker Object

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

Private Sub lstFields_Click()

  '//**************************************************************************
  '// lstFields_Click
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

  Dim sSection As String

    txtValue = ""

    '//****
    sSection = cboSection.Text

    '//****
    GetKeyDefinition (sSection)

    cmdSave.Enabled = False

End Sub

Private Sub cmdSave_Click()

  '//**************************************************************************
  '// cmdSave - Save the KeyValue entry to INI file under the Key and Section indicated
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

  Dim sLayoutINI As String
  Dim sWriteINI As Long

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"         ' Layouts.ini

    '//**** Check if KeyValue field blank
    If txtValue.Text = "" Then '

        '//**** Prompt to enter KeyValue
        MsgBox "Please enter Definition information to save", vbExclamation

        '//**** KeyValue field not blank
      Else

        '//**** Execute WriteToINI procedure
        sWriteINI = CPAINI.WriteToINI(cboSection.Text, _
                    sINIKey, _
                    txtValue.Text, sLayoutINI)

    End If

    '//**** Exit Sub/Function before error handler

    FillFieldList

    cmdSave.Enabled = False

Exit Sub

EH:
    sMessage = Err.Number & ": " & Err.Description & _
               " occurred during the WriteToINI procedure"
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

Private Sub txtValue_Validate(Cancel As Boolean)

  '//**** Put together the key that is getting updated

    sINIKey = Mid$(lstFields.Text, 1, 5) & "Description"

End Sub

Private Sub lstFields_GotFocus()

  Dim intMsg As Integer
  Dim sSection As String

    If lstFields.ListIndex <> -1 Then

        '//****
        If cmdSave.Enabled Then

            cmdSave.Enabled = False

            '//****
            intMsg = MsgBox("Import Layout Has Changed, Save Changes?", _
                     vbQuestion + vbYesNoCancel, Me.Caption)
            '//****
            Select Case intMsg

                '//****
              Case vbYes

                '//****
                Call cmdSave_Click

                '//****
              Case vbNo

                '//****
              Case vbCancel

                '//****
            End Select
            '//****
        End If

        '//****
    End If

End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)

    cmdSave.Enabled = True

End Sub

Private Sub SaveToggle()

    cmdSave.Enabled = Not cmdSave.Enabled

End Sub

Private Function GetKeyDefinition(sSection As String) As Boolean

  '//**************************************************************************
  '// GetKeyDifinition
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

  '//**** Evaluate if sSection = "GLNRL" GL Natural Account Record Layout

    If sSection = "GLNRL" Then

        '//**** Display Definition
        txtValue.Text = aGLNRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aGLNRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "GLSRL" GL Account Segment Record Layout
      ElseIf sSection = "GLSRL" Then

        '//**** Display Definition
        txtValue.Text = aGLSRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aGLSRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "GLARL" GL Account Record Layout
      ElseIf sSection = "GLARL" Then

        '//**** Display Definition
        txtValue.Text = aGLARL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aGLARL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "GLBRL" GL Transactions Batch Header Record Layout
      ElseIf sSection = "GLBRL" Then

        '//**** Display Definition
        txtValue.Text = aGLBRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aGLBRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "GLTRL" GL Transactions Record Layout
      ElseIf sSection = "GLTRL" Then

        '//**** Display Definition
        txtValue.Text = aGLTRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aGLTRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "GLFRL" GL Transactions Batch Footer Record Layout
      ElseIf sSection = "GLFRL" Then

        '//**** Display Definition
        txtValue.Text = aGLFRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aGLFRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "VPBRL" Posted Voucher Batch Record Layout array (17 Fields)
      ElseIf sSection = "VPBRL" Then

        '//**** Display Definition
        txtValue.Text = aVPBRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aVPBRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "VPHRL" Posted Voucher Header Record Layout array
      ElseIf sSection = "VPHRL" Then

        '//**** Display Definition
        txtValue.Text = aVPHRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aVPHRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "VPDRL" Posted Voucher Detail Record Layout array
      ElseIf sSection = "VPDRL" Then

        '//**** Display Definition
        txtValue.Text = aVPDRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aVPDRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "VPARL" Posted Voucher Application Record Layout array
      ElseIf sSection = "VPARL" Then

        '//**** Display Definition
        txtValue.Text = aVPARL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aVPARL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "VPXRL" Posted Voucher Tax Header Record Layout array
      ElseIf sSection = "VPXRL" Then

        '//**** Display Definition
        txtValue.Text = aVPXRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aVPXRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "VPTRL" Posted Voucher Tax Detail Record Layout array
      ElseIf sSection = "VPTRL" Then

        '//**** Display Definition
        txtValue.Text = aVPTRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aVPTRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "PVBRL" Pending Voucher Batch Record Layout
      ElseIf sSection = "PVBRL" Then

        '//**** Display Definition
        txtValue.Text = aPVBRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aPVBRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "PVHRL" Pending Voucher Header Record Layout
      ElseIf sSection = "PVHRL" Then

        '//**** Display Definition
        txtValue.Text = aPVHRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aPVHRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "PVDRL" Pending Voucher Detail Record Layout
      ElseIf sSection = "PVDRL" Then

        '//**** Display Definition
        txtValue.Text = aPVDRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aPVDRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "RIBRL" Posting Invoice Batch Record Layout array (17 Fields)
      ElseIf sSection = "RIBRL" Then

        '//**** Display Definition
        txtValue.Text = aRIBRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aRIBRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "RIRRL" Posting Invoice Header Record Layout array (84 Fields)
      ElseIf sSection = "RIRRL" Then

        '//**** Display Definition
        txtValue.Text = aRIRRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aRIRRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "RIDRL" Posting Invoice Detail Record Layout array (44 Fields)
      ElseIf sSection = "RIDRL" Then

        '//**** Display Definition
        txtValue.Text = aRIDRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aRIDRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "RIARL" Posting Invoice Application Record Layout array
      ElseIf sSection = "RIARL" Then

        '//**** Display Definition
        txtValue.Text = aRIARL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aRIARL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "RIERL" Posting Invoice Tax Header Record Layout array
      ElseIf sSection = "RIERL" Then

        '//**** Display Definition
        txtValue.Text = aRIERL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aRIERL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "RITRL" Posting Invoice Tax Detail Record Layout array
      ElseIf sSection = "RITRL" Then

        '//**** Display Definition
        txtValue.Text = aRITRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aRITRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "PIBRL" Pending Invoice Batch Record Layout array (17 Fields)
      ElseIf sSection = "PIBRL" Then

        '//**** Display Definition
        txtValue.Text = aPIBRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aPIBRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "PIHRL" Pending Invoice Header Record Layout array (84 Fields)
      ElseIf sSection = "PIHRL" Then

        '//**** Display Definition
        txtValue.Text = aPIHRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aPIHRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "PIDRL" Pending Invoice Detail Record Layout array (44 Fields)
      ElseIf sSection = "PIDRL" Then

        '//**** Display Definition
        txtValue.Text = aPIDRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aPIDRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "IMIRL" Inventory Management Item Record Layout array (61 Fields)
      ElseIf sSection = "IMIRL" Then

        '//**** Display Definition
        txtValue.Text = aIMIRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aIMIRL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "IMURL" Inventory Management Item Unit Of Measure Record Layout
      ElseIf sSection = "IMURL" Then

        '//**** Display Definition
        txtValue.Text = aIMURL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aIMURL(lstFields.ListIndex + 1, 3)

        '//**** Evaluate if sSection = "IMLRL" Inventory Management Landed Cost Factor Record Layout
      ElseIf sSection = "IMLRL" Then

        '//**** Display Definition
        txtValue.Text = aIMLRL(lstFields.ListIndex + 1, 2)
        '//**** Indicate Required Field or Not
        chkRequired.Value = aIMLRL(lstFields.ListIndex + 1, 3)

      Else

        GetKeyDefinition = True

    End If

    GetKeyDefinition = True

End Function

Private Sub cboSection_Click()

  '//**************************************************************************
  '// cboSection_Change
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

    FillFieldList

End Sub

Private Sub FillFieldList()

  Dim sSection As String

    sSection = cboSection.Text

    If GetLayout(sSection) Then
        Debug.Print "Got Layout"
    End If

End Sub

Private Sub cboSection_Change()

  '//**************************************************************************
  '// cboSection_Change
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

    FillFieldList

End Sub

Private Function GetLayout(sSection As String) As Boolean

  Dim iPos As Integer
  Dim iRows As Integer                                                           ' Total Rows in the Array
  Dim iColumns As Integer                                                        ' Total Columns in the Array

    lstFields.Clear

  Dim sLayoutINI As String

    '//**** Get path of CPAINI.ini
    sLayoutINI = CPAFiles.GetFilePath(App.Path & "\") & "Layouts.ini"             ' Layouts.ini

    '//**** Evaluate if sSection = "GLNRL" GL Natural Accounts Record Layout array (10 Fields)
    If sSection = "GLNRL" Then

        '//**** Define the dynamic array size variables
        iRows = 10
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aGLNRL(iRows, iColumns)

        '//**** PVBRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aGLNRL(iPos, 1) = CPAINI.GetFromINI("GLNRL", Trim$("N" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aGLNRL(iPos, 2) = CPAINI.GetFromINI("GLNRL", Trim$("N" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aGLNRL(iPos, 3) = CPAINI.GetFromINI("GLNRL", Trim$("N" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("N" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aGLNRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "GLSRL" GL Accounts Segment Record Layout array
      ElseIf sSection = "GLSRL" Then

        '//**** Define the dynamic array size variables
        iRows = 7
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aGLSRL(iRows, iColumns)

        '//**** PVBRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aGLSRL(iPos, 1) = CPAINI.GetFromINI("GLSRL", Trim$("S" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aGLSRL(iPos, 2) = CPAINI.GetFromINI("GLSRL", Trim$("S" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aGLSRL(iPos, 3) = CPAINI.GetFromINI("GLSRL", Trim$("S" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("S" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aGLSRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "GLARL" GL Accounts Record Layout array
      ElseIf sSection = "GLARL" Then

        '//**** Define the dynamic array size variables
        iRows = 8
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aGLARL(iRows, iColumns)

        '//**** PVBRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aGLARL(iPos, 1) = CPAINI.GetFromINI("GLARL", Trim$("A" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aGLARL(iPos, 2) = CPAINI.GetFromINI("GLARL", Trim$("A" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aGLARL(iPos, 3) = CPAINI.GetFromINI("GLARL", Trim$("A" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("A" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aGLARL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "GLTRL" GL Transaction Record Layout array
      ElseIf sSection = "GLTRL" Then

        '//**** Define the dynamic array size variables
        iRows = 19
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aGLTRL(iRows, iColumns)

        '//**** PVBRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aGLTRL(iPos, 1) = CPAINI.GetFromINI("GLTRL", Trim$("T" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aGLTRL(iPos, 2) = CPAINI.GetFromINI("GLTRL", Trim$("T" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aGLTRL(iPos, 3) = CPAINI.GetFromINI("GLTRL", Trim$("T" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("T" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & " " & aGLTRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "GLBRL" GL Transaction Batch Record Layout array
      ElseIf sSection = "GLBRL" Then

        '//**** Define the dynamic array size variables
        iRows = 4
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aGLBRL(iRows, iColumns)

        '//**** GLBRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aGLBRL(iPos, 1) = CPAINI.GetFromINI("GLBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aGLBRL(iPos, 2) = CPAINI.GetFromINI("GLBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aGLBRL(iPos, 3) = CPAINI.GetFromINI("GLBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("B" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aGLBRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "GLFRL" GL Transaction Footer Record Layout array
      ElseIf sSection = "GLFRL" Then

        '//**** Define the dynamic array size variables
        iRows = 2
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aGLFRL(iRows, iColumns)

        '//**** GLFRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aGLFRL(iPos, 1) = CPAINI.GetFromINI("GLFRL", Trim$("F" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aGLFRL(iPos, 2) = CPAINI.GetFromINI("GLFRL", Trim$("F" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aGLFRL(iPos, 3) = CPAINI.GetFromINI("GLFRL", Trim$("F" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("F" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aGLFRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "VPBRL" Posted Voucher Batch Record Layout array
      ElseIf sSection = "VPBRL" Then

        '//**** Define the dynamic array size variables
        iRows = 8
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aVPBRL(iRows, iColumns)

        '//**** VPBRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aVPBRL(iPos, 1) = CPAINI.GetFromINI("VPBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aVPBRL(iPos, 2) = CPAINI.GetFromINI("VPBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aVPBRL(iPos, 3) = CPAINI.GetFromINI("VPBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("B" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aVPBRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "VPHRL" Posted Voucher Header Record Layout array
      ElseIf sSection = "VPHRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 39
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aVPHRL(iRows, iColumns)

        '//**** VPHRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aVPHRL(iPos, 1) = CPAINI.GetFromINI("VPHRL", Trim$("V" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aVPHRL(iPos, 2) = CPAINI.GetFromINI("VPHRL", Trim$("V" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aVPHRL(iPos, 3) = CPAINI.GetFromINI("VPHRL", Trim$("V" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("V" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aVPHRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "VPDRL" Posted Voucher Detail Record Layout array
      ElseIf sSection = "VPDRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 25
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aVPDRL(iRows, iColumns)

        '//**** VPDRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aVPDRL(iPos, 1) = CPAINI.GetFromINI("VPDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aVPDRL(iPos, 2) = CPAINI.GetFromINI("VPDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aVPDRL(iPos, 3) = CPAINI.GetFromINI("VPDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("D" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aVPDRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "VPARL" Posted Voucher Application Record Layout array
      ElseIf sSection = "VPARL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 10
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aVPARL(iRows, iColumns)

        '//**** VPARL
        Do Until iPos = iRows
            iPos = iPos + 1
            aVPARL(iPos, 1) = CPAINI.GetFromINI("VPARL", Trim$("A" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aVPARL(iPos, 2) = CPAINI.GetFromINI("VPARL", Trim$("A" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aVPARL(iPos, 3) = CPAINI.GetFromINI("VPARL", Trim$("A" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("A" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aVPARL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "VPXRL" Posted Voucher Tax Header Record Layout array
      ElseIf sSection = "VPXRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 15
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aVPXRL(iRows, iColumns)

        '//**** VPXRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aVPXRL(iPos, 1) = CPAINI.GetFromINI("VPXRL", Trim$("X" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aVPXRL(iPos, 2) = CPAINI.GetFromINI("VPXRL", Trim$("X" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aVPXRL(iPos, 3) = CPAINI.GetFromINI("VPXRL", Trim$("X" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("X" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aVPXRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "VPTRL" Posted Voucher Tax Detail Record Layout array
      ElseIf sSection = "VPTRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 16
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aVPTRL(iRows, iColumns)

        '//**** VPTRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aVPTRL(iPos, 1) = CPAINI.GetFromINI("VPTRL", Trim$("T" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aVPTRL(iPos, 2) = CPAINI.GetFromINI("VPTRL", Trim$("T" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aVPTRL(iPos, 3) = CPAINI.GetFromINI("VPTRL", Trim$("T" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("T" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aVPTRL(iPos, 1))

                '//****
            End With

            '//****
        Loop


        '//**** Evaluate if sSection = "PVBRL" Pending Voucher Batch Record Layout array
      ElseIf sSection = "PVBRL" Then

        '//**** Define the dynamic array size variables
        iRows = 17
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aPVBRL(iRows, iColumns)

        '//**** PVBRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aPVBRL(iPos, 1) = CPAINI.GetFromINI("PVBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aPVBRL(iPos, 2) = CPAINI.GetFromINI("PVBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aPVBRL(iPos, 3) = CPAINI.GetFromINI("PVBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("B" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aPVBRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "PVHRL" Pending Voucher Header Record Layout array
      ElseIf sSection = "PVHRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 75
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aPVHRL(iRows, iColumns)

        '//**** PIHRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aPVHRL(iPos, 1) = CPAINI.GetFromINI("PVHRL", Trim$("V" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aPVHRL(iPos, 2) = CPAINI.GetFromINI("PVHRL", Trim$("V" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aPVHRL(iPos, 3) = CPAINI.GetFromINI("PVHRL", Trim$("V" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("V" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aPVHRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "PVDRL" Pending Voucher Detail Record Layout array
      ElseIf sSection = "PVDRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 38
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aPVDRL(iRows, iColumns)

        '//**** PIDRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aPVDRL(iPos, 1) = CPAINI.GetFromINI("PVDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aPVDRL(iPos, 2) = CPAINI.GetFromINI("PVDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aPVDRL(iPos, 3) = CPAINI.GetFromINI("PVDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("D" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aPVDRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "RIBRL" Posted Invoice Batch Record Layout array
      ElseIf sSection = "RIBRL" Then

        '//**** Define the dynamic array size variables
        iRows = 9
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aRIBRL(iRows, iColumns)

        '//**** RIBRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aRIBRL(iPos, 1) = CPAINI.GetFromINI("RIBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aRIBRL(iPos, 2) = CPAINI.GetFromINI("RIBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aRIBRL(iPos, 3) = CPAINI.GetFromINI("RIBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("B" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aRIBRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "RIRRL" Posted Invoice Header Record Layout array
      ElseIf sSection = "RIRRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 40
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aRIRRL(iRows, iColumns)

        '//**** RIRRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aRIRRL(iPos, 1) = CPAINI.GetFromINI("RIRRL", Trim$("I" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aRIRRL(iPos, 2) = CPAINI.GetFromINI("RIRRL", Trim$("I" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aRIRRL(iPos, 3) = CPAINI.GetFromINI("RIRRL", Trim$("I" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("I" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aRIRRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "RIDRL" Posted Invoice Detail Record Layout array
      ElseIf sSection = "RIDRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 28
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aRIDRL(iRows, iColumns)

        '//**** RIDRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aRIDRL(iPos, 1) = CPAINI.GetFromINI("RIDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aRIDRL(iPos, 2) = CPAINI.GetFromINI("RIDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aRIDRL(iPos, 3) = CPAINI.GetFromINI("RIDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("D" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aRIDRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "RIARL" Posted Invoice Application Record Layout array
      ElseIf sSection = "RIARL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 12
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aRIARL(iRows, iColumns)

        '//**** RIARL
        Do Until iPos = iRows
            iPos = iPos + 1
            aRIARL(iPos, 1) = CPAINI.GetFromINI("RIARL", Trim$("A" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aRIARL(iPos, 2) = CPAINI.GetFromINI("RIARL", Trim$("A" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aRIARL(iPos, 3) = CPAINI.GetFromINI("RIARL", Trim$("A" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("A" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aRIARL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "RIERL" Posted Invoice Tax Header Record Layout array
      ElseIf sSection = "RIERL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 10
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aRIERL(iRows, iColumns)

        '//**** RIERL
        Do Until iPos = iRows
            iPos = iPos + 1
            aRIERL(iPos, 1) = CPAINI.GetFromINI("RIERL", Trim$("E" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aRIERL(iPos, 2) = CPAINI.GetFromINI("RIERL", Trim$("E" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aRIERL(iPos, 3) = CPAINI.GetFromINI("RIERL", Trim$("E" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("E" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aRIERL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "RITRL" Posted Invoice Tax Detail Record Layout array
      ElseIf sSection = "RITRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 12
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aRITRL(iRows, iColumns)

        '//**** RITRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aRITRL(iPos, 1) = CPAINI.GetFromINI("RITRL", Trim$("T" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aRITRL(iPos, 2) = CPAINI.GetFromINI("RITRL", Trim$("T" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aRITRL(iPos, 3) = CPAINI.GetFromINI("RITRL", Trim$("T" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("T" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aRITRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "PIBRL" Pending Invoice Batch Record Layout array
      ElseIf sSection = "PIBRL" Then

        '//**** Define the dynamic array size variables
        iRows = 17
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aPIBRL(iRows, iColumns)

        '//**** PIBRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aPIBRL(iPos, 1) = CPAINI.GetFromINI("PIBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aPIBRL(iPos, 2) = CPAINI.GetFromINI("PIBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aPIBRL(iPos, 3) = CPAINI.GetFromINI("PIBRL", Trim$("B" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("B" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aPIBRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "PIHRL" Pending Invoice Header Record Layout array
      ElseIf sSection = "PIHRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 84
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aPIHRL(iRows, iColumns)

        '//**** PIHRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aPIHRL(iPos, 1) = CPAINI.GetFromINI("PIHRL", Trim$("I" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aPIHRL(iPos, 2) = CPAINI.GetFromINI("PIHRL", Trim$("I" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aPIHRL(iPos, 3) = CPAINI.GetFromINI("PIHRL", Trim$("I" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("I" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aPIHRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "PIDRL" Pending Invoice Detail Record Layout array
      ElseIf sSection = "PIDRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 44
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aPIDRL(iRows, iColumns)

        '//**** PIDRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aPIDRL(iPos, 1) = CPAINI.GetFromINI("PIDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aPIDRL(iPos, 2) = CPAINI.GetFromINI("PIDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aPIDRL(iPos, 3) = CPAINI.GetFromINI("PIDRL", Trim$("D" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("D" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aPIDRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "IMIRL" Inventory Management Item Record Layout array
      ElseIf sSection = "IMIRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 61
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aIMIRL(iRows, iColumns)

        '//**** IMIRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aIMIRL(iPos, 1) = CPAINI.GetFromINI("IMIRL", Trim$("I" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aIMIRL(iPos, 2) = CPAINI.GetFromINI("IMIRL", Trim$("I" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aIMIRL(iPos, 3) = CPAINI.GetFromINI("IMIRL", Trim$("I" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("I" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aIMIRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "IMURL" Inventory Management Item Unit of Measure Record Layout array
      ElseIf sSection = "IMURL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 8
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aIMURL(iRows, iColumns)

        '//**** IMURL
        Do Until iPos = iRows
            iPos = iPos + 1
            aIMURL(iPos, 1) = CPAINI.GetFromINI("IMURL", Trim$("U" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aIMURL(iPos, 2) = CPAINI.GetFromINI("IMURL", Trim$("U" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aIMURL(iPos, 3) = CPAINI.GetFromINI("IMURL", Trim$("U" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("U" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aIMURL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Evaluate if sSection = "IMLRL" Inventory Management Landed Cost Factor Record Layout array
      ElseIf sSection = "IMLRL" Then

        '//**** Define the size of the dynamic array MyArray()
        iRows = 3
        iColumns = 3
        iPos = 0

        '//**** Re Dimension the two-dimensional array
        '//     allocating elements based on iRows
        '//     and iColumns
        ReDim aIMLRL(iRows, iColumns)

        '//**** IMLRL
        Do Until iPos = iRows
            iPos = iPos + 1
            aIMLRL(iPos, 1) = CPAINI.GetFromINI("IMLRL", Trim$("L" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_FieldName"), sLayoutINI)
            aIMLRL(iPos, 2) = CPAINI.GetFromINI("IMLRL", Trim$("L" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Description"), sLayoutINI)
            aIMLRL(iPos, 3) = CPAINI.GetFromINI("IMLRL", Trim$("L" & _
                   CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & _
                   "_Required"), sLayoutINI)

            '//****
            With lstFields

                '//****
                .AddItem Trim$("L" & CPAStrings.FillAlign(CStr(iPos), 3, "0", True) & "_" & aIMLRL(iPos, 1))

                '//****
            End With

            '//****
        Loop

        '//**** Else Unknown Layout
      Else

        '//****
        GetLayout = False

        '//**** End if Layout Evaluation
    End If

    '//****
    lstFields.ListIndex = 0

    '//****
    GetLayout = True

End Function

Private Sub Form_Load()

  '//**************************************************************************
  '// Form_Load
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

    Set CPAINI = New cCPA_INI
    Set CPATracker = New cCPA_Tracker
    Set CPAStrings = New cCPA_Strings
    Set CPAFiles = New cCPA_Files

    With cboSection
        .AddItem "GLNRL" 'GL Natural Account Record
        .AddItem "GLSRL" 'GL Record
        .AddItem "GLARL" 'GL Account Record
        .AddItem "GLTRL" 'GL Transaction Record
        .AddItem "GLBRL" 'GL Transaction Batch Record
        .AddItem "GLFRL" 'GL Transaction Footer Record
        .AddItem "VPBRL" 'AP Posted Voucher Batch Record
        .AddItem "VPHRL" 'AP Posted Voucher Header Record
        .AddItem "VPDRL" 'AP Posted Voucher Detail Record
        .AddItem "VPARL" 'AP Posted Voucher Application Record
        .AddItem "VPXRL" 'AP Posted Voucher Tax Header Record
        .AddItem "VPTRL" 'AP Posted Voucher Tax Detail Record
        .AddItem "PVBRL" 'AP Pending Voucher Batch Record
        .AddItem "PVHRL" 'AP Pending Voucher Header Record
        .AddItem "PVDRL" 'AP Pending Voucher Detail Record
        .AddItem "RIBRL" 'AR Posted Invoice Batch Record
        .AddItem "RIRRL" 'AR Posted Invoice Header Record
        .AddItem "RIDRL" 'AR Posted Invoice Detail Record
        .AddItem "RIARL" 'AR Posted Invoice Application Record
        .AddItem "RIERL" 'AR Posted Invoice Tax Header Record
        .AddItem "RITRL" 'AR Posted Invoice Tax Detail Record
        .AddItem "PIBRL" 'AR Pending Invoice Batch
        .AddItem "PIHRL" 'AR Pending Header Record
        .AddItem "PIDRL" 'AR Pending Detail Record
        .AddItem "IMIRL" 'IM Item Record
        .AddItem "IMURL" 'IM Unit of Measure Record
        .AddItem "IMLRL" 'IM Record
    End With

    cboSection.ListIndex = 0

End Sub

Private Sub cmdClose_Click()

  '//**************************************************************************
  '// Close the application
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

  '//**** Unload the form

    Unload Me

    '//**** Show the Import Utility Main from
    frmImport_Utility.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

  '//**************************************************************************
  '// Form_Unload
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

    Set frmLayout_Maint = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  '//**************************************************************************
  '// Form_QueryUnload
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

End Sub

