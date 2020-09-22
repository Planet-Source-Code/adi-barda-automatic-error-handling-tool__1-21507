VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Auto Error handling 1.0  - (c) 2001 written by adi barda 052-721721"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowInterface 
      Caption         =   "Show interface"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   41
      ToolTipText     =   "Show selected file interface"
      Top             =   4590
      Width           =   1425
   End
   Begin VB.TextBox txtControlPrefixes 
      Height          =   255
      Left            =   2580
      TabIndex        =   40
      Text            =   "cmd,chk,lbl,cbo,lst,txt,opt,img"
      Top             =   7770
      Width           =   3255
   End
   Begin VB.CheckBox chkIgnoreControlsPrefix 
      Caption         =   "Ignore functions starting with:"
      Height          =   255
      Left            =   210
      TabIndex        =   39
      ToolTipText     =   "äúòìí îôåð÷öéåú ùîúçéìåú áî÷ãîéí äáàéí:"
      Top             =   7740
      Value           =   1  'Checked
      Width           =   2385
   End
   Begin VB.CommandButton cmdUnSelectFunc 
      Caption         =   "-"
      Height          =   285
      Left            =   10860
      TabIndex        =   38
      ToolTipText     =   "Un select"
      Top             =   4590
      Width           =   405
   End
   Begin VB.CommandButton cmdSelectFunc 
      Caption         =   "+"
      Height          =   285
      Left            =   11310
      TabIndex        =   37
      ToolTipText     =   "Select all"
      Top             =   4590
      Width           =   405
   End
   Begin VB.CommandButton cmdUnSelectFiles 
      Caption         =   "-"
      Height          =   285
      Left            =   6570
      TabIndex        =   36
      ToolTipText     =   "Un select"
      Top             =   4590
      Width           =   405
   End
   Begin VB.CommandButton cmdSelectFiles 
      Caption         =   "+"
      Height          =   285
      Left            =   7020
      TabIndex        =   35
      ToolTipText     =   "Select all"
      Top             =   4590
      Width           =   405
   End
   Begin VB.CheckBox chkIgnoreOnErr 
      Caption         =   "Ignore functions with ""ON ERROR"" commands"
      Height          =   255
      Left            =   210
      TabIndex        =   34
      ToolTipText     =   "äúòìí îôåð÷öéåú äîëéìåú èéôåì áùâéàåú"
      Top             =   7440
      Value           =   1  'Checked
      Width           =   3765
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   33
      ToolTipText     =   "Compare original file with the new file"
      Top             =   4590
      Width           =   795
   End
   Begin VB.CommandButton cmdCommit 
      Caption         =   "Commit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   32
      ToolTipText     =   "Add error handling to the selected files and functions"
      Top             =   4590
      Width           =   795
   End
   Begin VB.ListBox lstFunctions 
      Height          =   4335
      Left            =   7590
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   210
      Width           =   4125
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3360
      TabIndex        =   29
      ToolTipText     =   "Clear list"
      Top             =   4590
      Width           =   795
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Transfer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9270
      TabIndex        =   28
      ToolTipText     =   "äçìôú ä÷áöéí äî÷åøééí á÷áöéí äîòåáãéí"
      Top             =   7650
      Width           =   1125
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   10440
      TabIndex        =   27
      ToolTipText     =   "éöéàä îäîòøëú"
      Top             =   7650
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Err handling:"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   3270
      TabIndex        =   16
      Top             =   5280
      Width           =   8385
      Begin VB.OptionButton optUseErrFunc 
         Caption         =   "Use Error handling function"
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   4350
         TabIndex        =   26
         Top             =   210
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optUseFreeText 
         Caption         =   "Use free text"
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   90
         TabIndex        =   25
         Top             =   210
         Width           =   1305
      End
      Begin VB.TextBox txtExtraParam 
         Height          =   285
         Left            =   5640
         TabIndex        =   24
         Text            =   "Err_Handle_Mode"
         Top             =   1200
         Width           =   1755
      End
      Begin VB.CheckBox chkErrObj 
         Caption         =   "Err object"
         Height          =   255
         Left            =   7110
         TabIndex        =   23
         Top             =   810
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CheckBox chkModuleName 
         Caption         =   "Module name"
         Height          =   255
         Left            =   4440
         TabIndex        =   22
         Top             =   810
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chkExtraParam 
         Caption         =   "Extra param"
         Height          =   255
         Left            =   4440
         TabIndex        =   21
         Top             =   1170
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox chkProcName 
         Caption         =   "Proc name"
         Height          =   255
         Left            =   5820
         TabIndex        =   20
         Top             =   810
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.TextBox txtFuncName 
         Height          =   285
         Left            =   5370
         TabIndex        =   18
         Text            =   "Err_Handler"
         Top             =   480
         Width           =   1995
      End
      Begin VB.TextBox txtErrHndl 
         Enabled         =   0   'False
         Height          =   975
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Text            =   "frmMain.frx":0442
         Top             =   450
         Width           =   3915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Func name:"
         Height          =   255
         Index           =   7
         Left            =   4380
         TabIndex        =   19
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.TextBox txtExitLabel 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Text            =   "Exit_Proc"
      Top             =   5640
      Width           =   1995
   End
   Begin VB.TextBox txtTabLength 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Text            =   "4"
      Top             =   6720
      Width           =   885
   End
   Begin VB.CheckBox chkApplyOnFunc 
      Caption         =   "Apply on functions"
      Height          =   255
      Left            =   2130
      TabIndex        =   11
      ToolTipText     =   "äçì èéôåì áùâéàåú òì ôåð÷öéåú"
      Top             =   7110
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkApplyOnProc 
      Caption         =   "Apply on procedures"
      Height          =   255
      Left            =   210
      TabIndex        =   10
      ToolTipText     =   "äçì èéôåì áùâéàåú òì ôøåöãåøåú"
      Top             =   7110
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtUpperGap 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Text            =   "3"
      Top             =   6060
      Width           =   885
   End
   Begin VB.TextBox txtLowerGap 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "2"
      Top             =   6390
      Width           =   885
   End
   Begin VB.TextBox txtErrLbl 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Text            =   "Err_Proc"
      Top             =   5310
      Width           =   1995
   End
   Begin VB.ListBox lstSelectedFiles 
      Height          =   4335
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   210
      Width           =   7365
   End
   Begin VB.CommandButton cmdBrows 
      Caption         =   "Brows"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Add vb project to the list"
      Top             =   4590
      Width           =   735
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   4050
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDefine 
      Caption         =   "Define"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Parse the files and find the functions"
      Top             =   4590
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Selected functions:"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   7590
      TabIndex        =   31
      Top             =   0
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit label:"
      Height          =   255
      Index           =   8
      Left            =   210
      TabIndex        =   15
      Top             =   5640
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tab length:"
      Height          =   285
      Index           =   6
      Left            =   210
      TabIndex        =   13
      Top             =   6690
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lower gap:"
      Height          =   285
      Index           =   5
      Left            =   210
      TabIndex        =   6
      Top             =   6360
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Err label:"
      Height          =   255
      Index           =   4
      Left            =   210
      TabIndex        =   5
      Top             =   5310
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Upper gap:"
      Height          =   285
      Index           =   3
      Left            =   210
      TabIndex        =   4
      Top             =   6030
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Selected files:"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private m_SourceFiles()         As String
Const C_MODULE_NAME = 0
Const C_PROC_NAME = 1
Const C_SELECTED = 2
Const C_IGNORED = 3

Private m_FilesCounter          As Long
Private m_bAvoidClick           As Boolean
Private m_AFunctions()          As Variant
Private m_AControlsPrefix()     As String


'* Function: AddErrHandling
'* Purpose: Add error handling to a certain file

Private Function AddErrHandling(ByVal sFilePath As String, _
                                ByVal FileNum As Long, _
                                Optional ByVal UseDefinition As Boolean = False) As Boolean

    'Purpose: Add error handling to the temporary file
    '         if UseDefinition = true,

    On Error GoTo Err_Proc
    
    Const PROCESS_REMARK = "'"
    
    Dim ff          As Long 'source file
    Dim ffDest      As Long 'dest file
    Dim ffDesc      As Long 'description file
    Dim s           As String
    Dim sline       As String
    Dim sDest       As String
    Dim sDestFile   As String
    Dim sModuleName As String 'current module name
    Dim sProcName   As String 'current procedure name
    Dim ProcIndex   As Long 'function index in array
    Dim i           As Long
    
    Dim bStartSub     As Boolean 'recognize function
    Dim bStartFunc    As Boolean 'recognize function
    Dim bEndSub       As Boolean 'recognize end of sub or function
    Dim bAddOnErr     As Boolean 'flag saying need to add on error statement
    Dim bOnErrorAdded As Boolean 'flag indicated whether on error added
    Dim bAddErrLbl    As Boolean
    Dim bHasOnError   As Boolean 'the function allredy has on error
    Dim bFoundModuleName As Boolean 'flag-found thew module name
    
    Dim iTopIndex     As Long 'optimization flag
    Dim oDir          As Scripting.FileSystemObject
    Dim sDesc         As String 'temp variable to store function description
    Dim iDesc         As Long 'function description counter
    
    'Init interface description array
    ReDim g_InterfaceDesc(0)
    sDesc = ""
    iDesc = 1
    
    Set oDir = New Scripting.FileSystemObject
    
    'gets the array size-number of functions in the system
    iTopIndex = UBound(m_AFunctions, 2)
    If m_AFunctions(0, 0) <> -1 Then iTopIndex = iTopIndex + 1 'case its not the first time
    
    'init vars
    sModuleName = ""
    sProcName = ""
    
    'open source file
    ff = FreeFile
    Open sFilePath For Input As #ff
    
    'open temp dest file
    ffDest = FreeFile
    sDestFile = GetDestFileName(sFilePath)
    
    'ensures that the temp files folder exists
    If Not oDir.FolderExists(App.Path & "\DestTmp") Then
        oDir.CreateFolder App.Path & "\DestTmp"
    End If
    
    Open App.Path & "\DestTmp\" & sDestFile For Output As #ffDest
    
    'open description file
    ffDesc = FreeFile()
    s = App.Path & "\DestTmp\" & sDestFile & ".desc"
    Open s For Output As #ffDesc

    
    'init algorithm flags
    s = ""
    bStartSub = False
    bEndSub = False
    bAddOnErr = False
    bOnErrorAdded = False
    bAddErrLbl = False
    bStartFunc = False
    bFoundModuleName = False
    
    'main scanning loop
    Do Until EOF(ff)
    
        'read the current line from the file
        Line Input #ff, sline
        
        'init dest line
        sDest = ""
        
        '*Check for the module name
        If Not bFoundModuleName Then
            sModuleName = GetModuleName(sline)
            bFoundModuleName = (LenB(sModuleName) <> 0)
            If bFoundModuleName Then
                sDesc = vbNewLine & "*****   " & sModuleName & " INTERFACE   *****" & vbNewLine
                For i = 1 To 100
                    sDesc = sDesc & "-"
                Next i
                sDesc = sDesc & vbNewLine & "Printed in " & Now & vbNewLine
                sDesc = sDesc & vbNewLine & vbNewLine
            End If
        End If
        
        '* check if its a begining of a sub or function
        '* Check subs:
        If (Not bStartSub) Then
            If LCase(Left$(sline, 11)) = "public sub " Then
                sProcName = GetProcName(sline, 12)
                bStartSub = ((Me.chkApplyOnProc.Value = 1) And (FunctionSelected(FileNum, sProcName, UseDefinition)))
                If bStartSub Then bHasOnError = False
            ElseIf LCase(Left$(sline, 4)) = "sub " Then
                sProcName = GetProcName(sline, 5)
                bStartSub = ((Me.chkApplyOnProc.Value = 1) And (FunctionSelected(FileNum, sProcName, UseDefinition)))
                If bStartSub Then bHasOnError = False
            ElseIf LCase(Left$(sline, 12)) = "private sub " Then
                sProcName = GetProcName(sline, 13)
                bStartSub = ((Me.chkApplyOnProc.Value = 1) And (FunctionSelected(FileNum, sProcName, UseDefinition)))
                If bStartSub Then bHasOnError = False
            End If
        Else
            If LCase(Left$(sline, 7)) = "end sub" Then
                bEndSub = True
            End If
        End If
        
        '* Check functions:
        If (Not bStartFunc) Then
            If LCase(Left$(sline, 16)) = "public function " Then
                sProcName = GetProcName(sline, 17)
                bStartFunc = ((Me.chkApplyOnFunc.Value = 1) And (FunctionSelected(FileNum, sProcName, UseDefinition)))
                If bStartFunc Then bHasOnError = False
            ElseIf LCase(Left$(sline, 9)) = "function " Then
                sProcName = GetProcName(sline, 10)
                bStartFunc = ((Me.chkApplyOnFunc.Value = 1) And (FunctionSelected(FileNum, sProcName, UseDefinition)))
                If bStartFunc Then bHasOnError = False
            ElseIf LCase(Left$(sline, 17)) = "private function " Then
                sProcName = GetProcName(sline, 18)
                bStartFunc = ((Me.chkApplyOnFunc.Value = 1) And (FunctionSelected(FileNum, sProcName, UseDefinition)))
                If bStartFunc Then bHasOnError = False
            End If
        Else
            If LCase(Left$(sline, 12)) = "end function" Then
                bEndSub = True
            End If
        End If
        
        If ((bStartSub) And (Not bAddOnErr)) Or ((bStartFunc) And (Not bAddOnErr)) Then
            '* check if after the current row i should insert on error goto..
            bAddOnErr = CheckAddOnErr(sline)
            sDesc = sDesc & sline & vbNewLine 'function description
            iDesc = 1
        End If
        
        'Build function description:
        If (InStr(1, Trim(sline), PROCESS_REMARK) = 1) Then  'print the remark only if it starts the line
            If (bStartSub Or bStartFunc) Then
                sDesc = sDesc & vbNewLine & Space$(4) & iDesc & ")  " & sline & vbNewLine
                iDesc = iDesc + 1
            Else
                sDesc = sDesc & vbNewLine & sline & vbNewLine
            End If
            
        End If
        
        '*check if the function allready has on error statement
        If Not bHasOnError Then
            bHasOnError = (InStr(1, LCase$(sline), "on error") > 0)
        End If
        
        
        
        If ((bStartSub) And (bAddOnErr) And (Not bOnErrorAdded)) Or _
           ((bStartFunc) And (bAddOnErr) And (Not bOnErrorAdded)) Then
            '* Add on error goto...
            sDest = sDest & sline
            For i = 1 To CLng(Me.txtUpperGap.Text)
                sDest = sDest & vbNewLine
            Next i
            
            sDest = sDest & Space$(CLng(Me.txtTabLength.Text)) & "On error goto " & Me.txtErrLbl.Text
            
            bOnErrorAdded = True
        End If
        
        
        '*Check if its end of sub or function
        If (bEndSub) Then
            sDest = Me.txtExitLabel.Text & ":" & vbNewLine
            If bStartFunc Then
                sDest = sDest & Space$(CLng(Me.txtTabLength.Text)) & "Exit function" & vbNewLine
            Else
                sDest = sDest & Space$(CLng(Me.txtTabLength.Text)) & "Exit sub" & vbNewLine
            End If
            
            '*Add label text
            sDest = sDest & vbNewLine & vbNewLine
            sDest = sDest & Me.txtErrLbl.Text & ":" & vbNewLine
            
            '*Add reference to err handling
            If Me.optUseFreeText.Value Then
                sDest = sDest & Space$(CLng(Me.txtTabLength.Text)) & Me.txtErrHndl.Text & vbNewLine
            Else
                sDest = sDest & GetErrFunctionConst(sModuleName, sProcName) & vbNewLine
            End If
            
            '*Resume to exit point
            sDest = sDest & Space$(CLng(Me.txtTabLength.Text)) & "Resume " & Me.txtExitLabel & vbNewLine
            
            
            '*Update functions array:
            If Not UseDefinition Then
                ReDim Preserve m_AFunctions(3, iTopIndex) 'allocates new memory unit
                m_AFunctions(C_MODULE_NAME, iTopIndex) = FileNum  'file num in lst
                m_AFunctions(C_PROC_NAME, iTopIndex) = sProcName  'function name
                m_AFunctions(C_SELECTED, iTopIndex) = 1  'put err handling by default
                m_AFunctions(C_IGNORED, iTopIndex) = Abs(((bHasOnError) And (Me.chkIgnoreOnErr.Value = 1)) Or (HasControlPrefix(sProcName))) 'ignore this function or not
                
                iTopIndex = iTopIndex + 1
            End If
            
            'insert lower gap space
            For i = 1 To CLng(Me.txtLowerGap.Text)
                sDest = sDest & vbNewLine
            Next i

            sDest = sDest & sline
            sDesc = sDesc & vbNewLine & sline & vbNewLine 'function description
            
            'print to description file
            Print #ffDesc, sDesc
                
            '*Clear variables:
            bStartSub = False
            bEndSub = False
            bAddOnErr = False
            bOnErrorAdded = False
            bAddErrLbl = False
            bStartFunc = False
            sProcName = ""
            sDesc = ""
        End If
        
        
        '* if nessesary insert default value
        If LenB(sDest) = 0 Then
            sDest = sline
        End If
        
        'prints to destination temp file
        Print #ffDest, sDest
               
    Loop
    
    'close file ports
    Close #ff
    Close #ffDest
    Close #ffDesc
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "AddErrHandling", Err, Err_Handle_Mode
    'Resume
Resume Exit_Proc


End Function


Private Function FunctionSelected(ByVal ModuleIndex As Long, ByVal sProcName As String, ByVal UseDefinition As Boolean, _
                                  Optional ByRef ProcIndex As Long) As Boolean


    On Error GoTo Err_Proc

    '*Purpose: checks if the function was selected and not ignored
    '*         function is ignored when user unmark its checkbox
    
    Dim i           As Long
    
    FunctionSelected = True
    ProcIndex = -1
    
    If Not UseDefinition Then Exit Function
    
    'scan the functions array
    For i = 0 To UBound(m_AFunctions, 2)
        If (m_AFunctions(1, i) = sProcName) And (m_AFunctions(0, i) = ModuleIndex) Then
            FunctionSelected = ((m_AFunctions(C_SELECTED, i) = 1) And (m_AFunctions(C_IGNORED, i) = 0))
            ProcIndex = i
            Exit For
        End If
    Next i
    
    
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler " frmMain ", "FunctionSelected", Err, Err_Handle_Mode
    Resume Exit_Proc


End Function


Private Function GetErrFunctionConst(ModuleName, ProcName) As String

    '*Purpose: gets the code for referencing to global error handling function
    
    On Error GoTo Err_Proc

    Dim s           As String
    
    'insert tab
    s = Space$(CLng(Me.txtTabLength.Text)) & Me.txtFuncName.Text & " "
    
    'insert function params
    If Me.chkModuleName.Value = 1 Then s = s & Chr$(34) & ModuleName & Chr$(34) & ", "
    If Me.chkProcName.Value = 1 Then s = s & Chr$(34) & ProcName & Chr$(34)
    If Me.chkProcName.Value = 1 Then s = s & ",Err"
    If Me.chkExtraParam.Value = 1 Then s = s & "," & Me.txtExtraParam.Text
    
    GetErrFunctionConst = s
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "GetErrFunctionConst", Err, Err_Handle_Mode
Resume Exit_Proc


End Function


Private Function CheckAddOnErr(ByVal CurrentLine As String) As Boolean

    '*Purpose: check if the current line contains on error statement
    
    On Error GoTo Err_Proc

    CheckAddOnErr = (InStr(1, CurrentLine, ")") > 0)
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "CheckAddOnErr", Err, Err_Handle_Mode
Resume Exit_Proc


End Function



Private Function CheckValidation() As Boolean

    '*Purpose: check that all the nesesary fields has data
    
    On Error GoTo Err_Proc


    Dim obj     As Control
    
    CheckValidation = False
    
    
    For Each obj In frmMain
        If TypeOf obj Is TextBox Then
            If obj.Name <> "txtSource" And obj.Name <> "txtDest" Then
                If Trim(obj.Text) = "" Then
                    Exit Function
                End If
            End If
            
        End If
        
        
    Next obj
    
    CheckValidation = True
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "CheckValidation", Err, Err_Handle_Mode
Resume Exit_Proc


End Function

Private Sub cmdBrows_Click()

    '*Purpose:Brows for a vb project or just one more free file
    '*        if vb project found than i load all its relevant code files
    
    Dim sFileName       As String
    Dim oDir            As Scripting.FileSystemObject
    
    
    'open dialog box
    dlg1.ShowOpen
    
    Set oDir = New Scripting.FileSystemObject
    
    'checks for valid file
    If Not dlg1.CancelError Then
        sFileName = dlg1.FileName
        
        'checks for file type
        If oDir.GetExtensionName(sFileName) <> "vbp" Then
            AddFileName sFileName 'other file
        Else
            AddProject sFileName 'vb project - add all relevant files
        End If
        
    End If
    
    'allow defining the selected files:
    Me.cmdDefine.Enabled = (Me.lstSelectedFiles.ListCount > 0)
    Me.cmdShowInterface.Enabled = False

End Sub


Private Sub AddProject(ByVal sFileName As String)


    On Error GoTo Err_Proc

    '*Purpose: adds the selected project (all its files) to the system manager
    
    Dim oDir        As Scripting.FileSystemObject
    Dim ff          As Long
    Dim i           As Long
    Dim sline       As String
    Dim sObjectName As String
    Dim sPath       As String
    
    Set oDir = New Scripting.FileSystemObject
    
    '*ensures backslash is exists
    sPath = oDir.GetParentFolderName(sFileName)
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    'open file port
    ff = FreeFile
    
    Open sFileName For Input As #ff
    
    'scan vb project file
    Do Until EOF(ff)
        Line Input #ff, sline 'read next line in the project file
        
        'check for the next object:
        If InStr(1, LCase$(sline), "form=") > 0 Then
            i = InStr(1, sline, "=") + 1
            sObjectName = Mid$(sline, i, Len(sline) - i + 1) 'find object name
            
            '*check that there is no (") in the object name
            If InStr(1, sObjectName, Chr$(34)) = 0 Then
                AddFileName sPath & sObjectName 'add file to list
            End If
            
        End If
        
        If InStr(1, LCase$(sline), "class=") > 0 Then
            i = InStr(1, sline, ";") + 2
            sObjectName = Mid$(sline, i, Len(sline) - i + 1)
            AddFileName sPath & sObjectName 'add file to list
        End If
        
        If InStr(1, LCase$(sline), "module=") > 0 Then
            i = InStr(1, sline, ";") + 2
            sObjectName = Mid$(sline, i, Len(sline) - i + 1)
            AddFileName sPath & sObjectName 'add file to list
        End If
        
        If InStr(1, LCase$(sline), "usercontrol=") > 0 Then
            i = InStr(1, sline, "=") + 1
            sObjectName = Mid$(sline, i, Len(sline) - i + 1)
            AddFileName sPath & sObjectName 'add file to list
        End If
        
    Loop
    
    'close project file port
    Close #ff
    
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "AddProject", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

Private Sub AddFileName(ByVal sFileName As String)


    On Error GoTo Err_Proc

    '*Purpose:adds new file to the files list
    
    If LenB(sFileName) > 0 Then
        If Not FileInList(sFileName) Then
            lstSelectedFiles.AddItem sFileName
            lstSelectedFiles.Selected(lstSelectedFiles.NewIndex) = True
        End If
    End If

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "AddFileName", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

Private Function FileInList(ByVal sFileName As String) As Boolean

    'check if the specified file is allready in the list
    
    On Error GoTo Err_Proc
    
    Dim i           As Long
    
    FileInList = False
    
    For i = 0 To Me.lstSelectedFiles.ListCount - 1
        FileInList = (sFileName = Me.lstSelectedFiles.List(i))
        If FileInList Then Exit For
    Next i
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "FileInList", Err, Err_Handle_Mode
Resume Exit_Proc


End Function

Private Sub cmdClear_Click()

    'clear all the files from the list
    
    On Error GoTo Err_Proc
    
    If MsgBox("Are you sure you want to clear the list ?", vbOKCancel Or vbQuestion) = vbOK Then
        Me.lstSelectedFiles.Clear
        Me.lstFunctions.Clear
        ReDim m_AFunctions(3, 0)
        
        Me.cmdTransfer.Enabled = False
        Me.cmdCommit.Enabled = False
        Me.cmdDefine.Enabled = False
        Me.cmdView.Enabled = False
        Me.cmdShowInterface.Enabled = False
        
    End If
    
Exit_Proc:
Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "cmdClear_Click", Err, Err_Handle_Mode
Resume Exit_Proc


End Sub

Private Sub cmdCommit_Click()

    'Purpose:Make the temporary files (generate error handling code)
    
    ProcessFiles True 'parse and make new files using predefine rules like
    '                  wich function needs err handling
    
    Me.cmdView.Enabled = True
    Me.cmdCommit.Enabled = False 'have to press "define" again
    Me.cmdShowInterface.Enabled = False
    
End Sub

Private Sub cmdDefine_Click()

    '*Parse the selected files and make temporary new files on the fly
    ProcessFiles False 'dont use the previuse definition
    Me.cmdCommit.Enabled = True
    Me.cmdShowInterface.Enabled = True
    If Me.lstSelectedFiles.ListCount > 0 Then
        Me.lstSelectedFiles.ListIndex = 0 'focus on the first file
        lstSelectedFiles_Click 'force showing the first file's functions
    End If
    
End Sub


Private Sub ProcessFiles(Optional ByVal UseDefinitions As Boolean = False)


    '*Purpose: parse all the selected files in the files list and generate
    '          err handling code for all the selected functions
    
    On Error GoTo Err_Proc
    
    Dim i           As Long
    
    If CheckValidation() Then
        
        '*Init functions array:
        If Not UseDefinitions Then
            ReDim m_AFunctions(3, 0)
            m_AFunctions(0, 0) = -1
        End If
        
        'Prepare controls prefix array
        If Me.chkIgnoreControlsPrefix.Value = 1 Then
            Me.txtControlPrefixes.Text = TrimEX(Me.txtControlPrefixes.Text) 'cut all spaces
            m_AControlsPrefix = Split(Me.txtControlPrefixes.Text, ",")
        Else
            ReDim m_AControlsPrefix(0) 'clear array
        End If
        
         'scan the files list
         For i = 0 To Me.lstSelectedFiles.ListCount - 1
             If Me.lstSelectedFiles.Selected(i) Then
                 If LenB(Me.lstSelectedFiles.List(i)) > 0 Then
                    'add err handling to the destination temp file
                     AddErrHandling Me.lstSelectedFiles.List(i), i, UseDefinitions
                 End If
             End If
             
        Next i
        Me.cmdTransfer.Enabled = True
        'MsgBox "File definition completed successfully"
    Else
        MsgBox "Cannot commit because one of the parameters is wrong"
        
   End If
   
   
Exit_Proc:
Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "cmdAdd_Click", Err, Err_Handle_Mode
Resume Exit_Proc

End Sub


Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdSelectFiles_Click()

    'select all files
    
    Dim i           As Long
    
    For i = 0 To Me.lstSelectedFiles.ListCount - 1
        Me.lstSelectedFiles.ListIndex = i
        Me.lstSelectedFiles.Selected(i) = True
    Next i

End Sub

Private Sub cmdSelectFunc_Click()

    'select all functions
    
    Dim i           As Long
    
    For i = 0 To Me.lstFunctions.ListCount - 1
        Me.lstFunctions.ListIndex = i
        Me.lstFunctions.Selected(i) = True
    Next i

End Sub

Private Sub cmdShowInterface_Click()

    'view source code vs generated code
    If LenB(Trim(Me.lstSelectedFiles.Text)) > 0 Then
        frmView.ShowEX Me.lstSelectedFiles.Text, True
    End If

End Sub

Private Sub cmdTransfer_Click()


    '*Purpose: Replace the original files with the generated files
    '          the generated files has err handling code in every function
    
    On Error GoTo Err_Proc

    Dim i           As Long
    
    i = MsgBox("Are you sure you want to replace the original files with the " & _
             "error handled files ?", vbOKCancel Or vbQuestion)
    If i = vbOK Then
        ReplaceFiles 'the final step
        MsgBox "The file transfer completed successfully"
    End If
    
    
Exit_Proc:
Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "cmdTransfer_Click", Err, Err_Handle_Mode
Resume Exit_Proc


End Sub

Private Sub cmdUnSelectFiles_Click()

    'unselect all the files
    
    Dim i           As Long
    
    For i = 0 To Me.lstSelectedFiles.ListCount - 1
        Me.lstSelectedFiles.ListIndex = i
        Me.lstSelectedFiles.Selected(i) = False
    Next i
    
End Sub

Private Sub cmdUnSelectFunc_Click()

    'unselect all the functions
    
    Dim i           As Long
    
    For i = 0 To Me.lstFunctions.ListCount - 1
        Me.lstFunctions.ListIndex = i
        Me.lstFunctions.Selected(i) = False
    Next i

End Sub

Private Sub cmdView_Click()

    'view source code vs generated code
    
    If LenB(Trim(Me.lstSelectedFiles.Text)) > 0 Then
        frmView.ShowEX Me.lstSelectedFiles.Text
    End If
    
End Sub

Private Sub Form_Load()


    On Error GoTo Err_Proc
    
    'init module-scope variables:
    Err_Handle_Mode = True
    m_FilesCounter = 0
    
    ReDim m_SourceFiles(0) 'source files container
    ReDim m_AFunctions(2, 0) 'functions definition
    ReDim m_AControlsPrefix(0) 'controls prefix
    
    m_AFunctions(0, 0) = -1 'no functions by default
    m_bAvoidClick = False
    
Exit_Proc:
Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "Form_Load", Err, Err_Handle_Mode
Resume Exit_Proc


End Sub

Private Sub UpdateSelectedFile(ByVal Item As Long)


    On Error GoTo Err_Proc

    '*Purpose: Mark some function in the current file -
    '          whether to add err handling code or not
    
    Dim sProcName           As String
    Dim iModuleIndex        As Long
    Dim iProcIndex          As Long
    
    sProcName = Me.lstFunctions.Text 'function name
    iModuleIndex = Me.lstSelectedFiles.ListIndex 'module num
    FunctionSelected iModuleIndex, sProcName, True, iProcIndex
    
    'update the functions array
    m_AFunctions(2, iProcIndex) = Abs(Me.lstFunctions.Selected(Item))
    
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "UpdateSelectedFile", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

Private Sub lstFunctions_ItemCheck(Item As Integer)

    If m_bAvoidClick Then Exit Sub
    UpdateSelectedFile Item 'update functions array

End Sub

Private Sub lstSelectedFiles_Click()

    'show all the functions in the module
    m_bAvoidClick = True
    ShowFunctions Me.lstSelectedFiles.ListIndex
    m_bAvoidClick = False
    
End Sub

Private Sub ShowFunctions(ByVal FileIndex As Long)


    On Error GoTo Err_Proc

    '*Purpose: show all the functions in the selected module
    
    Dim i               As Long
    Dim s               As String
    Dim sFuncName       As String
    Dim bFirstElement   As Boolean
    Dim bNoMore         As Boolean
    Dim iTopIndx        As Long
    
    bFirstElement = False
    bNoMore = False
    Me.lstFunctions.Clear
    
    i = 0
    iTopIndx = UBound(m_AFunctions, 2)
    
    'scan the functions array
    Do
        If m_AFunctions(0, i) = FileIndex Then
            If (Not bFirstElement) Then bFirstElement = (Not bFirstElement)
            sFuncName = m_AFunctions(1, i)
            Me.lstFunctions.AddItem sFuncName
            'If m_AFunctions(2, i) = 1 Then
                Me.lstFunctions.Selected(Me.lstFunctions.NewIndex) = (m_AFunctions(2, i) = 1)
            'End If
            
        Else
            If bFirstElement Then 'no more relevant functions
                bNoMore = True
            End If
            
        End If
        
        i = i + 1
        bNoMore = (i > iTopIndx)
        
    Loop Until bNoMore 'no more relevant functions
    
    
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "ShowFunctions", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

Private Sub optUseErrFunc_Click()


    On Error GoTo Err_Proc
    CheckErrHandling
Exit_Proc:
Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "optUseErrFunc_Click", Err, Err_Handle_Mode
Resume Exit_Proc


End Sub

Private Sub CheckErrHandling()

    'enable disable controls on the form
    On Error GoTo Err_Proc

    '*Enable / disable objects attached to err handling frame
    Me.chkErrObj.Enabled = Me.optUseErrFunc.Value
    Me.chkExtraParam.Enabled = Me.optUseErrFunc.Value
    Me.chkModuleName.Enabled = Me.optUseErrFunc.Value
    Me.chkProcName.Enabled = Me.optUseErrFunc.Value
    Me.txtExtraParam.Enabled = Me.optUseErrFunc.Value
    
    Me.txtErrHndl.Enabled = Me.optUseFreeText.Value
    
Exit_Proc:
Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "CheckErrHandling", Err, Err_Handle_Mode
Resume Exit_Proc


End Sub

Private Sub optUseFreeText_Click()


    On Error GoTo Err_Proc
    CheckErrHandling
Exit_Proc:
Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "optUseFreeText_Click", Err, Err_Handle_Mode
Resume Exit_Proc


End Sub


Private Sub ReplaceFiles()

    '* Well this is the final step: replacing the old files with the new files
    '  - wich has error handling code in all their functions

    On Error GoTo Err_Proc

    Dim i           As Long
    Dim sFile       As String
    
    On Error GoTo Err_Proc
    
    For i = 0 To Me.lstSelectedFiles.ListCount - 1
        If Me.lstSelectedFiles.Selected(i) Then
            sFile = GetDestFileName(Me.lstSelectedFiles.List(i))
            FileCopy App.Path & "\DestTmp\" & sFile, Me.lstSelectedFiles.List(i)
        End If
        
    Next i
    
    Exit Sub
    
    
Exit_Proc:
Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "ReplaceFiles", Err, Err_Handle_Mode
Resume Exit_Proc


End Sub

Private Function HasControlPrefix(ByVal sProcName As String) As Boolean


    On Error GoTo Err_Proc

    'Purpose: check if the function's name has a control prefix
    '         if so - do not put error handling in that function
    
    Dim i           As Long
    Dim iIndx       As Long
    
    HasControlPrefix = False 'has no prefix by default
    sProcName = Trim(sProcName)
    
    For i = 0 To UBound(m_AControlsPrefix)
        iIndx = InStr(1, sProcName, m_AControlsPrefix(i))
        HasControlPrefix = (iIndx = 1 And Me.chkIgnoreControlsPrefix.Value = 1)
        If HasControlPrefix Then Exit For
    Next i
        
    
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler " frmMain ", "HasControlPrefix", Err, Err_Handle_Mode
    Resume Exit_Proc


End Function

Private Function TrimEX(ByVal str As String) As String


    On Error GoTo Err_Proc

    Dim s           As String
    Dim i           As Long
    
    str = Trim(str)
    s = ""
    For i = 1 To Len(str)
        s = s & IIf(Mid$(str, i, 1) <> " ", Mid$(str, i, 1), "")
    Next i
    TrimEX = s
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler " frmMain ", "TrimEX", Err, Err_Handle_Mode
    Resume Exit_Proc


End Function
