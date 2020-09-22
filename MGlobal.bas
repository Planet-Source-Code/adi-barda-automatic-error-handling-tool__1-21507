Attribute VB_Name = "MGlobal"
Option Explicit
Public Err_Handle_Mode    As Boolean
Public g_InterfaceDesc()  As String


Public Sub Err_Handler(ByVal Module As String, ByVal Proc As String, Err As ErrObject, Err_Handle_Mode As Boolean)


    On error goto Err_Proc

    'Centeral error handling procedure
    
    If Err_Handle_Mode Then
        MsgBox "çìä ùâéàä áîåãåì:" & Module & "    áôåð÷öéä:" & Proc & vbNewLine & _
        "úàåø äùâéàä: " & Err.Description
        
    End If
    
Exit_Proc:
    Exit sub


Err_Proc:
    Err_Handler "MGlobal", "Err_Handler",Err,Err_Handle_Mode
    Resume Exit_Proc


End Sub

Public Function GetModuleName(ByVal sline As String) As String

    'Purpose: parse the module name from the initializing line
    
    On Error GoTo Err_Proc


    Dim iStart          As Long
    Dim iEnd            As Long
    Dim sEndChar        As String
    
    GetModuleName = ""
    
    
    If InStr(1, sline, "Attribute VB_Name = ") <> 0 Then
        iStart = InStr(1, sline, "Attribute VB_Name") + 21
        sEndChar = Chr$(34)
    ElseIf InStr(1, sline, "Begin VB.Form") <> 0 Then
        iStart = InStr(1, sline, "Begin VB.Form") + 13
        sEndChar = " "
    End If
    
    
    If iStart > 0 Then
        iEnd = InStr(iStart, sline, sEndChar)
        GetModuleName = Mid$(sline, iStart, iEnd)
    End If
    
    If Right$(GetModuleName, 1) = Chr$(34) Then
        GetModuleName = Left$(GetModuleName, Len(GetModuleName) - 1)
    End If
    
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "GetModuleName", Err, Err_Handle_Mode
Resume Exit_Proc


End Function

Public Function GetProcName(ByVal sline As String, ByVal StartPoint As Long) As String

    'Purpose: parse the procedure name from the initializing line
    
    On Error GoTo Err_Proc

    Dim iStartBr    As Long
    
    iStartBr = InStr(1, sline, "(")
    GetProcName = Mid$(sline, StartPoint, (iStartBr - StartPoint))
        
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "GetProcName", Err, Err_Handle_Mode
Resume Exit_Proc


End Function


Public Function GetDestFileName(ByVal sFilePath As String) As String

    'Purpose: return the temporary file name wich is going to be worked on
    
    On Error GoTo Err_Proc

    Dim oDir        As Scripting.FileSystemObject
    Dim sFileName   As String
    
    Set oDir = New Scripting.FileSystemObject
    
    sFileName = oDir.GetFileName(sFilePath)
    
    GetDestFileName = sFileName & ".tmp"
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "GetDestFileName", Err, Err_Handle_Mode
Resume Exit_Proc


End Function


