VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "System42"
   ClientHeight    =   210
   ClientLeft      =   1755
   ClientTop       =   4305
   ClientWidth     =   1560
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "LOAD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   210
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Loader..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1572
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim lstrSourcePath As String
Dim lstrDestinationPath As String
Dim lstrSourceFile As String
Dim lstrCheckFile As String
Dim lstrDestinationFile As String
Dim lbooCopyDone As Boolean
Dim llngShellID As Long
Dim lstrPassParam As String
Dim lstrErrorTrap As String
Dim lintProgCount As Integer
Dim lintArrInc As Integer
Dim lstrServerSupportPath As String
Dim lstrExeName As String
Dim lstrParam As String
Dim lintLenOfFileName As Integer
Dim lstrEFNTA As String

    Me.Top = 0
    Me.Left = 0
    
    Me.WindowState = vbNormal
    
    On Error Resume Next
    
    If UCase(Dir(Trim$(App.Path) & "\" & gconstrStaticLdr, vbNormal)) = UCase(gconstrStaticLdr) Then
        CheckStaticCipher
    Else
        MsgBox "Your Static loader file is not present!" & vbCrLf & vbCrLf & _
            "Please contact your System Administrator!", vbCritical
        End
    End If
    
    'Copy the latest static.ldr from the server folder
    FileCopy gstrStatic.strTrueLiveServerPath & gconstrStaticLdr, Trim$(App.Path) & "\" & gconstrStaticLdr
    
    'Re-read static.ldr
    If UCase(Dir(Trim$(App.Path) & "\" & gconstrStaticLdr, vbNormal)) = UCase(gconstrStaticLdr) Then
        CheckStaticCipher
    End If
    
    On Error GoTo LoadErrors
    
    lstrErrorTrap = "One"
        
    lstrPassParam = "/X" '" /X"
    
    'Setup paths
    If InStr(UCase(Command$), "/TEST") > 0 Then
        lstrErrorTrap = "Two"
        lstrPassParam = lstrPassParam & " /TEST"
        'When testing client, reps and admin exe files, TEST must be suffixed! /TEST for also me passed as a loader argument
        lstrEFNTA = "TEST"
        lstrSourcePath = gstrStatic.strServerTestNewPath
        lstrErrorTrap = "Three"
        lstrServerSupportPath = gstrStatic.strServerTestNewPath
        lstrErrorTrap = "Four"
    Else
        lstrErrorTrap = "Five"
        lstrEFNTA = ""
        lstrSourcePath = gstrStatic.strServerPath
        lstrErrorTrap = "Six"
        lstrServerSupportPath = gstrStatic.strSupportPath
        lstrErrorTrap = "Seven"
    End If
    gstrSourcePath = lstrSourcePath
        
    lstrDestinationPath = AppPath
    
    lstrErrorTrap = "Eight"
    
    'Check for loader.log and clear
    If Dir(App.Path & "\Loader.log") <> "" Then
        Kill App.Path & "\Loader.log"
    End If
    
    'Installs any new ocx / dll files
    InstallOCX lstrServerSupportPath
    
    lstrErrorTrap = "Nine"
    Me.Show
    DoEvents
    
    'Update Minder.exe
    lblStatus.Caption = "Downloading."
    DoEvents
    lstrErrorTrap = "Ten"
    lbooCopyDone = FileCopyIfNewer(lstrSourcePath & "Minder.exe", lstrDestinationPath & "Minder.exe")
    lstrErrorTrap = "Eleven"

    'CHK files are a copy of their server versions and used to tell if the server version has changed.
    'Obviously mdb file copies will change as they are used and can not be used to determine sever changes.
    
    'Setup variables for local.mdb check
    If InStr(UCase(Command$), "/TEST") > 0 Then
        lstrErrorTrap = "Twelve"
        lstrSourceFile = lstrSourcePath & gstrStatic.strShortLocalTestingDBFile
        
        lstrErrorTrap = "Thirteen"
        lintLenOfFileName = Len(gstrStatic.strShortLocalTestingDBFile)
        lstrErrorTrap = "Fourteen"
        lstrCheckFile = lstrDestinationPath & Left$(gstrStatic.strShortLocalTestingDBFile, lintLenOfFileName - 3) & "chk"
        
        lstrErrorTrap = "Fifteen"
        lstrDestinationFile = lstrDestinationPath & gstrStatic.strShortLocalTestingDBFile
        lstrErrorTrap = "Sixteen"
    Else
        lstrErrorTrap = "Seventeen"
        lstrSourceFile = lstrSourcePath & gstrStatic.strShortLocalDBFile
        lstrErrorTrap = "Eightteen"
        lintLenOfFileName = Len(gstrStatic.strShortLocalDBFile)
        lstrErrorTrap = "Nineteen"
        lstrCheckFile = lstrDestinationPath & Left$(gstrStatic.strShortLocalDBFile, lintLenOfFileName - 3) & "chk"
        lstrErrorTrap = "Twenty"
        lstrDestinationFile = lstrDestinationPath & gstrStatic.strShortLocalDBFile
        lstrErrorTrap = "TwentyOne"
    End If

    'Ensure that chk is available
    If Trim$(Dir(lstrCheckFile)) = "" Then
        lstrErrorTrap = "TwentyOneOne"
        FileCopy lstrSourceFile, lstrCheckFile
    End If
    
    lstrErrorTrap = "TwentyOneTwo"
    
    'Get mdb if newer
    If FileCopyIfNewer(lstrSourceFile, lstrCheckFile) Then
        lstrErrorTrap = "TwentyTwo"
        'Always overwite localdb if new
        FileCopy lstrCheckFile, lstrDestinationFile
        lstrErrorTrap = "TwentyThree"
    ElseIf DateValue(Date) > DateValue(FileDateTime(lstrDestinationFile)) Then
        lstrErrorTrap = "TwentyFour"
        'Overwite localdb once per day
        FileCopy lstrCheckFile, lstrDestinationFile
        lstrErrorTrap = "TwentyFive"
    End If
    
    lstrErrorTrap = "TwentySix"
    
    'Setup paths for Reps.mdb check
    If InStr(UCase(Command$), "/TEST") > 0 Then
        lstrErrorTrap = "TwentySeven"
        lstrSourceFile = lstrSourcePath & gstrStatic.strShortReportsTestingDBFile
        lstrErrorTrap = "TwentyEight"
        lintLenOfFileName = Len(gstrStatic.strShortReportsTestingDBFile)
        lstrErrorTrap = "TwentyNine"
        If lintLenOfFileName > 0 Then
            lstrCheckFile = lstrDestinationPath & Left$(gstrStatic.strShortReportsTestingDBFile, lintLenOfFileName - 3) & "chk"
        End If
        lstrErrorTrap = "Thirty"
        lstrDestinationFile = lstrDestinationPath & gstrStatic.strShortReportsTestingDBFile
        lstrErrorTrap = "ThirtyOne"
    Else
        lstrErrorTrap = "ThirtyTwo"
        lstrSourceFile = lstrSourcePath & gstrStatic.strShortReportsDBFile
        lstrErrorTrap = "ThirtyThree"
        lintLenOfFileName = Len(gstrStatic.strShortReportsDBFile)
        lstrErrorTrap = "ThirtyFour"
        If lintLenOfFileName > 0 Then
            lstrCheckFile = lstrDestinationPath & Left$(gstrStatic.strShortReportsDBFile, lintLenOfFileName - 3) & "chk"
        End If
        lstrErrorTrap = "ThirtyFive"
        lstrDestinationFile = lstrDestinationPath & gstrStatic.strShortReportsDBFile
        lstrErrorTrap = "ThirtySix"
    End If

    If Trim$(Dir(lstrCheckFile)) = "" Then
        FileCopy lstrSourceFile, lstrCheckFile
    End If
    
    If UCase$(Right$(Trim$(gstrStatic.strShortReportsDBFile), 3)) = "MDB" Then
        'Get reps.mdb if newer
        If FileCopyIfNewer(lstrSourceFile, lstrCheckFile) Then
            lstrErrorTrap = "ThirtySeven"
            'Always overwite localdb if new
            FileCopy lstrCheckFile, lstrDestinationFile
            lstrErrorTrap = "ThirtyEight"
        ElseIf DateValue(Date) > DateValue(FileDateTime(lstrDestinationFile)) Then
            lstrErrorTrap = "ThirtyNine"
            'Overwite localdb once per day
            FileCopy lstrCheckFile, lstrDestinationFile
            lstrErrorTrap = "Forty"
        End If
    End If
    
    Dim lbooParamaterUnknown As Boolean

    'Setup paths for each program supplied from static.ldr
    If UCase(Dir(Trim$(App.Path) & "\" & gconstrStaticLdr, vbNormal)) = UCase(gconstrStaticLdr) Then
        lintProgCount = UBound(gstrStatic.strPrograms)
        lstrErrorTrap = "FortyOne"
        For lintArrInc = 0 To lintProgCount
            lstrErrorTrap = "FortyTwo"
            lstrExeName = gstrStatic.strPrograms(lintArrInc).strProgram
            lstrErrorTrap = "FortyThree"
            lstrParam = "/" & gstrStatic.strPrograms(lintArrInc).strParam
            lstrErrorTrap = "FortyFour"
            If InStr(UCase(Command$), lstrParam) > 0 Then
                lbooParamaterUnknown = False
                Exit For
            Else
                lbooParamaterUnknown = True
            End If
            lstrErrorTrap = "FortyFive"
        Next lintArrInc
    End If
    
    'Stop loader from running an unknown program
    If lbooParamaterUnknown = True Then
        MsgBox "Your Loader Program has not been Setup correctly!", vbInformation, UCase(Command$)
        End
    End If
            
    lstrSourceFile = lstrSourcePath & lstrExeName
    lstrDestinationFile = lstrDestinationPath & lstrEFNTA & lstrExeName
    
    lstrErrorTrap = "FortySix"
    
    lbooCopyDone = FileCopyIfNewer(lstrSourceFile, lstrDestinationFile)
    lstrErrorTrap = "FortySeven"
    lblStatus.Caption = "Running."
    DoEvents
    
    'Debug.Print lstrDestinationFile & lstrPassParam
    'llngShellID = Shell(lstrDestinationFile & lstrPassParam, vbNormalFocus)
    
    ShellExecute Me.hwnd, "open", lstrDestinationFile, lstrPassParam, "C:\", 1

    lstrErrorTrap = "FortyEight"
    
    Me.Hide
    'If llngShellID <> 0 Then
        lstrErrorTrap = "FortyNine"
    '    AppActivate llngShellID
    'End If
    lstrErrorTrap = "Fifty"
    
NormalExit:
    Unload Me
    End
    Exit Sub
    
LoadErrors:
    If Err = 68 Then
        MsgBox "Could not find " & lstrSourceFile & vbCrLf & lstrErrorTrap, vbCritical
        Resume Next
    ElseIf Err = 70 Then
        'permission denied as open
        Resume Next
    End If
    If Err = 6 Then
        'ignore overflow because of long returned from shell
        Resume Next
    End If
    MsgBox "Loader Error " & Err & ": " & Err.Description & vbCrLf & lstrErrorTrap, vbCritical
    Resume NormalExit

End Sub

Sub InstallOCXOld()

Dim lstrWindowsPath As String
Dim lstrSourcePath As String
Dim llngReturnValue As Long
Dim llngPathBufferSize As Long
Dim lstrPathBuff As String
Dim lstrFileName As Variant
Dim lvarRetval As Variant
Dim lstrFileArray() As String
Dim lintArrInc As Integer

lstrWindowsPath = ""
lstrSourcePath = ""
lstrFileName = ""

On Error GoTo ErrHandler
llngPathBufferSize = 255
lstrPathBuff = Space$(llngPathBufferSize)

' call system path routine to get system directory
' from windows API

llngReturnValue = GetSystemDirectory(lstrPathBuff, llngPathBufferSize)
   
If llngReturnValue = 0 Then
    Exit Sub
End If

'return path after stripping terminator
lstrWindowsPath = Left$(lstrPathBuff, (Len(Trim$(lstrPathBuff)) - 1))
    
If lstrWindowsPath = "" Then
    'lstrwindows path will be "" if an error occured during reading!
    Exit Sub
End If

lstrSourcePath = "\\SERVER\TMOS\SETUP\SUPPORT\"

lstrFileName = Dir(lstrSourcePath & "*.OCX", vbNormal)
ReDim lstrFileArray(0)
lintArrInc = 0
Do While lstrFileName <> ""
    lstrFileName = Dir
    ReDim Preserve lstrFileArray(lintArrInc)
    lintArrInc = lintArrInc + 1
    lstrFileArray(UBound(lstrFileArray)) = lstrFileName
Loop

For lintArrInc = 0 To UBound(lstrFileArray)
    If Trim$(Dir$(lstrWindowsPath & "\" & lstrFileArray(lintArrInc))) = "" Then
        FileCopy (lstrSourcePath & lstrFileArray(lintArrInc)), (lstrWindowsPath & "\" & lstrFileArray(lintArrInc))
        FileCopy (lstrSourcePath & "regsvr32.exe"), (lstrWindowsPath & "\" & "regsvr32.exe")
        lvarRetval = Shell(lstrWindowsPath & "\" & "regsvr32 /s " & (lstrWindowsPath & "\" & lstrFileArray(lintArrInc)), 1)
        Debug.Print lstrFileArray(lintArrInc)
    End If
Next lintArrInc

lstrFileName = Dir(lstrSourcePath & "*.DLL", vbNormal)
ReDim lstrFileArray(0)
lintArrInc = 0
Do While lstrFileName <> ""
    lstrFileName = Dir
    ReDim Preserve lstrFileArray(lintArrInc)
    lintArrInc = lintArrInc + 1
    lstrFileArray(UBound(lstrFileArray)) = lstrFileName
Loop


For lintArrInc = 0 To UBound(lstrFileArray)
    If Left$(lstrFileArray(lintArrInc), 3) <> "DAO" Then
        If Trim$(Dir$(lstrWindowsPath & "\" & lstrFileArray(lintArrInc))) = "" Then
            FileCopy (lstrSourcePath & lstrFileArray(lintArrInc)), (lstrWindowsPath & "\" & lstrFileArray(lintArrInc))
            FileCopy (lstrSourcePath & "regsvr32.exe"), (lstrWindowsPath & "\" & "regsvr32.exe")
            lvarRetval = Shell(lstrWindowsPath & "\" & "regsvr32 " & (lstrWindowsPath & "\" & lstrFileArray(lintArrInc)), 1)
            Debug.Print lstrFileArray(lintArrInc)
        End If
    
    End If
Next lintArrInc

Exit Sub

ErrHandler:
    Exit Sub

End Sub

Function CreateUpdateVBS() As String
Dim lintFileNum As Integer
Dim lstrFileName As String
    
    lstrFileName = GetTempDir & "~" & Format(Now(), "MMDDSSN") & ".vbs"
    lintFileNum = FreeFile
    
    Open lstrFileName For Output As lintFileNum

    Print #lintFileNum, "ldatNow = Now()"
    Print #lintFileNum, "Do While DateAdd(" & Chr(34) & "s" & Chr(34) & ", 3, ldatNow) > Now()"
    Print #lintFileNum, "'"
    Print #lintFileNum, "Loop"
    Print #lintFileNum, ""
    Print #lintFileNum, "msgbox " & Chr(34) & "Your Loader program in out of date!" & Chr(34) & " & vbcrlf & _"
    Print #lintFileNum, "   " & Chr(34) & "Please Click OK to update, then " & Chr(34) & " & vbcrlf & _"
    Print #lintFileNum, "   " & Chr(34) & "re-run your program!" & Chr(34) & ",," & Chr(34) & "Loader Update Needed!" & Chr(34) & ""
    Print #lintFileNum, "   "
    Print #lintFileNum, "Set fso = CreateObject(" & Chr(34) & "Scripting.FileSystemObject" & Chr(34) & ")"
    Print #lintFileNum, "Set a = fso.GetFile(" & Chr(34) & gstrSourcePath & "loader.exe" & Chr(34) & ")"
    Print #lintFileNum, "a.Copy (" & Chr(34) & App.Path & "\" & "loader.exe" & Chr(34) & ")"
    Close #lintFileNum
    
    CreateUpdateVBS = lstrFileName
    
End Function

Sub Log(pstrString As String)
Dim lintFileNum As Integer
Dim lstrFileName As String

    On Error Resume Next
    
    If gbooErrorFound = True Then
        pstrString = pstrString & " Not Found!"
        gbooErrorFound = False
    Else
        pstrString = pstrString & " OK!"
    End If
    lintFileNum = FreeFile
    
    Open App.Path & "\Loader.log" For Append As lintFileNum
    Print #lintFileNum, pstrString
    Close #lintFileNum
End Sub
Sub InstallOCX(pstrSupportPath As String)
Dim lstrErrorStage As String
Dim lstrWindowsPath As String
Dim lstrWindowsSystemWow4Path As String
Dim llngReturnValue As Long
Dim llngPathBufferSize As Long
Dim lstrPathBuff As String
Dim lstrFileName As Variant
Dim lvarRetval As Variant
Dim lstrFileArray() As String
Dim lintArrInc As Integer
Dim lstrSystemFolder As String

    lstrWindowsPath = ""
    lstrWindowsSystemWow4Path = ""
    lstrFileName = ""
    
    On Error GoTo ErrHandler

    llngPathBufferSize = 255
    lstrPathBuff = Space$(llngPathBufferSize)
    
    llngReturnValue = GetSystemDirectory(lstrPathBuff, llngPathBufferSize)
       
    If llngReturnValue = 0 Then
        Exit Sub
    End If
    
    lstrWindowsPath = Left$(lstrPathBuff, (Len(Trim$(lstrPathBuff)) - 1))
    lstrWindowsSystemWow4Path = Replace(lstrWindowsPath, "System32", "SysWow64")
    lstrWindowsSystemWow4Path = Replace(lstrWindowsPath, "system32", "SysWow64")
    
    If lstrWindowsPath = "" Then
        Exit Sub
    End If
    
    lstrErrorStage = "One"
    If Trim$(pstrSupportPath) = "" Then
        Exit Sub
    End If
    Log ".Install OCX - Loader options .. " & pstrSupportPath
    'OCX's in Windows directory
    lstrFileName = Dir(pstrSupportPath & "*.OCX", vbNormal)

    lstrErrorStage = "Two"
    ReDim lstrFileArray(0)
    lintArrInc = 0
    Do While lstrFileName <> ""
        ReDim Preserve lstrFileArray(lintArrInc)
        lintArrInc = lintArrInc + 1
        lstrFileArray(UBound(lstrFileArray)) = lstrFileName
        lstrFileName = Dir
        lstrErrorStage = "Three"
    Loop
    
    For lintArrInc = 0 To UBound(lstrFileArray)
        lstrErrorStage = "Four"
        
        lstrSystemFolder = lstrWindowsSystemWow4Path & "\"
        
        If Trim$(Dir$(lstrSystemFolder & lstrFileArray(lintArrInc))) = "" Then
            lstrSystemFolder = lstrWindowsPath & "\"
        End If
        
        Log "..OCX - Win Dir - Check Time Stamp - Source - " & pstrSupportPath & _
            lstrFileArray(lintArrInc) & " (" & FileDateTime(pstrSupportPath & lstrFileArray(lintArrInc)) & ")"
            
        Log "..OCX - Win Dir - Check Time Stamp - Destination - " & lstrSystemFolder & _
            lstrFileArray(lintArrInc) & " (" & FileDateTime(lstrSystemFolder & lstrFileArray(lintArrInc)) & ")"
            
        If FileDateTime(pstrSupportPath & lstrFileArray(lintArrInc)) > FileDateTime(lstrSystemFolder & lstrFileArray(lintArrInc)) Then
            lstrErrorStage = "Five"
            FileCopy (pstrSupportPath & "regsvr32.exe"), (lstrSystemFolder & "regsvr32.exe")
            Log "..OCX - Win Dir - Copying Regsvr32.."
            lvarRetval = Shell(lstrSystemFolder & "regsvr32 /u /s " & (lstrSystemFolder & lstrFileArray(lintArrInc)), 1)
            Log "..OCX - Win Dir - Shelling Regsvr32 Unregister - " & lstrSystemFolder & lstrFileArray(lintArrInc)
            lstrErrorStage = "Six"
            SetAttr lstrSystemFolder & lstrFileArray(lintArrInc), vbNormal
            Log "..OCX - Win Dir - SetAttr to Normal - " & lstrSystemFolder & lstrFileArray(lintArrInc)
            lstrErrorStage = "Seven"
            FileCopy (pstrSupportPath & lstrFileArray(lintArrInc)), (lstrSystemFolder & lstrFileArray(lintArrInc))
            
            Log "..OCX - Win Dir - Copy - Source - " & pstrSupportPath & _
            lstrFileArray(lintArrInc) & " (" & FileDateTime(pstrSupportPath & lstrFileArray(lintArrInc)) & ")"

            Log "..OCX - Win Dir - Copy - Destination - " & lstrSystemFolder & _
                lstrFileArray(lintArrInc) & " (" & FileDateTime(lstrSystemFolder & lstrFileArray(lintArrInc)) & ")"

            lvarRetval = Shell(lstrSystemFolder & "regsvr32 /s " & (lstrSystemFolder & lstrFileArray(lintArrInc)), 1)
            Log "..OCX - Win Dir - Shelling Regsvr32 Register - " & lstrSystemFolder & lstrFileArray(lintArrInc)
            
            Debug.Print lstrFileArray(lintArrInc)
        End If
    Next lintArrInc
    lstrErrorStage = "Eight"
    
    lstrFileName = Dir(pstrSupportPath & "*.DLL", vbNormal)
    ReDim lstrFileArray(0)
    lintArrInc = 0
    lstrErrorStage = "Sixteen"
    Do While lstrFileName <> ""
        ReDim Preserve lstrFileArray(lintArrInc)
        lintArrInc = lintArrInc + 1
        lstrFileArray(UBound(lstrFileArray)) = lstrFileName
        lstrErrorStage = "Seventeen"
        lstrFileName = Dir
    Loop
    
    For lintArrInc = 0 To UBound(lstrFileArray)
        If Left$(lstrFileArray(lintArrInc), 3) <> "DAO" Then
            
            lstrSystemFolder = lstrWindowsSystemWow4Path & "\"
            
            If Trim$(Dir$(lstrSystemFolder & lstrFileArray(lintArrInc))) = "" Then
                lstrSystemFolder = lstrWindowsPath & "\"
            End If
        
            If FileDateTime(pstrSupportPath & lstrFileArray(lintArrInc)) > FileDateTime(lstrSystemFolder & lstrFileArray(lintArrInc)) Then
                FileCopy (pstrSupportPath & "regsvr32.exe"), (lstrSystemFolder & "regsvr32.exe")
                lvarRetval = Shell(lstrSystemFolder & "regsvr32 /u /s " & (lstrSystemFolder & lstrFileArray(lintArrInc)), 1)
                SetAttr lstrSystemFolder & lstrFileArray(lintArrInc), vbNormal
                lstrErrorStage = "Eighteen"
                FileCopy (pstrSupportPath & lstrFileArray(lintArrInc)), (lstrSystemFolder & lstrFileArray(lintArrInc))
                lvarRetval = Shell(lstrSystemFolder & "regsvr32 /s " & (lstrSystemFolder & lstrFileArray(lintArrInc)), 1)
                Debug.Print lstrFileArray(lintArrInc)
                lstrErrorStage = "Nineteen"
            End If
        
        End If
    Next lintArrInc
    
    Exit Sub

ErrHandler:
    gbooErrorFound = True
    Resume Next

End Sub

Function XStaticInErrorCheck() As Boolean
Dim lintProgCount As Integer
Dim lintArrInc As Integer
Dim lstrErrMsg As String
Dim lstrErrorPosition As String

    On Error GoTo ErrHandler
    lstrErrorPosition = "One"
    If InStr(UCase(Command$), "/TEST") > 0 Then
        lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "SysFileInfo", "SrvTestPth")
        lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "SysFileInfo", "SupTestPth")
        lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "DB", "LocalTest")
        lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "DB", "RepsTest")
    Else
        lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "SysFileInfo", "ServerPath")
        lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "SysFileInfo", "SuppPath")
        lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "DB", "Local")
        lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "DB", "Reps")
    End If

    lstrErrorPosition = "Two"
    lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "SysFileInfo", "AppPath")
    
    lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "Programs", "ProgCount")
    lintProgCount = CInt(Val(GetPrivateINI(Trim$(App.Path) & "\" & gconstrStaticIni, "Programs", "ProgCount")))
    
    lstrErrorPosition = "Three"
    For lintArrInc = 1 To lintProgCount
        lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "Programs", "Prog" & lintArrInc)
        lstrErrMsg = lstrErrMsg & CheckPrivateIniForError(Trim$(App.Path) & "\" & gconstrStaticIni, "Programs", "Prog" & lintArrInc & "Param")
    Next lintArrInc
    
    lstrErrorPosition = "Four"

    If lstrErrMsg = "" Then
        XStaticInErrorCheck = False
    Else
        lstrErrMsg = "Please contact your IT Techical support office!" & vbCrLf & _
            "The necessary information in your STATIC.INI file is not present!" & _
            vbCrLf & vbCrLf & lstrErrMsg & vbCrLf & "Please quote the above message!!"
        MsgBox lstrErrMsg, vbCritical, "Loader Error!"
        XStaticInErrorCheck = True
    End If
    
    Exit Function
ErrHandler:
    MsgBox Err.Number & " " & Err.Description & vbCrLf & "Error Position = " & lstrErrorPosition & " from function StaticErrorCheck"
    
End Function

Function CheckPrivateIniForError(ByVal PrivateINI As String, ByVal AppName As String, ByVal keyword As String) As String
Dim lstrReturnValue As String

    lstrReturnValue = GetPrivateINI(PrivateINI, AppName, keyword)
    
    If Trim$(lstrReturnValue) = "" Then
        CheckPrivateIniForError = Spacer("Section [" & AppName & "] ", 25) & vbTab & Spacer(" Keyword = " & keyword, 25) & vbTab & " is missing!" & vbCrLf
    Else
        CheckPrivateIniForError = ""
    End If
End Function

Function Spacer(pstrString, pintLength As Integer, Optional pstrParam As String) As String
Dim lintWordLength As Integer

    If IsMissing(pstrParam) Then
        pstrParam = ""
    End If
    
    If Len(pstrString) > pintLength Then
        pstrString = Left$(pstrString, pintLength)
    End If
    lintWordLength = Len(Trim$(pstrString))
    Select Case pstrParam
    Case "L"
        Spacer = Space(pintLength - lintWordLength) & Trim$(pstrString)
    Case Else
        Spacer = Trim$(pstrString) & Space(pintLength - lintWordLength)
    End Select

End Function
