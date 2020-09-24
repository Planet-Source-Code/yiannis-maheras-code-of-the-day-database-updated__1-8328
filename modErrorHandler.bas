Attribute VB_Name = "modErrorHandler"
Dim CallStackSize As Integer
Dim CallStack() As String
Global gstErrorFlag As String
Public Function DirExists(ByVal strDirName As String) As Integer
1000    Call CallStackPush("frmError", "DirExists") '*Y-CA*

   Dim strDummy As String

   On Error Resume Next

1010    If Right$(strDirName, 1) <> "\" Then
1020       strDirName = strDirName & "\"
1030    End If

1040    strDummy = Dir$(strDirName & "*.*", vbDirectory)
1050    DirExists = Not (strDummy = vbNullString)

1060    Err = 0

1070    Call CallStackPop '*Y-CA*
End Function


Public Sub WriteError(ModName As String, ProcName As String, sErrorLineNo As String)
1000    Call CallStackPush("frmError", "WriteError") '*Y-CA*
   On Error GoTo ErrorHandler ' *Y-CA*

   Dim sErrorNumber As String
   Dim sErrorDescription As String
   Dim sErrFile As String
   Dim i As Integer
   Dim sMsg$, starSeparator As String
   Dim ErrFile As String
   
   '*use to prevent recording multiple errors in loop
1010    Static LastErrorRecorded As String

   ' save current values of error object properties
1020    sErrorDescription = CStr(Err.Description)
1030    sErrorNumber = CStr(Err.Number)
    
1040    If sErrorNumber & sErrorDescription & ModName & ProcName = LastErrorRecorded Then
1050    Call CallStackPop '*Y-CA*
1060       Exit Sub
1070    Else
      ' make sure the folder for error loggin exists
1080       If Not DirExists(App.Path + "\ErrorFolder") Then
1090          MkDir App.Path + "\ErrorFolder"
1100       End If

      ' include the current date into the errorlog file name
1110       ErrFile = App.Path + "\ErrorFolder\ErrorLog_" & Format(Now, "yyyy_mm_dd") & ".log"
      ' make sure the errorlog file exists and open/create the file
1120       FileExists = Dir$(ErrFile)
1130       If Len(FileExists) > 0 Then
1140          fl% = FreeFile
1150          Open ErrFile For Append As fl%
1160       Else
1170          fl% = FreeFile
1180          Open ErrFile For Output As fl%
1190       End If
   
1200       starSeparator = String(70, "*")

      ' write error message into the text file
1210       Print #fl%, starSeparator

      ' error source procedure name
1220       Print #fl%, "* Source: " & ModName & "(" & ProcName & ")"

      ' define procedure section containing the error
1230       Print #fl%, "* Error Number: " & sErrorNumber
1240       Print #fl%, "* Error Line Number: " & sErrorLineNo
1250       Print #fl%, "* Description:"

      ' save sErrorDescription string in predefined format
1260       Print #fl%, sErrorDescription

      ' call cascade description
1270       Print #fl%, "* Error Call History: "
1280       For i = 0 To CallStackSize - 1
1290          Print #fl%, "*    - " & CallStack(i)
1300       Next i

      ' put the time stamp
1310       Print #fl%, "* Date/Time: " & Now
1320       Print #fl%, starSeparator
1330       Close #fl%

1340       LastErrorRecorded = sErrorNumber & sErrorDescription & ModName & ProcName
1350    End If

1360 Exit Sub ' *Y-CA*
' *Y-CA*
ErrorHandler:    ' *Y-CA*
   Call ErrMsg("frmError", "WriteError", Err, Erl()) ' *Y-CA*
   Select Case gstErrorFlag ' *Y-CA*
      Case "CANCEL" ' *Y-CA*
         Exit Sub ' *Y-CA*
      Case "RETRY" ' *Y-CA*
         Resume ' *Y-CA*
      Case "IGNORE" ' *Y-CA*
         Resume Next ' *Y-CA*
   End Select ' *Y-CA*
   ' *Y-CA*
1370    Call CallStackPop '*Y-CA*
End Sub

' add next element into stack using collection object
Public Sub CallStackPush(sModule As String, sProcedure As String)
1000    Call CallStackPush("frmError", "CallStackPush") '*Y-CA*
   On Error GoTo ErrorHandler ' *Y-CA*

1010    If CallStackSize = 0 Then ReDim CallStack(0)
1020    ReDim Preserve CallStack(CallStackSize)
1030    CallStack(CallStackSize) = sModule & "(" & sProcedure & ")"
1040    CallStackSize = CallStackSize + 1

1050    Call CallStackPop '*Y-CA*
1060 Exit Sub ' *Y-CA*
' *Y-CA*
ErrorHandler:    ' *Y-CA*
   Call ErrMsg("frmError", "CallStackPush", Err, Erl()) ' *Y-CA*
   Select Case gstErrorFlag ' *Y-CA*
      Case "CANCEL" ' *Y-CA*
         Exit Sub ' *Y-CA*
      Case "RETRY" ' *Y-CA*
         Resume ' *Y-CA*
      Case "IGNORE" ' *Y-CA*
         Resume Next ' *Y-CA*
   End Select ' *Y-CA*
   ' *Y-CA*
1070    Call CallStackPop '*Y-CA*
End Sub
' remove last element out of stack using collection object
Public Sub CallStackPop()
1000    Call CallStackPush("frmError", "CallStackPop") '*Y-CA*
   On Error GoTo ErrorHandler ' *Y-CA*

1010    CallStackSize = CallStackSize - 1
1020    If CallStackSize = 0 Then
1030       ReDim CallStack(0)
1040    Else
1050       ReDim Preserve CallStack(CallStackSize - 1)
1060    End If

1070    Call CallStackPop '*Y-CA*
1080 Exit Sub ' *Y-CA*
' *Y-CA*
ErrorHandler:    ' *Y-CA*
   Call ErrMsg("frmError", "CallStackPop", Err, Erl()) ' *Y-CA*
   Select Case gstErrorFlag ' *Y-CA*
      Case "CANCEL" ' *Y-CA*
         Exit Sub ' *Y-CA*
      Case "RETRY" ' *Y-CA*
         Resume ' *Y-CA*
      Case "IGNORE" ' *Y-CA*
         Resume Next ' *Y-CA*
   End Select ' *Y-CA*
   ' *Y-CA*
1090    Call CallStackPop '*Y-CA*
End Sub



Sub ErrMsg(FormName As String, Rout As String, ErrNo As Integer, ErrLine As Integer)
1000    Call CallStackPush("frmError", "ErrMsg") '*Y-CA*
   On Error GoTo ErrorHandler ' *Y-CA*
   
1010    Call WriteError(FromName, Rout, CStr(ErrLine))
1020    Load frmError
   
1030    frmError!cFormName.Text = FormName
1040    frmError!cRoutine.Text = Rout
1050    frmError!cErrNo.Text = ErrNo
1060    frmError!cErrLine.Text = ErrLine
1070    frmError!cErrDes.Text = Error(ErrNo)
   
1080    If Screen.MousePointer <> 0 Then
1090       MP = Screen.MousePointer
1100       Screen.MousePointer = 0
1110    End If
   
1120    frmError.Show 1
   
1130    DoEvents

1140    If MP <> 0 Then Screen.MousePointer = MP

1150    Call CallStackPop '*Y-CA*
1160 Exit Sub ' *Y-CA*
' *Y-CA*
ErrorHandler:    ' *Y-CA*
   Call ErrMsg("frmError", "ErrMsg", Err, Erl()) ' *Y-CA*
   Select Case gstErrorFlag ' *Y-CA*
      Case "CANCEL" ' *Y-CA*
         Exit Sub ' *Y-CA*
      Case "RETRY" ' *Y-CA*
         Resume ' *Y-CA*
      Case "IGNORE" ' *Y-CA*
         Resume Next ' *Y-CA*
   End Select ' *Y-CA*
   ' *Y-CA*
1170    Call CallStackPop '*Y-CA*
End Sub

