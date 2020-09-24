Attribute VB_Name = "GlobalModule"
Public gobjOutlook As Outlook.Application
Public gobjNamespace As Outlook.NameSpace

Global gCurrentDB As Database

Global gstDBFolder As String
Global gstDBName As String
Global gstPOP3 As String
Global gstSMTP As String
Global gstLoginName As String
Global gstPassword As String
Global gstMailAddress As String
Global gstYourName As String
Global gstMessageRules As String

Dim StartRow As Integer
Dim EndRow As Integer

Function AllReadyImported() As Boolean
   
   On Error GoTo ErrorHandler
   Call CallStackPush("GlobalModule", "AllReadyImported")

   Dim SQL As String
   Dim Rs As Recordset
   
   AllReadyImported = False
   
   SQL = "SELECT MailsTable.MailSubject, MailsTable.MailDate From MailsTable"
   SQL = SQL + " WHERE (((MailsTable.MailSubject)='" & frmImportControls!cMailSubject.Text & "') AND ((MailsTable.MailDate)='" & frmImportControls!cMailDate.Text & "'));"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   If GetNumbRecsRs(Rs) > 0 Then AllReadyImported = True
   
   Rs.Close
   Set Rs = Nothing

   Call CallStackPop

Exit Function
 
ErrorHandler:
   Call ErrMsg("GlobalModule", "AllReadyImported", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Function
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select
   
End Function

Function CreateOutlookInstance() As Boolean
   
   On Error GoTo ErrorHandler
   Call CallStackPush("GlobalModule", "CreateOutlookInstance")

   'Create application and namespace root objects
   Set gobjOutlook = CreateObject("Outlook.Application.9")
   If Err Then
      MsgBox "Could not create Outlook Application object!", vbCritical
      CreateOutlookInstance = False
      Call CallStackPop
      Exit Function
   End If

   Set gobjNamespace = gobjOutlook.GetNamespace("MAPI")
   If Err Then
      MsgBox "Could not create MAPI Namespace!", vbCritical
      CreateOutlookInstance = False
      Call CallStackPop
      Exit Function
   End If

   CreateOutlookInstance = True

   Call CallStackPop

Exit Function
 
ErrorHandler:
   Call ErrMsg("GlobalModule", "CreateOutlookInstance", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Function
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select
   
End Function



Public Function Encode(vText As String)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("GlobalModule", "Encode")

   Dim CurSpc As Integer
   Dim varLen As Integer
   Dim varChr As String
   Dim varFin As String

   varLen = Len(vText)

   Do While CurSpc <= varLen

      DoEvents
      CurSpc = CurSpc + 1
      varChr = Mid(vText, CurSpc, 1)

      Select Case varChr
         'lower case
         Case "a"
            varChr = "coe"
         Case "b"
            varChr = "wer"
         Case "c"
            varChr = "ibq"
         Case "d"
            varChr = "am7"
         Case "e"
            varChr = "pm1"
         Case "f"
            varChr = "mop"
         Case "g"
            varChr = "9v4"
         Case "h"
            varChr = "qu6"
         Case "i"
            varChr = "zxc"
         Case "j"
            varChr = "4mp"
         Case "k"
            varChr = "f88"
         Case "l"
            varChr = "qe2"
         Case "m"
            varChr = "vbn"
         Case "n"
            varChr = "qwt"
         Case "o"
            varChr = "pl5"
         Case "p"
            varChr = "13s"
         Case "q"
            varChr = "c%l"
         Case "r"
            varChr = "w$w"
         Case "s"
            varChr = "6a@"
         Case "t"
            varChr = "!2&"
         Case "u"
            varChr = "(=c"
         Case "v"
            varChr = "wvf"
         Case "w"
            varChr = "dp0"
         Case "x"
            varChr = "w$-"
         Case "y"
            varChr = "vn&"
         Case "z"
            varChr = "c*4"

            'numbers
         Case "1"
            varChr = "aq@"
         Case "2"
            varChr = "902"
         Case "3"
            varChr = "2.&"
         Case "4"
            varChr = "/w!"
         Case "5"
            varChr = "|pq"
         Case "6"
            varChr = "ml|"
         Case "7"
            varChr = "t'?"
         Case "8"
            varChr = ">^s"
         Case "9"
            varChr = "<s^"
         Case "0"
            varChr = ";&c"

            'caps
         Case "A"
            varChr = "$)c"
         Case "B"
            varChr = "-gt"
         Case "C"
            varChr = "|p*"
         Case "D"
            varChr = "1" & Chr(34) & "r"
         Case "E"
            varChr = "c>:"
         Case "F"
            varChr = "@+x"
         Case "G"
            varChr = "v^a"
         Case "H"
            varChr = "]eE"
         Case "I"
            varChr = "aP0"
         Case "J"
            varChr = "{=1"
         Case "K"
            varChr = "cWv"
         Case "L"
            varChr = "cDc"
         Case "M"
            varChr = "*,!"
         Case "N"
            varChr = "fW" & Chr(34)
         Case "O"
            varChr = ".?T"
         Case "P"
            varChr = "%<8"
         Case "Q"
            varChr = "@:a"
         Case "R"
            varChr = "&c$"
         Case "S"
            varChr = "WnY"
         Case "T"
            varChr = "{Sh"
         Case "U"
            varChr = "_%M"
         Case "V"
            varChr = "}'$"
         Case "W"
            varChr = "QlU"
         Case "X"
            varChr = "Im^"
         Case "Y"
            varChr = "l|P"
         Case "Z"
            varChr = ".>#"
            'Special characters
         Case "!"
            varChr = "\" & Chr(34) & "]"
         Case "@"
            varChr = "cY,"
         Case "#"
            varChr = "x%B"
         Case "$"
            varChr = "a*v"
         Case "%"
            varChr = "'&T"
         Case "^"
            varChr = ";%R"
         Case "&"
            varChr = "eG_"
         Case "*"
            varChr = "Z/e"
         Case "("
            varChr = "rG\"
         Case ")"
            varChr = "]*F"
         Case "_"
            varChr = "@B*"
         Case "-"
            varChr = "+Hc"
         Case "="
            varChr = "&|D"
         Case "+"
            varChr = "(:#"
         Case "["
            varChr = "SlW"
         Case "]"
            varChr = "'QB"
         Case "{"
            varChr = "{D>"
         Case "}"
            varChr = "+c%"
         Case ":"
            varChr = "(s:"
         Case ";"
            varChr = "^a("
         Case "'"
            varChr = "16."
         Case Chr(34)
            varChr = "s.*"
         Case ","
            varChr = "&?W"
         Case "."
            varChr = "GPQ"
         Case "<"
            varChr = "SK*"
         Case ">"
            varChr = "RL^"
         Case "/"
            varChr = "40C"
         Case "?"
            varChr = "?#9"
         Case "\"
            varChr = "_?/"
         Case "|"
            varChr = "(_@"
         Case " "
            varChr = "=#B"
      End Select
      varFin = varFin & varChr

      DoEvents
   Loop

   Encode = varFin

   Call CallStackPop

Exit Function
 
ErrorHandler:
   Call ErrMsg("GlobalModule", "Encode", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Function
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select
   
End Function
Public Function DeCode(vText As String)
   
   On Error GoTo ErrorHandler
   Call CallStackPush("GlobalModule", "DeCode")

   Dim CurSpc As Integer
   Dim varLen As Integer
   Dim varChr As String
   Dim varFin As String
   
   CurSpc = CurSpc + 1
   varLen = Len(vText)

   Do While CurSpc <= varLen
      DoEvents
      varChr = Mid(vText, CurSpc, 3)
      Select Case varChr
         'lower case
         Case "coe"
            varChr = "a"
         Case "wer"
            varChr = "b"
         Case "ibq"
            varChr = "c"
         Case "am7"
            varChr = "d"
         Case "pm1"
            varChr = "e"
         Case "mop"
            varChr = "f"
         Case "9v4"
            varChr = "g"
         Case "qu6"
            varChr = "h"
         Case "zxc"
            varChr = "i"
         Case "4mp"
            varChr = "j"
         Case "f88"
            varChr = "k"
         Case "qe2"
            varChr = "l"
         Case "vbn"
            varChr = "m"
         Case "qwt"
            varChr = "n"
         Case "pl5"
            varChr = "o"
         Case "13s"
            varChr = "p"
         Case "c%l"
            varChr = "q"
         Case "w$w"
            varChr = "r"
         Case "6a@"
            varChr = "s"
         Case "!2&"
            varChr = "t"
         Case "(=c"
            varChr = "u"
         Case "wvf"
            varChr = "v"
         Case "dp0"
            varChr = "w"
         Case "w$-"
            varChr = "x"
         Case "vn&"
            varChr = "y"
         Case "c*4"
            varChr = "z"

            'numbers
         Case "aq@"
            varChr = "1"
         Case "902"
            varChr = "2"
         Case "2.&"
            varChr = "3"
         Case "/w!"
            varChr = "4"
         Case "|pq"
            varChr = "5"
         Case "ml|"
            varChr = "6"
         Case "t'?"
            varChr = "7"
         Case ">^s"
            varChr = "8"
         Case "<s^"
            varChr = "9"
         Case ";&c"
            varChr = "0"

            'caps
         Case "$)c"
            varChr = "A"
         Case "-gt"
            varChr = "B"
         Case "|p*"
            varChr = "C"
         Case "1" & Chr(34) & "r"
            varChr = "D"
         Case "c>:"
            varChr = "E"
         Case "@+x"
            varChr = "F"
         Case "v^a"
            varChr = "G"
         Case "]eE"
            varChr = "H"
         Case "aP0"
            varChr = "I"
         Case "{=1"
            varChr = "J"
         Case "cWv"
            varChr = "K"
         Case "cDc"
            varChr = "L"
         Case "*,!"
            varChr = "M"
         Case "fW" & Chr(34)
            varChr = "N"
         Case ".?T"
            varChr = "O"
         Case "%<8"
            varChr = "P"
         Case "@:a"
            varChr = "Q"
         Case "&c$"
            varChr = "R"
         Case "WnY"
            varChr = "S"
         Case "{Sh"
            varChr = "T"
         Case "_%M"
            varChr = "U"
         Case "}'$"
            varChr = "V"
         Case "QlU"
            varChr = "W"
         Case "Im^"
            varChr = "X"
         Case "l|P"
            varChr = "Y"
         Case ".>#"
            varChr = "Z"
            'Special characters
         Case "\" & Chr(34) & "]"
            varChr = "!"
         Case "cY,"
            varChr = "@"
         Case "x%B"
            varChr = "#"
         Case "a*v"
            varChr = "$"
         Case "'&T"
            varChr = "%"
         Case ";%R"
            varChr = "^"
         Case "eG_"
            varChr = "&"
         Case "Z/e"
            varChr = "*"
         Case "rG\"
            varChr = "("
         Case "]*F"
            varChr = ")"
         Case "@B*"
            varChr = "_"
         Case "+Hc"
            varChr = "-"
         Case "&|D"
            varChr = "="
         Case "(:#"
            varChr = "+"
         Case "SlW"
            varChr = "["
         Case "'QB"
            varChr = "]"
         Case "{D>"
            varChr = "{"
         Case "+c%"
            varChr = "}"
         Case "(s:"
            varChr = ":"
         Case "^a("
            varChr = ";"
         Case "16."
            varChr = "'"
         Case "s.*"
            varChr = Chr(34)
         Case "&?W"
            varChr = ","
         Case "GPQ"
            varChr = "."
         Case "SK*"
            varChr = "<"
         Case "RL^"
            varChr = ">"
         Case "40C"
            varChr = "/"
         Case "?#9"
            varChr = "?"
         Case "_?/"
            varChr = "\"
         Case "(_@"
            varChr = "|"
         Case "=#B"
            varChr = " "
      End Select
      varFin = varFin & varChr
      CurSpc = CurSpc + 3
      DoEvents
   Loop
   DeCode = varFin

   Call CallStackPop

Exit Function
 
ErrorHandler:
   Call ErrMsg("GlobalModule", "DeCode", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Function
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select
   
End Function
Function GetNumbRecsRs(Rs As Recordset) As Long

   Dim rsClone As Recordset

   On Error GoTo RsErr
   Call CallStackPush("GlobalModule", "GetNumbRecsRs")
      
   Set rsClone = Rs.Clone()
   If Not rsClone.EOF Then rsClone.MoveLast
   GetNumbRecsRs = rsClone.RecordCount
   rsClone.Close
   Set rsClose = Nothing
   
   Call CallStackPop

Exit Function

RsErr:
   'just return because row count is non critical
   GetNumbRecsRs = -1
   Call CallStackPop
   Exit Function

End Function


Public Sub RefreshTreeView()

   Dim SQL As String
   Dim Rs As Recordset
   Dim CatRs As Recordset
   Dim TitleRs As Recordset
   Dim nodX As Node
   Dim i, j As Integer
   Dim FirstRoot As Boolean
   
   On Error GoTo ErrorHandler
   Call CallStackPush("GlobalModule", "RefreshTreeView")

   SQL = "SELECT DISTINCTROW Code.Category From Code GROUP BY Code.Category;"
   Set CatRs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   frmMain!TreeView1.Nodes.Clear
   Set nodX = frmMain!TreeView1.Nodes.Add(, , "Root", "Planet Source Code")
   
   If GetNumbRecsRs(CatRs) > 0 Then
      CatRs.MoveFirst: i = 1
      Do Until CatRs.EOF
         Key$ = "Cat" & i
      
         SQL = "SELECT DISTINCTROW Code.Category, Code.Title, Code.Code"
         SQL = SQL + " From Code"
         SQL = SQL + " GROUP BY Code.Category, Code.Title, Code.Code"
         SQL = SQL + " HAVING (((Code.Category)='" & CatRs(0) & "'));"
         Set TitleRs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
         
         frmSplash!cActions.Caption = CatRs(0) & ", retriving code..."
         DoEvents
         
         RNo = GetNumbRecsRs(TitleRs)
         Set nodX = frmMain!TreeView1.Nodes.Add("Root", tvwChild, Key$, CatRs(0) + "(" & RNo & ")")
         If RNo > 0 Then
            TitleRs.MoveFirst
            Do Until TitleRs.EOF
               ChildKey$ = "Code_" & TitleRs(2)
               Set nodX = frmMain!TreeView1.Nodes.Add(Key$, tvwChild, ChildKey$, TitleRs(1))
               TitleRs.MoveNext
            Loop
         End If
         
         TitleRs.Close
         Set TitleRs = Nothing
      
         CatRs.MoveNext: i = i + 1
      Loop
      frmMain!TreeView1.Nodes(Key$).EnsureVisible
   End If
      
   CatRs.Close
   Set CatRs = Nothing
   
   SQL = "SELECT * FROM Code"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   RecNo% = GetNumbRecsRs(Rs)
   
   Rs.Close
   Set Rs = Nothing
   
   SQL = "SELECT TOP 1 Code.MailDate From code GROUP BY Code.MailDate ORDER BY Code.MailDate DESC;"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenSnapshot)
   
   If GetNumbRecsRs(Rs) > 0 Then
      frmMain.Caption = "Planet Source Code Database (The DB contain " & RecNo% & " pieces of code) (Last updated at " & Rs(0) & ")"
   Else
      frmMain.Caption = "Planet Source Code Database (The DB contain " & RecNo% & " pieces of code)"
   End If
   
   Rs.Close
   Set Rs = Nothing

   Call CallStackPop

Exit Sub

ErrorHandler:
   Call ErrMsg("Global", "RefreshTreeView", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub


Sub ImportMails()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("GlobalModule", "ImportMails")

   i% = 1
   Do Until i% = frmImportControls!cMessageText.rows - 2
      MessLine$ = frmImportControls!cMessageText.TextMatrix(i%, 0)
   
      If (Left$(MessLine$, 9) = "Category:" And Left$(frmImportControls!cMessageText.TextMatrix(i% + 1, 0), 6) = "Level:") Or (Left$(MessLine$, 9) = "Category:" And Left$(frmImportControls!cMessageText.TextMatrix(i% + 1, 0), 12) = "Description:") Then
         b% = i%
         Do
            If frmImportControls!cMessageText.TextMatrix(b%, 0) = String(48, "*") Or frmImportControls!cMessageText.TextMatrix(b%, 0) = String(48, "=") Then
               Exit Do
            Else
               b% = b% - 1
            End If
         Loop
         If b% + 2 = i% Then
            StartRow = i% - 1
         Else
            Txt = frmImportControls!cMessageText.TextMatrix(b% + 1, 0)
            For k% = (b% + 2) To (i% - 1)
               Txt = Txt + " " + frmImportControls!cMessageText.TextMatrix(k%, 0)
               frmImportControls!cMessageText.RemoveItem k%
            Next k%
            StartRow = b% + 1
            frmImportControls!cMessageText.TextMatrix(StartRow, 0) = Txt
         End If
      End If
      If Left$(MessLine$, 12) = "Submitted on" Then
         EndRow = i%
      End If
      If UCase(Left$(Trim(MessLine$), 6)) = "MONDAY" Or UCase(Left$(Trim(MessLine$), 7)) = "TUESDAY" Or UCase(Left$(Trim(MessLine$), 9)) = "WEDNESDAY" Or UCase(Left$(Trim(MessLine$), 6)) = "FRIDAY" Or UCase(Left$(Trim(MessLine$), 8)) = "SATURDAY" Or UCase(Left$(Trim(MessLine$), 6)) = "SUNDAY" Or UCase(Left$(Trim(MessLine$), 8)) = "THURSDAY" Then
         b% = InStr(MessLine$, ",")
         If b% > 0 Then
            MailDate = Trim(Right$(MessLine$, Len(MessLine$) - b%))
            frmImportControls!cMailDate.Text = Right$(MailDate, 4) + "/" + Left$(MailDate, 2) + "/" + Mid$(MailDate, 4, 2)
         End If
      End If
      If StartRow > 0 And EndRow > 0 Then
         Call StripCode
         Call SaveData
         StartRow = 0: EndRow = 0
         If UCase(frmImportControls!cCodeExist.Caption) = "CANCEL" Then Exit Do
      End If
      i% = i% + 1
   Loop

   frmImportControls!cMessageText.Clear
   frmImportControls!cMessageText.rows = 2

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("GlobalModule", "ImportMails", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select
   
End Sub

Private Sub SaveData()

   Dim SQL As String
   Dim Rs As Recordset
   Static Proccess%
   
   On Error GoTo ErrorHandler
   Call CallStackPush("GlobalModule", "SaveData")
   
   SQL = "SELECT DISTINCTROW Code.Code, Code.Title, Code.Category, Code.Level, Code.Description, Code.URL, Code.Compatibility, Code.Submitted, Code.MailDate FROM Code"
   SQL = SQL + " WHERE (((Code.URL) Like '" & frmImportControls!cURL.Text & "'));"
   Set Rs = gCurrentDB.OpenRecordset(SQL, dbOpenDynaset)
   
   If GetNumbRecsRs(Rs) = 0 Then
      Rs.AddNew
      Rs(1) = frmImportControls!cTitle.Text
      Rs(2) = frmImportControls!cCategory.Text
      Rs(3) = frmImportControls!cLevel.Text
      Rs(4) = frmImportControls!cDes.Text
      Rs(5) = frmImportControls!cURL.Text
      Rs(6) = frmImportControls!cCompatibility.Text
      Rs(7) = frmImportControls!cSubmitted.Text
      Rs(8) = frmImportControls!cMailDate.Text
      Rs.Update
      frmProgress!cRecsImport.Caption = "Records imported: " & Proccess%
      DoEvents: Proccess% = Proccess% + 1
   Else
      frmImport.MousePointer = 0
      Load frmCodeExist
      GoSub FillForm
      frmCodeExist.Show 1
      frmImport.MousePointer = 11
      frmImport.Refresh
      frmMain.Refresh
      frmProgress.Refresh
      DoEvents
      Select Case frmImportControls!cCodeExist.Caption
         Case "Save"
            Rs.AddNew
            Rs(1) = frmImportControls!cTitle.Text
            Rs(2) = frmImportControls!cCategory.Text
            Rs(3) = frmImportControls!cLevel.Text
            Rs(4) = frmImportControls!cDes.Text
            Rs(5) = frmImportControls!cURL.Text
            Rs(6) = frmImportControls!cCompatibility.Text
            Rs(7) = frmImportControls!cSubmitted.Text
            Rs(8) = frmImportControls!cMailDate.Text
            Rs.Update
            frmProgress!cRecsImport.Caption = "Records imported: " & Proccess%
            DoEvents: Proccess% = Proccess% + 1
         Case "Update"
            Rs.Edit
            Rs(1) = frmImportControls!cTitle.Text
            Rs(2) = frmImportControls!cCategory.Text
            Rs(3) = frmImportControls!cLevel.Text
            Rs(4) = frmImportControls!cDes.Text
            Rs(5) = frmImportControls!cURL.Text
            Rs(6) = frmImportControls!cCompatibility.Text
            Rs(7) = frmImportControls!cSubmitted.Text
            Rs(8) = frmImportControls!cMailDate.Text
            Rs.Update
            frmProgress!cRecsImport.Caption = "Records imported: " & Proccess%
            DoEvents: Proccess% = Proccess% + 1
         Case "Ignore"
            'nothing to do
         Case "Cancel"
            'nothing to do
      End Select
   End If
   
   Rs.Close
   Set Rs = Nothing

   Call CallStackPop

Exit Sub

FillForm:
   frmCodeExist!cTitleNew.Text = frmImportControls!cTitle.Text
   frmCodeExist!cCategoryNew.Text = frmImportControls!cCategory.Text
   frmCodeExist!cLevelNew.Text = frmImportControls!cLevel.Text
   frmCodeExist!cDesNew.Text = frmImportControls!cDes.Text
   frmCodeExist!cURLNew.Text = frmImportControls!cURL.Text
   frmCodeExist!cCompatibilityNew.Text = frmImportControls!cCompatibility.Text
   frmCodeExist!cSubmittedNew.Text = frmImportControls!cSubmitted.Text
   frmCodeExist!cMailDateNew.Text = frmImportControls!cMailDate.Text

   frmCodeExist!cTitleOld.Text = Rs(1)
   frmCodeExist!cCategoryOld.Text = Rs(2)
   frmCodeExist!cLevelOld.Text = Rs(3)
   frmCodeExist!cDesOld.Text = Rs(4)
   frmCodeExist!cURLOld.Text = Rs(5)
   frmCodeExist!cCompatibilityOld.Text = Rs(6)
   frmCodeExist!cSubmittedOld.Text = Rs(7)
   frmCodeExist!cMailDateOld.Text = Rs(8)

   Return

ErrorHandler:
   Call ErrMsg("Global", "SaveData", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select

End Sub
Private Sub StripCode()
   
   On Error GoTo ErrorHandler
   Call CallStackPush("GlobalModule", "StripCode")

   For i% = StartRow To EndRow
      MessLine$ = frmImportControls!cMessageText.TextMatrix(i%, 0)
      If i% = StartRow Then
         ' Title
         Txt = Trim(MessLine$)
         A% = InStr(Txt, ")")
         If A% > 0 Then frmImportControls!cTitle.Text = Trim(Right$(Txt, Len(Txt) - A%))
      End If
      If Left$(MessLine$, 9) = "Category:" Then
         ' Category
         Txt = Trim(MessLine$)
         A% = InStr(Txt, ":")
         If A% > 0 Then frmImportControls!cCategory.Text = Trim(Right$(Txt, Len(Txt) - A%))
      End If
      If Left$(MessLine$, 6) = "Level:" Then
         ' Level
         Txt = Trim(MessLine$)
         A% = InStr(Txt, ":")
         If A% > 0 Then frmImportControls!cLevel.Text = Trim(Right$(Txt, Len(Txt) - A%))
      End If
      If Left$(MessLine$, 12) = "Description:" Then
         ' Description
         Txt = Trim(MessLine$)
         A% = InStr(Txt, ":")
         If A% > 0 Then frmImportControls!cDes.Text = Trim(Right$(Txt, Len(Txt) - A%))
         Row% = i% + 1
         Do
            If Trim(frmImportControls!cMessageText.TextMatrix(Row%, 0)) = "Complete source code is at:" Then Exit Do
            frmImportControls!cDes.Text = frmImportControls!cDes.Text + " " + Trim(frmImportControls!cMessageText.TextMatrix(Row%, 0))
            Row% = Row% + 1
         Loop
      End If
      If Trim(MessLine$) = "Complete source code is at:" Then
         ' URL
         frmImportControls!cURL.Text = Trim(frmImportControls!cMessageText.TextMatrix(i% + 1, 0))
      End If
      If Left$(MessLine$, 14) = "Compatibility:" Then
         ' Compatibility
         Txt = Trim(MessLine$)
         A% = InStr(Txt, ":")
         If A% > 0 Then frmImportControls!cCompatibility.Text = Trim(Right$(Txt, Len(Txt) - A%))
         Row% = i% + 1
         Do
            If Left$(frmImportControls!cMessageText.TextMatrix(Row%, 0), 12) = "Submitted on" Then Exit Do
            frmImportControls!cCompatibility.Text = frmImportControls!cCompatibility.Text + " " + Trim(frmImportControls!cMessageText.TextMatrix(Row%, 0))
            Row% = Row% + 1
         Loop
      End If
      If Left$(MessLine$, 12) = "Submitted on" Then
         ' Submitted on
         Txt = Trim(MessLine$)
         A% = InStr(Txt, "and")
         If A% > 0 Then frmImportControls!cSubmitted.Text = Trim(Mid$(Txt, 14, A% - 14))
      End If
   Next i%

   Call CallStackPop

Exit Sub
 
ErrorHandler:
   Call ErrMsg("GlobalModule", "StripCode", Err, Erl())
   Select Case gstErrorFlag
      Case "CANCEL"
         Call CallStackPop
         Exit Sub
      Case "RETRY"
         Resume
      Case "IGNORE"
         Resume Next
   End Select
   
End Sub
