VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Syphen Scan - [none]"
   ClientHeight    =   2940
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6045
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2940
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Wate 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   1000
      Left            =   2400
      Top             =   1920
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   1
      Left            =   1920
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   600
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Scan scripts (*.scn)|*.scn"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":059E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0856
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1206
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "opens"
                  Text            =   "Open Script"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "openss"
                  Text            =   "Open Scan"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Play"
            Object.ToolTipText     =   "Play"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   5
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Settings"
            Object.ToolTipText     =   "Settings"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Script"
            Object.ToolTipText     =   "Script"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Log"
            Object.ToolTipText     =   "Log"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Text            =   "bob.com"
      Top             =   60
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2670
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5821
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1402
            MinWidth        =   1412
            Text            =   "Threads:"
            TextSave        =   "Threads:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   707
            MinWidth        =   707
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1402
            MinWidth        =   1412
            Text            =   "Found:"
            TextSave        =   "Found:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   707
            MinWidth        =   707
            Text            =   "0"
            TextSave        =   "0"
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   5000
      Left            =   1440
      Top             =   1440
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   1920
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfSource 
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393217
      DisableNoScroll =   -1  'True
      TextRTF         =   $"Form1.frx":1362
   End
   Begin MSComctlLib.ListView LVstat 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Server"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   4657
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Starting server:"
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   60
      Width           =   1095
   End
   Begin VB.Menu menfile 
      Caption         =   "File"
      Begin VB.Menu menOpens 
         Caption         =   "Open Script"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu menSave 
         Caption         =   "Save Scan"
      End
      Begin VB.Menu menOpenss 
         Caption         =   "Open Scan"
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu menquit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu menview 
      Caption         =   "View"
      Begin VB.Menu menScript 
         Caption         =   "Script"
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu menLog 
         Caption         =   "View Log"
      End
      Begin VB.Menu menScaned 
         Caption         =   "View Scanned"
      End
      Begin VB.Menu menToScan 
         Caption         =   "View To Scan"
      End
      Begin VB.Menu space4 
         Caption         =   "-"
      End
      Begin VB.Menu menSet 
         Caption         =   "Settings"
      End
      Begin VB.Menu menabout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bob
Dim list2 As New Collection
Dim List1 As New Collection
Dim sScript As String
Dim Test
Dim Saved
Dim FoundServers As String

Sub ClearLists()

    For i = 1 To List1.Count
        List1.Remove (1)
    Next i
    For i = 1 To list2.Count
        list2.Remove (1)
    Next i

End Sub
Function removeNotes(Script)

    On Error Resume Next
    'go threw each line removing notes
    Dim newScript As String
    For i = 0 To 1000
        If Left(Split(Script, vbCrLf)(i), 1) <> "'" Then
            If i = 0 Then
                newScript = Split(Script, vbCrLf)(i)
            Else
                newScript = newScript & vbCrLf & Split(Script, vbCrLf)(i)
            End If
        End If
    Next i
    removeNotes = newScript

End Function
Public Sub TestScript(StartServer)
    
    Test = True
    sScript = removeNotes(frmScript.Text1.Text)
    'remove http://
    If Left(StartServer, 7) = "http://" Then StartServer = Replace(StartServer, "http://", "")
    'remove www.
    If Left(StartServer, 4) = "www." Then StartServer = Replace(StartServer, "www.", "")
    Winsock1(1).Connect StartServer, Split(sScript, vbCrLf & "#")(0)

End Sub
Sub info(Text, Index)

    On Error Resume Next
    If Right(Text, Len(vbCrLf)) = vbCrLf Then Text = Left(Text, Len(Text) - Len(vbCrLf))
    If Test = True Then
        frmScript.Text4 = frmScript.Text4 & vbCrLf & Text
    Else
        With LVstat.ListItems(wTFI(Index))
            If .Text <> Index Then .Text = Index
            If .SubItems(1) <> Winsock1(Index).RemoteHost Then .SubItems(1) = Winsock1(Index).RemoteHost
            If .SubItems(2) <> Text Then .SubItems(2) = Text
        End With
    End If

End Sub

Function find(FindIN As String, findWhat)

    Dim bob
    bob = False
    For i = 1 To Len(FindIN)
        If Mid(LCase(FindIN), i, Len(findWhat)) = LCase(findWhat) Then
        bob = True
        
        End If
    Next i
    find = bob

End Function

Private Sub excute(Index)

Dim LineToRun As String
Dim LastData As String
With Winsock1(Index)

    LineToRun = Split(sScript, vbCrLf & "#")(Split(Winsock1(Index).Tag, ">|<")(1))
    LineToRun = Replace(LineToRun, "$host", .RemoteHost)
    LineToRun = Replace(LineToRun, "$ip", .RemoteHostIP)
    LastData = Split(.Tag, ">|<")(0)
    .Tag = Split(.Tag, ">|<")(0) & ">|<" & Split(.Tag, ">|<")(1) + 1 & ">|<FALSE"
    'start the thing!
    If Left(LCase(LineToRun), 4) = "wait" Then
        info "looking for exploit: Waiting...", Index
        TimeOut(Index).Enabled = True
        .Tag = Split(.Tag, ">|<")(0) & ">|<" & Split(.Tag, ">|<")(1) & ">|<TRUE"
    ElseIf Left(LCase(LineToRun), 4) = "send" Then
        info "looking for exploit: Sending - " & Right(LineToRun, (Len(LineToRun) - 5)), Index
        .SendData Right(LineToRun, (Len(LineToRun) - 5))
        excute (Index)
    ElseIf Left(LCase(LineToRun), 4) = "quit" Then
        info "looking for exploit: QUITING!", Index
        List1.Add .RemoteHost
        .Close
        info "looking for exploit: QUITTED!", Index
        next_Server (Index)
    ElseIf Left(LCase(LineToRun), 3) = "add" Then
        info "looking for exploit: ADDING!", Index
        foundExp Index, Right(LineToRun, (Len(LineToRun) - 3))
        excute (Index)
    ElseIf Left(LCase(LineToRun), 2) = "if" Then
        'there is an IF argument
        'is it a Then or an Else?
        Dim ThenElse
        If find(LCase(LineToRun), "then") = True Then
            ThenElse = "then"
            Value = Right(Split(LCase(LineToRun), " then")(0), Len(Split(LCase(LineToRun), " then")(0)) - 3)
        ElseIf find(LCase(LineToRun), "else") = True Then
            ThenElse = "else"
            Value = Right(Split(LCase(LineToRun), " else")(0), Len(Split(LCase(LineToRun), " else")(0)) - 3)
        Else
            StatusBar.Panels(1) = "IF with no THEN or ELSE"
            Exit Sub
        End If
        Dim Vlue As String
        'IF
            Dim IsValueTrue
            IsValueTrue = False
            If Left(LCase(Value), 4) = "left" Then
                If Left(LCase(LastData), Len(Value) - 5) = Right(Value, Len(Value) - 5) Then IsValueTrue = True
            ElseIf Left(LCase(Value), 5) = "right" Then
                If Right(LCase(LastData), Len(Value - 6)) = Right(Value, Len(Value) - 6) Then IsValueTrue = True
            ElseIf Left(LCase(Value), 5) = "cords" Then
                Dim ValOne As String
                Dim ValTwo
                ValOne = Val(Split(Right(Value, Len(Value) - 6), " ")(0))
                ValTwo = Right(Value, Len(Value) - Len(ValOne) - 7)
                If Mid(LCase(LastData), ValOne, Len(ValTwo)) = ValTwo Then IsValueTrue = True
            ElseIf Left(LCase(Value), 4) = "find" Then
                IsValueTrue = find(LCase(LastData), Right(Value, Len(Value) - 5))
            End If
        'THEN or ELSE
            Dim doStuff
            doStuff = False
            If ThenElse = "then" And IsValueTrue = True Then
                doStuff = True
            ElseIf ThenElse = "else" And IsValueTrue = False Then
                doStuff = True
            End If
            If doStuff = True Then
                Dim ThenDo
                If ThenElse = "then" Then
                    ThenDo = Split(LCase(LineToRun), "then ")(1)
                ElseIf ThenElse = "else" Then
                    ThenDo = Split(LCase(LineToRun), "else ")(1)
                End If
                
                If LCase(ThenDo) = "quit" Then
                    info "looking for exploit: QUITING!", Index
                    List1.Add .RemoteHost
                    .Close
                    info "looking for exploit: QUITTED!", Index
                    next_Server (Index)
                ElseIf Left(LCase(ThenDo), 3) = "add" Then
                    info "looking for exploit: ADDING!", Index
                    foundExp Index, Right(ThenDo, (Len(ThenDo) - 3))
                    excute (Index)
                ElseIf Left(LCase(ThenDo), 4) = "send" Then
                    info "looking for exploit: Sending - " & Right(ThenDo, (Len(ThenDo) - 5)), Index
                    .SendData Right(ThenDo, (Len(ThenDo) - 5))
                    excute (Index)
                ElseIf LCase(ThenDo) = "wait" Then
                    info "looking for exploit: Waiting...", Index
                    TimeOut(Index).Enabled = True
                    .Tag = Split(.Tag, ">|<")(0) & ">|<" & Split(.Tag, ">|<")(1) & ">|<TRUE"
                End If
            Else
                excute (Index)
            End If
    End If
End With

End Sub
Private Sub Form_Load()

    Call SetWindowPos(FrmSettings.hwnd, -1, 0&, 0&, 0&, 0&, &H2 Or &H1)
    bob = 0
    Saved = False
    'load settings
    FrmSettings.txtLog.Text = GetSetting(App.Title, "Settings", "Logfile", LCase(App.Path) & "\Log.txt")
    FrmSettings.Text4.Text = GetSetting(App.Title, "Settings", "Threads", 10)
    FrmSettings.CBcolor.Value = GetSetting(App.Title, "Settings", "ColScript", 1)
    'check settings
    Call FrmSettings.Command1_Click
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    LVstat.Width = Me.ScaleWidth
    LVstat.Height = Me.ScaleHeight - LVstat.Top - StatusBar.Height
    txtStart.Width = LVstat.Width - txtStart.Left

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    End

End Sub

Private Sub menabout_Click()
    If frmAbout.Visible = False Then Unload frmAbout
    frmMain.Enabled = False
    frmScript.Enabled = False
Unload frmAbout
frmAbout.Show
End Sub

Private Sub menLog_Click()
    
    Shell "C:\windows\notepad.exe " & FrmSettings.txtLog.Text, vbNormalFocus

End Sub

Private Sub menOpens_Click()

    frmScript.opens

End Sub

Private Sub menOpenss_Click()
    
    On Error Resume Next
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Open CommonDialog1.FileName For Input As #1
        Do Until EOF(1)
            Line Input #1, lineoftext$
            If alltext$ <> "" Then
                alltext$ = alltext$ & vbCrLf & lineoftext$
            Else
                alltext$ = lineoftext$
            End If
            ScanData = alltext$
        Loop
        Close #1
        Dim sd1a As String
        Dim sd2a As String
        frmScript.Text1.Text = Split(ScanData, vbCrLf & ":#:")(0)
        sd1a = Split(ScanData, vbCrLf & ":#:")(2)
        sd2a = Split(ScanData, vbCrLf & ":#:")(1)
        'scanned
        ClearLists
        For i = 1 To 1000
            If Split(sd1a, vbCrLf)(i) <> "" Then List1.Add Split(sd1a, vbCrLf)(i)
        Next i
        'To scan
        For i = 1 To 50
            If Split(sd2a, vbCrLf)(i) <> "" Then list2.Add Split(sd2a, vbCrLf)(i)
        Next i
        Saved = True
        frmMain.Caption = "Syphen Scan - " & CommonDialog1.FileName
    End If
    
End Sub

Private Sub menquit_Click()

    End

End Sub

Private Sub menSave_Click()
    
    Dim strData As String
    Dim lenList2
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        strData = frmScript.Text1.Text
        Do While Right(strData, Len(vbCrLf)) = vbCrLf
            strData = Left(strData, Len(strData) - Len(vbCrLf))
        Loop
        strData = strData & vbCrLf & ":#:"
        If list2.Count > 50 Then
            lenList2 = 50
        Else
            lenList2 = list2.Count
        End If
        For i = 1 To lenList2
            strData = strData & vbCrLf & list2.Item(i)
        Next i
        strData = strData & vbCrLf & ":#:"
        For i = 1 To List1.Count
            strData = strData & vbCrLf & List1.Item(i)
        Next i
        Open CommonDialog1.FileName For Output As #1
        Print #1, strData
        Close #1
    End If
    
End Sub

Private Sub menScaned_Click()
    
    On Error Resume Next
    strData = List1.Item(1)
    For i = 2 To List1.Count
        strData = strData & vbCrLf & List1.Item(i)
    Next i
    Open App.Path & "/Scaned.tmp" For Output As #1
        Print #1, strData
    Close #1
    Shell "C:\windows\notepad.exe " & App.Path & "/Scaned.tmp", vbNormalFocus
    Kill App.Path & "/Scaned.tmp"

End Sub

Private Sub menScript_Click()
    
    frmScript.Show

End Sub

Private Sub menSet_Click()
    
    frmMain.Enabled = False
    frmScript.Enabled = False
    frmAbout.Enabled = False
    FrmSettings.Show

End Sub

Private Sub menToScan_Click()

    On Error Resume Next
    strData = list2.Item(1)
    For i = 2 To list2.Count
        strData = strData & vbCrLf & list2.Item(i)
    Next i
    Open App.Path & "/ToScan.tmp" For Output As #1
        Print #1, strData
    Close #1
    Shell "C:\windows\notepad.exe " & App.Path & "/ToScan.tmp", vbNormalFocus
    Kill App.Path & "/ToScan.tmp"

End Sub

Private Sub TimeOut_Timer(Index As Integer)

    TimeOut(Index).Enabled = False
    info "looking for exploit: Timed out, Closing conection!", Index
    Winsock1(Index).Close
    'go on to next one!
    next_Server (Index)

End Sub

Private Sub Timer1_Timer()
    
    StatusBar.Panels(3).Text = LVstat.ListItems.Count
    If LVstat.ListItems.Count = 0 And frmScript.Text1.Text <> "" And Toolbar1.Buttons(4).Enabled = False Then
        Toolbar1.Buttons(4).Enabled = True
        StatusBar.Panels(1).Text = "Syphen Scan Reset!"
    ElseIf LVstat.ListItems.Count = 0 And frmScript.Text1.Text = "" And Toolbar1.Buttons(4).Enabled = True Then
        Toolbar1.Buttons(4).Enabled = False
    End If

    'check if all strings are wating for a server
    Dim Allwaiting  As Boolean
    Allwaiting = True
    Dim NewServer As String
    If LVstat.ListItems.Count <> 0 Then
        For i = 1 To LVstat.ListItems.Count
            If LVstat.ListItems(i).SubItems(2) <> "waiting for server to scan" Or Allwaiting = False Then
                Allwaiting = False
            Else
                Allwaiting = True
            End If
        Next i
        If Allwaiting = True Then
            NewServer = InputBox("the Scanner has run out of servers to scan. enter another to continue or press cancel to end the scan", "new sever needed")
            If NewServer <> "" Then
                list2.Add NewServer
                next_Server (1)
            Else
                'cancel scan!
                Toolbar1.Buttons(5).Value = tbrPressed
            End If
            Timer1.Enabled = False
        End If
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Key = "Settings" Then
        frmMain.Enabled = False
        frmScript.Enabled = False
        frmAbout.Enabled = False
        FrmSettings.Show
    ElseIf Button.Key = "Log" Then
        Shell "C:\windows\notepad.exe " & FrmSettings.txtLog.Text, vbNormalFocus
    ElseIf Button.Key = "Script" Then
        frmScript.Show
    ElseIf Button.Key = "Play" Then
        'change \/
        FoundServers = ""
        
        'dissable all
        FrmSettings.Text4.Locked = True
        StatusBar.Panels(1).Text = "Starting Scan"
        'PLAY
        Test = False
        LVstat.ListItems.Add , , "1"
        StatusBar.Panels(5).Text = "0"
        If Saved = False Then
            ClearLists
            sScript = removeNotes(frmScript.Text1.Text)
            'remove http://
            If Left(txtStart, 7) = "http://" Then txtStart = Replace(txtStart, "http://", "")
            'remove www.
            If Left(txtStart, 4) = "www." Then txtStart = Replace(txtStart, "www.", "")
            check_url 1, txtStart
        Else
            check_url 1, list2.Item(1)
        End If
    ElseIf Button.Key = "open" Then
        frmScript.opens
    ElseIf Button.Key = "QuickStop" Then
        LVstat.ListItems.Remove (wTFI(1))
        Winsock1(1).Close
        For i = 2 To Winsock1.Count
            Winsock1(i).Close
            LVstat.ListItems.Remove (wTFI(i))
            Unload Winsock1(i)
            Unload Winsock2(i)
            Unload rtfSource(i)
            Unload TimeOut(i)
            Unload Wate(i)
        Next i
    ElseIf Button.Key = "Save" Then
    menSave_Click
    ElseIf Button.Key = "Stop" Then
        Saved = False
        Toolbar1.Buttons(4).Enabled = False
        StatusBar.Panels(1).Text = "Stoping Scan"
        FrmSettings.Text4.Locked = False
    End If

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    'On Error Resume Next
    Dim ScanData
    If ButtonMenu.Key = "opens" Then
        frmScript.opens
    ElseIf ButtonMenu.Key = "openss" Then
        menOpenss_Click
    End If

End Sub

Private Sub Wate_Timer(Index As Integer)
Wate(Index).Enabled = False
next_Server (Index)
End Sub

Private Sub Winsock1_Connect(Index As Integer)
    
    TimeOut(Index).Enabled = False
    info "Conected", Index
    Winsock1(Index).Tag = ">|<1>|<FALSE"
    excute Index

End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    TimeOut(Index).Enabled = False
    On Error Resume Next '<<<<<< REMOVE!!
    Dim strData As String
    With Winsock1(Index)
        .GetData strData
        strData = Replace(strData, Chr(10), vbCrLf)
        
        info "looking for exploit: Receved - " & strData, Index
        .Tag = strData & ">|<" & Split(.Tag, ">|<")(1) & ">|<" & Split(.Tag, ">|<")(2)
        If Split(.Tag, ">|<")(2) = "TRUE" Then
            excute Index
        End If
    End With

End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    If Description = "No buffer space is available" Then
        StatusBar.Panels(1).Text = "too much strings open, closing string " & Index
        LVstat.ListItems.Remove (wTFI(Index))
        StatusBar.Panels(3).Text = LVstat.ListItems.Count
        If Index = 1 Then Exit Sub
        Unload Winsock1(Index)
        Unload Winsock2(Index)
        Unload rtfSource(Index)
        Unload TimeOut(Index)
        Unload Wate(Index)
    Else
        info "ERROR:" & Description, Index
        Winsock1(Index).Close 'note it does'nt add it to the checked urls list
        next_Server (Index)
    End If

End Sub
Function Is_this_str_In_this_list(strServer As String, lstLook As Collection)
    
    'small function to check if a server is in a list
    For i = 1 To lstLook.Count
        If lstLook.Item(i) = strServer Then  'sorry man its there
            Is_this_str_In_this_list = True
            Exit Function
        End If
    Next
    Is_this_str_In_this_list = False
    
End Function

Function Is_listView(strServer As String)

    For i = 1 To LVstat.ListItems.Count
        If LCase(LVstat.ListItems(i).SubItems(1)) = LCase(strServer) Then
            Is_this_str_In_this_listView = True
            Exit Function
        End If
    Next i
    Is_this_str_In_this_listView = False

End Function
Public Sub Add_Server_To_Check(strServer As String)
    
    'i put the yahoo thing in cos theres no point in scanning there network
    strServer = Replace(strServer, "www.", "")
    'if the server is not already in list2 add it
    If Is_this_str_In_this_list(strServer, list2) = False And Is_this_str_In_this_list(strServer, List1) = False And Is_listView(strServer) = False Then
        If Right(strServer, Len("yimg.com")) <> "yimg.com" And _
        Right(strServer, Len("yahoo.com")) <> "yahoo.com" Then 'its not there
            list2.Add strServer
        End If
    End If
    
End Sub

Function wTFI(Number)

    For i = 1 To LVstat.ListItems.Count
        If LVstat.ListItems(i).Text = Number Then wTFI = i
    Next i

End Function
Public Sub next_Server(Index)

    On Error Resume Next
    If Test = True Then Exit Sub

    If Toolbar1.Buttons(5).Value = tbrPressed Then
        'Winsock1(index).Close
        LVstat.ListItems.Remove (wTFI(Index))
        StatusBar.Panels(3).Text = LVstat.ListItems.Count
        If Index = 1 Then Exit Sub
        Unload Winsock1(Index)
        Unload Winsock2(Index)
        Unload rtfSource(Index)
        Unload TimeOut(Index)
        Unload Wate(Index)
    Else
        If list2.Count = 0 Then
            info "waiting for server to scan", Index
            'next_Server (Index)
            Wate(Index).Enabled = True
            Exit Sub
        End If
        Timer1.Enabled = True
        Dim strServer As String
        'opens new strings
        For i = (Winsock1.Count + 1) To FrmSettings.Text4
            If list2.Count > 1 And FrmSettings.Text4 > Winsock1.Count And Toolbar1.Buttons(5).Value = tbrUnpressed Then
                Load Winsock2(i)
                Load Winsock1(i)
                Load rtfSource(i)
                Load TimeOut(i)
                Load Wate(i)
                LVstat.ListItems.Add LVstat.ListItems.Count + 1, , i
                strServer = list2.Item(1)
                list2.Remove (1)
                check_url i, strServer
                info "Making new Thread: " & i, Index
            End If
        Next
        StatusBar.Panels(3).Text = LVstat.ListItems.Count
        strServer = ""
        strServer = list2.Item(1)
        list2.Remove (1)
        info "Moving to next server: " & strServer, Index
        If strServer <> "" Then
            check_url Index, strServer
        Else
            info "damn!: " & strServer, Index
        End If
    End If

End Sub

Public Sub check_url(Index, URL)

    info "Geting Index.html", Index
    LVstat.ListItems(wTFI(Index)).SubItems(1) = URL
    Winsock2(Index).Connect URL, 80
    
End Sub

Private Sub Winsock2_Connect(Index As Integer)
info "Geting Index.html-Downloading", Index
rtfSource(Index).Text = ""
LVstat.ListItems(wTFI(Index)).SubItems(1) = Winsock2(Index).RemoteHost
Winsock2(Index).Tag = Winsock2(Index).RemoteHost

Data2Send = "GET / HTTP/1.1" & vbCrLf & _
"Host: " & Winsock2(Index).RemoteHost & vbCrLf & _
"Connection: Close " & vbCrLf & _
"User-Agent: Sam Spade 1.14" & vbCrLf & _
vbCrLf

Winsock2(Index).SendData Data2Send
End Sub
Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    On Error Resume Next
    Dim strData As String
    Winsock2(Index).GetData strData
    strData = Replace(strData, Chr(10), vbCrLf)
    rtfSource(Index).Text = rtfSource(Index).Text & strData
    If Winsock2(Index).RemoteHost = "bob.com" Then Clipboard.SetText rtfSource(Index).Text
    
End Sub
Private Sub Winsock2_Close(Index As Integer)

    On Error Resume Next
    
    Dim ServerStart
    Dim End1
    Dim End2
    Dim ServerEnd
    
    With rtfSource(Index)
        Winsock2(Index).Close
        info "Extracting servers", Index
        .find "http://", 0
        Do While .SelText = "http://"
            ServerStart = .SelStart + 7
            .find """", ServerStart
            If .SelText = """" Then End1 = .SelStart
            .find "/", ServerStart
            If .SelText = "/" Then End2 = .SelStart
            If End1 < End2 Then ' " is closer than /
                ServerEnd = End1
            End If
            If End1 > End2 Then ' / is closer than "
                ServerEnd = End2
            End If
            .SelStart = ServerStart
            .SelLength = ServerEnd - ServerStart
            'seltext is now the server
            If .SelText <> "" Then Add_Server_To_Check (.SelText)
            .find "http://", ServerStart
        Loop
        .Text = ""
        info "looking for exploit: connecting", Index
        LVstat.ListItems(wTFI(Index)).SubItems(1) = Winsock2(Index).RemoteHost
        TimeOut(Index).Enabled = True
        Winsock1(Index).Close
        Winsock1(Index).Connect Winsock2(Index).Tag, Split(sScript, vbCrLf & "#")(0)
    End With
    Exit Sub
    

End Sub

Sub foundExp(Index, data)

    Dim strData
    'open the list
    Open FrmSettings.txtLog For Input As #1
        Do Until EOF(1)
            Line Input #1, lineoftext$
            If alltext$ <> "" Then
                alltext$ = alltext$ & vbCrLf & lineoftext$
            Else
                alltext$ = lineoftext$
            End If
            strData = alltext$
        Loop
    Close #1
    'check if the server is already found
    If find(FoundServers, Winsock1(Index).RemoteHost) = False Then
        FoundServers = FoundServers & Winsock1(Index).RemoteHost & vbCrLf
        StatusBar.Panels(5).Text = StatusBar.Panels(5).Text + 1
        'add the found server to the list
        strData = strData & vbCrLf & Winsock1(Index).RemoteHost & data
        Open FrmSettings.txtLog For Output As #1
            'save the new list
            Print #1, strData
        Close #1
    End If

End Sub

Private Sub Winsock2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
        info "error: " & Description, Index
        Winsock2(Index).Close 'note it does'nt add it to the checked urls list
        next_Server (Index)
End Sub
