VERSION 5.00
Begin VB.Form FrmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Threads"
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "10"
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Scripting"
      Height          =   855
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
      Begin VB.CheckBox CBcolor 
         Caption         =   "Color Script"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Log file"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtLog 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "C:\windows\desktop\CGIlog2.txt"
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "okey dokey"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBcolor_Click()
    
    frmScript.Coloring

End Sub

Public Sub Command1_Click()
    
    frmMain.Enabled = False
    frmScript.Enabled = False
    frmAbout.Enabled = False
    'strings
    If Text4.Text <> "" And Val(Text4.Text) <> 0 Then
        Text4.Text = Val(Text4.Text)
    Else
        Call SetWindowPos(FrmSettings.hwnd, -2, 0&, 0&, 0&, 0&, &H2 Or &H1)
        MsgBox "invalid number of threads!"
        Call SetWindowPos(FrmSettings.hwnd, -1, 0&, 0&, 0&, 0&, &H2 Or &H1)
        FrmSettings.Show
        Exit Sub
    End If
    'Check log
    If DoesLOG = False Then
        Call SetWindowPos(FrmSettings.hwnd, -2, 0&, 0&, 0&, 0&, &H2 Or &H1)
        If MsgBox("The file " & txtLog & " does not exist, do you want to create it?", vbYesNo, "Settings") = vbYes Then
            Open txtLog For Output As #1
                Print #1, "[Syphen Scan Logfile]" & vbCrLf & "_____________________" & vbCrLf & "¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯"
            Close #1
        Else
            Call SetWindowPos(FrmSettings.hwnd, -1, 0&, 0&, 0&, 0&, &H2 Or &H1)
            FrmSettings.Show
            Exit Sub
        End If
    End If
    frmMain.Enabled = True
    frmScript.Enabled = True
    frmAbout.Enabled = True
    FrmSettings.Hide
    'save settings
    SaveSetting App.Title, "Settings", "Logfile", FrmSettings.txtLog.Text
    SaveSetting App.Title, "Settings", "Threads", FrmSettings.Text4.Text
    SaveSetting App.Title, "Settings", "ColScript", FrmSettings.CBcolor.Value
    
End Sub

Private Function DoesLOG()
    
    On Error GoTo nope
    Open txtLog For Input As #1
    DoesLOG = True
    Close #1
    Exit Function
nope:
    DoesLOG = False

End Function

Private Sub Form_Unload(Cancel As Integer)

    If frmMain.Visible = True Then
        Cancel = True
        Me.Hide
    End If

End Sub
