VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmScript 
   Caption         =   "Script - [none]"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4440
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox text2 
      Height          =   135
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form2.frx":058A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3975
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7011
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form2.frx":060C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "ini Scripts (*.ini)|*.ini"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "test"
            Object.ToolTipText     =   "Test script"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stop"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Check"
            Object.ToolTipText     =   "Check Script for errors"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Script"
            ImageIndex      =   5
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "showTest"
            ImageIndex      =   6
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0689
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":07E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0941
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0A9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0BF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1195
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":15E9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Text            =   "Domain.com"
      Top             =   90
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   3975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Test server:"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
On Error Resume Next
Text1.Width = Me.ScaleWidth
Text1.Height = Me.ScaleHeight - Text1.Top
Text4.Width = Text1.Width
Text4.Height = Text1.Height
Text3.Width = Me.ScaleWidth - Text3.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmMain.Visible = True Then
    Cancel = True
    Me.Hide
End If
End Sub

Private Sub Text1_Change()
If FrmSettings.CBcolor.Value = 1 Then text2.Text = Text1.Text: Coloring
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Key = "test" Then
    Text4.Text = "Testing..."
    Text1.Visible = False
    Text4.Visible = True
    Toolbar1.Buttons(9).Value = tbrPressed
    frmMain.TestScript (Text3.Text)
ElseIf Button.Key = "Script" Then
    Text4.Visible = False
    Text1.Visible = True
ElseIf Button.Key = "showTest" Then
    Text1.Visible = False
    Text4.Visible = True
ElseIf Button.Key = "stop" Then
    frmMain.Winsock1(1).Close
    Text4.Text = Text4.Text & vbCrLf & "Stoped"
ElseIf Button.Key = "open" Then
    opens
ElseIf Button.Key = "save" Then
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Open CommonDialog1.FileName For Output As #1
        Print #1, Text1.Text
        Me.Caption = "Script - " & CommonDialog1.FileName
        Close #1
    End If
ElseIf Button.Key = "Check" Then
 Coloring
End If
End Sub
Public Sub Coloring()
'i know its fucked up but its the only way i can think of doing it :(
On Error Resume Next
Dim whereWasit
With text2
    whereWasit = Text1.SelStart
    .SelStart = 0
    .SelLength = Len(.Text)
    .SelColor = &H80000006
    If FrmSettings.CBcolor.Value = 1 Then
        .find vbCrLf & "#", 0
        Do While .SelText = vbCrLf & "#"
            .SelStart = .SelStart + 3
            .SelLength = 3
            If LCase(.SelText) = "if " Then
                .SelColor = &H800080            '3rd
                Dim Sstart
                Dim nextcrlf
                Sstart = .SelStart
                .find vbCrLf, .SelStart
                nextcrlf = .SelStart
                .find " then ", Sstart, nextcrlf
                If LCase(.SelText) = " then " Then
                    .SelColor = &H800080
                    .SelStart = .SelStart + 6
                    .SelLength = 3
                    If LCase(.SelText) = "add" Then
                        .SelColor = &HFF0000
                    Else
                        .SelLength = 4
                        If LCase(.SelText) = "wait" Or LCase(.SelText) = "send" Or LCase(.SelText) = "quit" Then .SelColor = &HFF0000
                    End If
                End If
                .find " else ", Sstart, nextcrlf
                If LCase(.SelText) = " else " Then
                    .SelColor = &H800080
                    .SelStart = .SelStart + 6
                    .SelLength = 3
                    If LCase(.SelText) = "add" Then
                        .SelColor = &HFF0000
                    Else
                        .SelLength = 4
                        If LCase(.SelText) = "wait" Or LCase(.SelText) = "send" Or LCase(.SelText) = "quit" Then .SelColor = &HFF0000
                    End If
                End If
            Else
                If LCase(.SelText) = "add" Then
                    .SelColor = &HFF0000        '2nd
                Else
                    .SelLength = 4
                    If LCase(.SelText) = "wait" Or LCase(.SelText) = "send" Or LCase(.SelText) = "quit" Then
                        .SelColor = &HFF0000        '2nd
                    Else
                        'theres an error
                        Dim linestart
                        Dim lineend
                        linestart = .SelStart
                        .find vbCrLf, linestart
                        lineend = .SelStart
                        .SelStart = linestart
                        .SelLength = lineend - linestart
                        .SelColor = &HFF&
                    End If
                End If
            End If
            .find vbCrLf & "#", .SelStart
        Loop
        'find note on first line
        If Left(.Text, 1) = "'" Then
            .SelStart = 0
            .SelLength = Len(Split(.Text, vbCrLf)(0))
            .SelColor = &H8000&
        End If
        'find notes on other lines
        .find vbCrLf & "'", 0
        Do While .SelText = vbCrLf & "'"
            Dim noteStart
            Dim noteend
            noteStart = .SelStart + 2
            .find vbCrLf, noteStart
            noteend = .SelStart
            .SelStart = noteStart
            .SelLength = noteend - noteStart
            .SelColor = &H8000&
            .find vbCrLf & "'", noteend
        Loop
    End If
    If Text1.TextRTF <> .TextRTF Then Text1.TextRTF = .TextRTF
    Text1.SelStart = whereWasit
End With
End Sub

Public Sub opens()

    
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
            Text1.Text = alltext$
        Loop
        Close #1
        Me.Caption = "Script - " & CommonDialog1.FileName
        frmMain.Caption = "Syphen Scan - " & CommonDialog1.FileName
    End If
End Sub

