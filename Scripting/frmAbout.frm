VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3240
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5430
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2236.306
   ScaleMode       =   0  'User
   ScaleWidth      =   5099.05
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   960
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "yea what ever!"
      Default         =   -1  'True
      Height          =   465
      Left            =   3840
      TabIndex        =   0
      Top             =   2640
      Width           =   1500
   End
   Begin VB.PictureBox P1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   120
      ScaleHeight     =   2070
      ScaleWidth      =   5115
      TabIndex        =   6
      Top             =   120
      Width           =   5175
      Begin VB.Label LabError 
         Caption         =   "gfgdfhdhdfh"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.PictureBox P2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   120
      ScaleHeight     =   142
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "E-mail me:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Go to the site:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Syphen_2k@hotmail.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "www.hackuk.net"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4958.192
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   4958.192
      Y1              =   1697.936
      Y2              =   1697.936
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim thetop As Long
Dim p1hgt As Long
Dim p1wid As Long
Dim theleft As Long
Dim Tempstring As String

Private Sub cmdOK_Click()

    frmMain.Enabled = True
    frmScript.Enabled = True
    LabError.Visible = False
    Unload Me

End Sub

Private Sub Form_Load()
    
    On Error GoTo Err
    Me.Caption = "About " & App.Title & " " & App.Major & "." & App.Minor
        P1.AutoRedraw = True
        P1.Visible = False
        P1.FontSize = 8
        P1.ForeColor = vbBlack
        P1.BackColor = &H80000004
        P1.ScaleMode = 3
        ScaleMode = 3
        Open (App.Path & "\credits.txt") For Input As #1
        Line Input #1, Tempstring
        P1.Height = (Val(Tempstring) * P1.TextHeight("Test Height")) + 200
        Do Until EOF(1)
            Line Input #1, Tempstring
            PrintText Tempstring
        Loop
        Close #1
        theleft = P2.ScaleLeft
        thetop = P2.ScaleHeight
        p1hgt = P1.ScaleHeight
        p1wid = P1.ScaleWidth
        Timer1.Enabled = True
        
    Exit Sub
Err:
    P1.Visible = True
    LabError.Visible = True
    LabError.Caption = "file " & vbCrLf & App.Path & "\credits.txt" & vbCrLf & "was not found!"
    Timer1.Enabled = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Enabled = True
    frmScript.Enabled = True
    Unload Me

End Sub

Sub Timer1_Timer()
Dim X As Long
Dim Txt As String
        X = BitBlt(P2.hDC, theleft, thetop, p1wid, p1hgt, P1.hDC, 0, 0, &HCC0020)
        thetop = thetop - 1
        If thetop < -p1hgt Then
        Timer1.Enabled = False
        Txt = "Credits Completed"
        CurrentY = P2.ScaleHeight / 2
        CurrentX = (P2.ScaleWidth - P2.TextWidth(Txt)) / 2
        P2.Print Txt
        End If
End Sub

Sub PrintText(Text As String)
Dim X As Long
Dim Y As Long
P1.CurrentX = (P1.ScaleWidth / 2) - (P1.TextWidth(Text) / 2)
P1.ForeColor = 0: X = P1.CurrentX: Y = P1.CurrentY
'For i = 1 To 3
    'P1.Print Text
    X = X + 1: Y = Y + 1: P1.CurrentX = X: P1.CurrentY = Y
'Next i
P1.ForeColor = vbBlack
P1.Print Text
End Sub


Private Sub Label1_Click()
On Error Resume Next
Dim Web_WWW As Long
Dim WebPage As String
WebPage = "http://www.hackuk.net"
Web_WWW = ShellExecute(Me.hwnd, vbNullString, WebPage, vbNullString, "c:\", SW_SHOWNORMAL)

End Sub

Private Sub Label2_Click()
On Error Resume Next
Dim Web_WWW As Long
Dim WebPage As String
WebPage = "mailto:syphen_2k@hotmail.com"
Web_WWW = ShellExecute(Me.hwnd, vbNullString, WebPage, vbNullString, "c:\", SW_SHOWNORMAL)

End Sub
