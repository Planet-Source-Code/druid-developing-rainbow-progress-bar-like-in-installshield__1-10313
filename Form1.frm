VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Cool Rainbow ProgressBar Demo"
   ClientHeight    =   6045
   ClientLeft      =   3945
   ClientTop       =   3345
   ClientWidth     =   6135
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6135
   Begin VB.Frame Frame3 
      Caption         =   "Color Options :"
      Height          =   1695
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   5655
      Begin VB.PictureBox Picture3 
         Height          =   615
         Left            =   1440
         ScaleHeight     =   555
         ScaleWidth      =   1035
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Select BackColor"
         Height          =   615
         Left            =   2640
         Picture         =   "Form1.frx":030A
         Style           =   1  'Grafisch
         TabIndex        =   12
         Top             =   960
         Width           =   2895
      End
      Begin VB.PictureBox Picture2 
         Height          =   615
         Left            =   1440
         ScaleHeight     =   555
         ScaleWidth      =   1035
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Select ProgressColor"
         Height          =   615
         Left            =   2640
         Picture         =   "Form1.frx":0894
         Style           =   1  'Grafisch
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "BackColor :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "ProgressColor :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -120
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "ProgressBar :"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5895
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ClipControls    =   0   'False
         FillColor       =   &H00FF0000&
         ForeColor       =   &H000000FF&
         Height          =   384
         Left            =   120
         ScaleHeight     =   330
         ScaleWidth      =   5535
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   5595
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Click on the buttons to give a Value to the ProgressBar :"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
      Begin VB.CommandButton Command9 
         Caption         =   "Apply"
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Text            =   "0"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Timer Demo"
         Height          =   615
         Left            =   120
         Picture         =   "Form1.frx":0E1E
         Style           =   1  'Grafisch
         TabIndex        =   6
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Random Value"
         Height          =   615
         Left            =   2880
         Picture         =   "Form1.frx":13C4
         Style           =   1  'Grafisch
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Value 4"
         Height          =   615
         Left            =   2880
         Picture         =   "Form1.frx":1986
         Style           =   1  'Grafisch
         TabIndex        =   4
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Value 3"
         Height          =   615
         Left            =   120
         Picture         =   "Form1.frx":1F48
         Style           =   1  'Grafisch
         TabIndex        =   3
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Value 2"
         Height          =   615
         Left            =   2880
         Picture         =   "Form1.frx":250A
         Style           =   1  'Grafisch
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Value 1"
         Height          =   615
         Left            =   120
         Picture         =   "Form1.frx":2ACC
         Style           =   1  'Grafisch
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Custom Value:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   -120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Select Case Command1.Caption
    Case "Start Timer Demo"
        Timer1.Enabled = True
        Command1.Caption = "Stop Timer Demo"
    Case "Stop Timer Demo"
        Timer1.Enabled = False
        Command1.Caption = "Start Timer Demo"
    End Select
End Sub

Private Sub Command2_Click()
    PaintProgress Picture1, 0.2
End Sub

Private Sub Command3_Click()
    PaintProgress Picture1, 0.4
End Sub

Private Sub Command4_Click()
    PaintProgress Picture1, 0.6
End Sub

Private Sub Command5_Click()
    PaintProgress Picture1, 0.8
End Sub

Private Sub Command6_Click()
    PaintProgress Picture1, Rnd
End Sub

Private Sub Command7_Click()
    CommonDialog1.ShowColor
    Picture2.BackColor = CommonDialog1.Color
End Sub

Private Sub Command8_Click()
    CommonDialog1.ShowColor
    Picture3.BackColor = CommonDialog1.Color
End Sub

Private Sub Command9_Click()
    PaintProgress Picture1, Text1 / 100
End Sub

Private Sub Form_Load()
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Picture1.FillStyle = vbFSSolid
    Picture2.BackColor = vbRed
    Picture3.BackColor = vbWhite
End Sub

Private Sub Timer1_Timer()
    xx = xxbefore + 0.01
    PaintProgress Picture1, xx
    xxbefore = xx
End Sub

Sub PaintProgress(pic As PictureBox, ByVal Percent As Single, Optional ByVal fBorderCase)
    Dim strPercent As String
    Dim intX As Integer
    Dim intY As Integer
    Dim intWidth As Integer
    Dim intHeight As Integer
    If IsMissing(fBorderCase) Then fBorderCase = False
    pic.ForeColor = Picture2.BackColor
    pic.BackColor = Picture3.BackColor
    Dim intPercent
    intPercent = Int(100 * Percent + 0.5)
    If intPercent = 0 Then
        If Not fBorderCase Then
            intPercent = 1
        End If
    ElseIf intPercent = 100 Then
        If Not fBorderCase Then
            intPercent = 99
        End If
    End If
    strPercent = Format$(intPercent) & "%"
    intWidth = pic.TextWidth(strPercent)
    intHeight = pic.TextHeight(strPercent)
    intX = pic.Width / 2 - intWidth / 2
    intY = pic.Height / 2 - intHeight / 2
    pic.DrawMode = 13
    pic.Line (intX, intY)-Step(intWidth, intHeight), pic.BackColor, BF
    pic.CurrentX = intX
    pic.CurrentY = intY
    pic.Print strPercent
    pic.DrawMode = 10
    If Percent > 0 Then
        pic.Line (0, 0)-(pic.Width * Percent, pic.Height), pic.ForeColor, BF
    Else
        pic.Line (0, 0)-(pic.Width, pic.Height), pic.BackColor, BF
    End If
    pic.Refresh
End Sub
