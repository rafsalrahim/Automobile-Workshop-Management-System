VERSION 5.00
Begin VB.Form frm_splash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4755
   ClientLeft      =   4590
   ClientTop       =   2460
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_splash.frx":0000
   ScaleHeight     =   4755
   ScaleWidth      =   12270
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   11280
      Top             =   3840
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   9720
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   2640
      ScaleHeight     =   435
      ScaleWidth      =   6435
      TabIndex        =   3
      Top             =   3960
      Width           =   6495
      Begin VB.Image Image1 
         Height          =   405
         Left            =   0
         Picture         =   "frm_splash.frx":34FDF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   825
      End
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Files Loading...."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Copyright Researved"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   1800
      Width           =   1920
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Platform : Windows 7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Top             =   1080
      Width           =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "AUTOMOBILE WORKSHOP AND SERVICE MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   420
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   11085
   End
End
Attribute VB_Name = "frm_splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y As Integer
Option Explicit

Private Sub Form_Load()
Y = 0
File1.FileName = App.Path
X = File1.ListCount
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load MDIForm1
MDIForm1.Show
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Timer1_Timer()
Y = Y + 1
If (Image1.Left <= 6480) Then
    Image1.Left = Image1.Left + 200
Else
    Image1.Left = 0
End If
If Y = 20 Then Unload Me

End Sub

