VERSION 5.00
Begin VB.Form frm_vehclass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   7410
   ClientTop       =   1845
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_vehclass.frx":0000
   ScaleHeight     =   5295
   ScaleWidth      =   6255
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txt_rate 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txt_model 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frm_vehclass.frx":3830
      Left            =   3000
      List            =   "frm_vehclass.frx":3843
      TabIndex        =   2
      Text            =   "__SELECT ONE__"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lbl_rate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Filed required"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   2520
      Width           =   1365
   End
   Begin VB.Label lbl_mod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Filed required"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   1800
      Width           =   1365
   End
   Begin VB.Label lbl_vhcom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Filed required"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wheel allignment Rate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   2760
      Width           =   2010
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Model "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Company"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADD NEW VEHICLE MODEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   4230
   End
End
Attribute VB_Name = "frm_vehclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Dim STRSQL1 As String
Dim RS1 As ADODB.Recordset
Dim rate As Integer

Private Sub Combo1_Change()
MsgBox "Should Select One"
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
lbl_vhcom.Visible = True
KeyAscii = 0
End Sub
Private Sub Command1_Click()
If validation Then
STRSQL1 = "select * from tbl_wheelrate where service= '" & txt_model.Text & "'"
Set RS1 = adocn.Execute(STRSQL1)
If RS1.RecordCount < 1 Then
inseart
MsgBox "Model Added"
Unload Me
MDIForm1.Show
Else
MsgBox "Already Exist"
Unload Me
MDIForm1.Show
End If
Else
MsgBox "Some Problem While Insearting"
cleatlbl
End If
End Sub
Public Sub inseart()
rate = val(txt_rate.Text)
STRSQL = "insert into tbl_wheelrate(vh_company,service,rate) values ('" & Combo1.List(Combo1.ListIndex) & "','" & txt_model.Text & "','" & rate & "') "
Set RS = adocn.Execute(STRSQL)
End Sub
Public Function validation()
Dim v1 As Boolean
If Combo1.Text = "__SELECT ONE__" Then
        lbl_vhcom.Visible = True
        v1 = False
    Else
    v1 = True
     End If
If Trim(txt_model.Text) = "" Then
        lbl_mod.Visible = True
        v1 = False
    Else
     v1 = True
     End If
If Trim(txt_rate.Text) = "" Then
        lbl_rate.Visible = True
        v1 = False
    Else
     v1 = True
     End If
   If (v1 = True) Then
   validation = True
   Else
   validation = False
   End If
End Function

Private Sub Command2_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub Form_Load()
cleatlbl
End Sub

Private Sub txt_model_Change()
If Trim(txt_model.Text) = "" Then
        lbl_mod.Visible = True
    Else
        lbl_mod.Visible = False
    End If
End Sub

Private Sub txt_model_KeyPress(KeyAscii As Integer)
If Len(txt_model.Text) = 30 Then
MsgBox "Limit Crossed", vhinformation
txt_model.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_rate_Change()
If Trim(txt_rate.Text) = "" Then
        lbl_rate.Visible = True
    Else
        lbl_rate.Visible = False
    End If
 End Sub
Public Sub cleatlbl()
 lbl_rate.Visible = False
 lbl_mod.Visible = False
  lbl_vhcom.Visible = False
End Sub

Private Sub txt_rate_KeyPress(KeyAscii As Integer)
If Len(txt_rate.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_rate.Text = ""
End If
If KeyAscii > 48 And KeyAscii > 58 And KeyAscii <> 8 And KeyAscii <> 13 Then
MsgBox "Only integer is allowed"
KeyAscii = 0
End If
End Sub
