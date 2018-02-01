VERSION 5.00
Begin VB.Form frm_delevary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivary Status"
   ClientHeight    =   5205
   ClientLeft      =   6990
   ClientTop       =   3120
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_delevary.frx":0000
   ScaleHeight     =   5205
   ScaleWidth      =   6000
   Begin VB.TextBox txt_vno 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txt_cstat 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox comb_vid 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lbl_vhno 
      AutoSize        =   -1  'True
      Caption         =   "Vehicle No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   2880
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Current Status"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vehicle ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   1185
   End
End
Attribute VB_Name = "frm_delevary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Public Sub Autofil()
STRSQL = " select * from tbl_vehicleregistration where status='Finished' "
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
        comb_vid.AddItem RS!vh_id
        RS.MoveNext
    Loop
End Sub

Private Sub comb_vid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_vid_Click()
STRSQL = " select * from tbl_vehicleregistration where vh_id='" & comb_vid.List(comb_vid.ListIndex) & "' "
Set RS = adocn.Execute(STRSQL)
txt_cstat.Text = RS!Status
txt_vno.Text = RS!regi_no
End Sub

Private Sub Command1_Click()
If fnValidation Then
STRSQL = "update tbl_vehicleregistration set  status='DELEVARED' where vh_id='" & comb_vid.List(comb_vid.ListIndex) & "' "
        adocn.Execute (STRSQL)
        worker
MsgBox "Vehicle Delevared"
Unload Me
MDIForm1.Show
Else
MsgBox "Select The Vehicle ID"
End If
End Sub
Private Sub Form_Load()
Autofil
End Sub
Public Sub worker()
STRSQL = "update tbl_wstatus set  status='FINISHED' where vh_id='" & comb_vid.List(comb_vid.ListIndex) & "' "
        adocn.Execute (STRSQL)
MsgBox "Vehicle Delevared"
End Sub
Public Function fnValidation()
Dim ok1 As Boolean
If (Trim(txt_cstat.Text) = "") Then
ok1 = False
Else
 If (Trim(txt_vno.Text) = "") Then
  ok1 = False
  Else
  ok1 = True
  End If
  End If
  fnValidation = ok1
End Function
