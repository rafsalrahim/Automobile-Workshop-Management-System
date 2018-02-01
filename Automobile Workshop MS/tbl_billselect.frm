VERSION 5.00
Begin VB.Form frm_billselect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bill Print"
   ClientHeight    =   4650
   ClientLeft      =   6810
   ClientTop       =   2745
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "tbl_billselect.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ComboBox comb_vhid 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   2
      Text            =   "----Select One----"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox comb_work 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "tbl_billselect.frx":9AE1
      Left            =   3120
      List            =   "tbl_billselect.frx":9AF1
      TabIndex        =   1
      Text            =   "----Select one----"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SELECT"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle ID"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Work"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   2400
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT THE BILL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   3180
   End
End
Attribute VB_Name = "frm_billselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Public Sub autofill()
STRSQL = "select * from tbl_bodywork where  status='FINISHED' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb_vhid.AddItem RS!vh_id
        RS.MoveNext
    Loop
End Sub

Private Sub comb_vhid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_work_Change()
MsgBox "Should Select One"
End Sub

Private Sub Command1_Click()
If validation Then
If (comb_work.List(comb_work.ListIndex) = "BODY WORK") Then
VECH_ID = comb_vhid.List(comb_vhid.ListIndex)
'frm_bdbill.lbl_id.Caption = VECH_ID
frm_bdbill.Show
ElseIf (comb_work.List(comb_work.ListIndex) = "MECHANICAL REPAIR") Then
VECH_ID = comb_vhid.List(comb_vhid.ListIndex)
frm_mechbill.Show
ElseIf (comb_work.List(comb_work.ListIndex) = "WATER SERVICE") Then
VECH_ID = comb_vhid.List(comb_vhid.ListIndex)
frm_waterbill.Show
ElseIf (comb_work.List(comb_work.ListIndex) = "WHEEL ALLAINGNMENT") Then
VECH_ID = comb_vhid.List(comb_vhid.ListIndex)
frm_wwheelbill.Show
End If
Else
MsgBox "Select Fields"
End If
End Sub

Private Sub Command2_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub Form_Load()
autofill
End Sub
Public Function validation()
Dim ok As Boolean
If (comb_vhid.Text = "----Select One----") Then
ok = False
ElseIf (comb_work.Text = "----Select one----") Then
ok = False
Else
ok = True
End If
validation = ok
End Function


