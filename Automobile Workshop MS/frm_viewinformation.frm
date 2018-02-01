VERSION 5.00
Begin VB.Form frm_viewinformation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Information"
   ClientHeight    =   7860
   ClientLeft      =   2430
   ClientTop       =   1470
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   12165
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.ComboBox combo_stwork 
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
         ItemData        =   "frm_viewinformation.frx":0000
         Left            =   2160
         List            =   "frm_viewinformation.frx":0010
         TabIndex        =   24
         Text            =   "----SELECT ONE------"
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox txt_cstatus 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   23
         Top             =   4200
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   22
         Top             =   6360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   21
         Top             =   6360
         Width           =   1335
      End
      Begin VB.TextBox txt_comp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txt_model 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox txt_regno 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txt_desc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   975
         Left            =   5400
         TabIndex        =   12
         Top             =   1080
         Width           =   6135
         Begin VB.CheckBox Check1 
            Caption         =   "Body"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Glass"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Paint"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   14
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Others"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4440
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   5400
         TabIndex        =   7
         Top             =   2280
         Width           =   6135
         Begin VB.CheckBox Check5 
            Caption         =   "Engine"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Oil Change"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Break"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Others"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   5400
         TabIndex        =   3
         Top             =   3480
         Width           =   6135
         Begin VB.CheckBox Check9 
            Caption         =   "Full Body"
            Enabled         =   0   'False
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
            Left            =   840
            TabIndex        =   6
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Chase"
            Enabled         =   0   'False
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
            Left            =   2520
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Others"
            Enabled         =   0   'False
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
            Left            =   4200
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.ComboBox comb_change 
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
         ItemData        =   "frm_viewinformation.frx":0055
         Left            =   8520
         List            =   "frm_viewinformation.frx":005F
         TabIndex        =   2
         Text            =   "----SELECT ONE------"
         Top             =   5280
         Width           =   2415
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
         Left            =   2160
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Image Image4 
         Height          =   5640
         Left            =   4800
         Picture         =   "frm_viewinformation.frx":0079
         Top             =   960
         Width           =   6915
      End
      Begin VB.Label lbl_cstat 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7680
         TabIndex        =   35
         Top             =   5280
         Width           =   135
      End
      Begin VB.Label lbl_wtyp 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1440
         TabIndex        =   34
         Top             =   3600
         Width           =   135
      End
      Begin VB.Image Image3 
         Height          =   4080
         Left            =   5280
         Picture         =   "frm_viewinformation.frx":A571
         Top             =   1080
         Width           =   6315
      End
      Begin VB.Image Image2 
         Height          =   2745
         Left            =   5400
         Picture         =   "frm_viewinformation.frx":16409
         Top             =   2040
         Width           =   6120
      End
      Begin VB.Image Image1 
         Height          =   2505
         Left            =   5400
         Picture         =   "frm_viewinformation.frx":1CD27
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Current Status"
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
         Left            =   240
         TabIndex        =   33
         Top             =   4200
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "UPDATE    STATUS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         TabIndex        =   32
         Top             =   240
         Width           =   2580
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Work Type"
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
         Left            =   240
         TabIndex        =   31
         Top             =   3600
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Company"
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
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Model"
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
         Left            =   240
         TabIndex        =   29
         Top             =   2400
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Registration NO."
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
         Left            =   240
         TabIndex        =   28
         Top             =   3000
         Width           =   1545
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Description"
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
         TabIndex        =   27
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Change Status"
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
         Left            =   6240
         TabIndex        =   26
         Top             =   5280
         Width           =   1305
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Left            =   360
         TabIndex        =   25
         Top             =   1200
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_viewinformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset


Public Sub vhidfill()
STRSQL = "select a.u_id,a.vh_id,b.vh_id,b.vh_compnam,b.vh_model,b.regi_no,b.chai_no,b.status from tbl_wstatus a inner join tbl_vehicleregistration b on a.vh_id = b.vh_id where a.u_id = '" & mech_id & "'"
Set RS = adocn.Execute(STRSQL)
Do While Not RS.EOF
        comb_vhid.AddItem RS!vh_id
        RS.MoveNext
    Loop
End Sub

Private Sub comb_vhid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_vhid_Click()
 STRSQL = "select * from tbl_vehicleregistration  where vh_id= " _
             & " '" & comb_vhid.List(comb_vhid.ListIndex) & "'"
   Set RS = adocn.Execute(STRSQL)
    txt_comp = RS!vh_compnam
    txt_model = RS!vh_model
    txt_rno = RS!regi_no
    txt_regno.Text = RS!regi_no
End Sub

Private Sub combo_stwork_Change()
MsgBox "Should Select One"
End Sub

Private Sub combo_stwork_Click()
If (combo_stwork.List(combo_stwork.ListIndex) = "BODY WORK") Then
subadd
ElseIf (combo_stwork.List(combo_stwork.ListIndex) = "MECHANICAL REPAIR") Then
subadd2
ElseIf (combo_stwork.List(combo_stwork.ListIndex) = "WHEEL ALLAINGNMENT") Then
subadd3
ElseIf (combo_stwork.List(combo_stwork.ListIndex) = "WATER SERVICE") Then
subadd4
End If
End Sub

'Private Sub subsetgrid()
'gridview.Cols = 4
'gridview.Rows = 2
'gridview.FixedRows = 1
'gridview.TextMatrix(0, 1) = "USER ID"
'gridview.TextMatrix(0, 2) = "VEHICLE ID"
'gridview.TextMatrix(0, 3) = "REGISTOR NO"
'gridview.ColWidth(0) = 0
'gridview.ColWidth(1) = 1000
'gridview.ColWidth(2) = 1730
'gridview.ColWidth(3) = 1730
'End Sub
Public Sub subadd()
STRSQL = " select * from  tbl_bodywork where vh_id = '" & comb_vhid.List(comb_vhid.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
txt_desc = RS!Description
txt_cstatus.Text = RS!Status
If RS!bdywork = "SELECTED" Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If RS!glass = "SELECTED" Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If RS!Paint = "SELECTED" Then
Check3.Value = 1
Else
Check3.Value = 1
End If
If RS!others = "SELECTED" Then
Check4.Value = 1
Else
Check4.Value = 1
End If
Frame3.Visible = False
Frame4.Visible = False
Frame2.Visible = True
Image1.Visible = False
Image2.Visible = True
Image3.Visible = False
Image4.Visible = False
txt_desc.Visible = True
Label7.Visible = True
Else
MsgBox "NOT REGISTERED"
End If
End Sub

Private Sub Command1_Click()
If validation Then
If (combo_stwork.List(combo_stwork.ListIndex) = "BODY WORK") Then
STRSQL = "update tbl_bodywork set status='" & comb_change.List(comb_change.ListIndex) & "' where " _
                    & " vh_id = '" & comb_vhid.List(comb_vhid.ListIndex) & "'"
        adocn.Execute (STRSQL)
        MsgBox "User details updated . . ."
ElseIf (combo_stwork.List(combo_stwork.ListIndex) = "MECHANICAL REPAIR") Then
STRSQL = "update tbl_repare set status='" & comb_change.List(comb_change.ListIndex) & "' where " _
                    & " vh_id = '" & comb_vhid.List(comb_vhid.ListIndex) & "'"
        adocn.Execute (STRSQL)
        MsgBox "User details updated . . ."
ElseIf (combo_stwork.List(combo_stwork.ListIndex) = "WHEEL ALLAINGNMENT") Then
STRSQL = "update tbl_wheel set status='" & comb_change.List(comb_change.ListIndex) & "' where " _
                    & " vh_id = '" & comb_vhid.List(comb_vhid.ListIndex) & "'"
        adocn.Execute (STRSQL)
        MsgBox "User details updated . . ."
ElseIf (combo_stwork.List(combo_stwork.ListIndex) = "WATER SERVICE") Then
STRSQL = "update tbl_water set status ='" & comb_change.List(comb_change.ListIndex) & "' where " _
                    & " vh_id = '" & comb_vhid.List(comb_vhid.ListIndex) & "'"
        adocn.Execute (STRSQL)
        MsgBox "User details updated . . ."
End If
Else
MsgBox "Fill Details"
End If
End Sub

Private Sub Command2_Click()
Unload frm_viewinformation
MDIForm1.Show
End Sub

Private Sub Form_Load()
vhidfill
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
End Sub
'Private Sub subsetgrid2()
'gridview.Cols = 4
'gridview.Rows = 2
'gridview.FixedRows = 1
'gridview.TextMatrix(0, 1) = "DATE"
'gridview.TextMatrix(0, 2) = "MODEL"
'gridview.TextMatrix(0, 3) = "REGISTOR NO"
'gridview.ColWidth(0) = 0
'gridview.ColWidth(1) = 1000
'gridview.ColWidth(2) = 1730
'gridview.ColWidth(3) = 1730
'End Sub
Public Sub subadd2()
STRSQL = "select * from tbl_repare where vh_id='" & comb_vhid.List(comb_vhid.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
txt_desc.Text = RS!Discription
txt_cstatus.Text = RS!Status
If RS!engine = "SELECTED" Then
Check5.Value = 1
Else
Check5.Value = 0
End If
If RS!brake = "SELECTED" Then
Check6.Value = 1
Else
Check6.Value = 0
End If
If RS!oil = "SELECTED" Then
Check7.Value = 1
Else
Check7.Value = 1
End If
If RS!others = "SELECTED" Then
Check8.Value = 1
Else
Check8.Value = 1
End If
Label7.Visible = True
txt_desc.Visible = True
Frame2.Visible = False
Frame4.Visible = False
Frame3.Visible = True
Image1.Visible = False
Image2.Visible = False
Image3.Visible = True
Image4.Visible = False
Else
MsgBox "NOT REGISTERED"
End If
End Sub
'Private Sub subsetgrid3()
'gridview.Cols = 4
'gridview.Rows = 2
'gridview.FixedRows = 1
'gridview.TextMatrix(0, 1) = "DATE"
'gridview.TextMatrix(0, 2) = "MODEL"
'gridview.TextMatrix(0, 3) = "REGISTOR NO"
'gridview.ColWidth(0) = 0
'gridview.ColWidth(1) = 1000
'gridview.ColWidth(2) = 1730
'gridview.ColWidth(3) = 1730
'End Sub
Public Sub subadd3()
STRSQL = "select * from tbl_wheel where vh_id='" & comb_vhid.List(comb_vhid.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
txt_cstatus.Text = RS!Status
Frame3.Visible = False
Frame4.Visible = False
Frame2.Visible = False
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
txt_desc.Visible = False
Label7.Visible = False
Image4.Visible = True
Else
MsgBox "NOT REGISTERED"
End If
End Sub
'Private Sub subsetgrid4()
'gridview.Cols = 4
'gridview.Rows = 2
'gridview.FixedRows = 1
'gridview.TextMatrix(0, 1) = "DATE"
'gridview.TextMatrix(0, 2) = "MODEL"
'gridview.TextMatrix(0, 3) = "REGISTOR NO"
'gridview.ColWidth(0) = 0
'gridview.ColWidth(1) = 1000
'gridview.ColWidth(2) = 1730
'gridview.ColWidth(3) = 1730
'End Sub
Public Sub subadd4()
STRSQL = "select * from tbl_water where vh_id='" & comb_vhid.List(comb_vhid.ListIndex) & "'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
txt_cstatus.Text = RS!Status
If RS!fulbdy = "SELECTED" Then
Check9.Value = 1
Else
Check9.Value = 0
End If
If RS!chase = "SELECTED" Then
Check10.Value = 1
Else
Check10.Value = 0
End If
If RS!others = "SELECTED" Then
Check11.Value = 1
Else
Check11.Value = 1
End If
Frame3.Visible = False
Frame2.Visible = False
Frame4.Visible = True
Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Else
MsgBox "NOT REGISTERED"
End If
End Sub

Private Sub gridview_Click()
If (combo_stwork.List(combo_stwork.ListIndex) = "BODY WORK") Then
fill
ElseIf (combo_stwork.List(combo_stwork.ListIndex) = "MECHANICAL REPAIR") Then
fill2
ElseIf (combo_stwork.List(combo_stwork.ListIndex) = "WHEEL ALLAINGNMENT") Then
fill3
ElseIf (combo_stwork.List(combo_stwork.ListIndex) = "WATER SERVICE") Then
fill4
End If
End Sub
Public Sub fill()
If gridview.Rows > 1 Then
STRSQL = "select a.u_id,a.vh_id,b.vh_id,b.paint,b.bdywork,b.glass,b.others,b.description,b.fdate,b.status,b.vh_compnam,b.vh_model,b.regi_no from tbl_wstatus a inner join tbl_bodywork b on a.vh_id=b.vh_id where a.u_id = '" & mech_id & "'"
Set RS = adocn.Execute(STRSQL)
txt_comp.Text = RS!vh_compnam
txt_model.Text = RS!vh_model
txt_regno.Text = RS!regi_no
txt_desc.Text = RS!Description
txt_cstatus.Text = RS!Status
txt_uid.Text = RS!u_id
If RS!bdywork = "SELECTED" Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If RS!glass = "SELECTED" Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If RS!Paint = "SELECTED" Then
Check3.Value = 1
Else
Check3.Value = 1
End If
If RS!others = "SELECTED" Then
Check4.Value = 1
Else
Check4.Value = 1
End If
End If
End Sub
Public Sub fill2()
If gridview.Rows > 1 Then
STRSQL = "select a.u_id,a.vh_id,a.status,b.vh_id,b.engine,b.brake,b.oil,b.others,b.discription,b.frm_date,b.status,b.vh_company,b.vh_model,b.regi_no from tbl_wstatus a inner join tbl_repare b on a.vh_id=b.vh_id where a.u_id ='" & mech_id & "'"
Set RS = adocn.Execute(STRSQL)
'comb_vid.AddItem = RS!vh_id
txt_comp.Text = RS!vh_compnam
txt_model.Text = RS!vh_model
txt_regno.Text = RS!regi_no
txt_desc.Text = RS!Discription
txt_cstatus.Text = RS!Status
txt_uid.Text = RS!u_id
If RS!engine = "SELECTED" Then
Check5.Value = 1
Else
Check5.Value = 0
End If
If RS!brake = "SELECTED" Then
Check6.Value = 1
Else
Check6.Value = 0
End If
If RS!oil = "SELECTED" Then
Check7.Value = 1
Else
Check7.Value = 1
End If
If RS!others = "SELECTED" Then
Check8.Value = 1
Else
Check8.Value = 1
End If
End If
End Sub
Public Sub fill3()
If gridview.Rows > 1 Then
STRSQL = "select a.u_id,a.vh_id,a.status,b.vh_id,b.t_type,b.f_date,b.status,b.vh_compnam,b.vh_model,b.regi_no from tbl_wstatus a inner join tbl_wheel b on a.vh_id=b.vh_id where a.u_id='" & mech_id & "'"
Set RS = adocn.Execute(STRSQL)
'comb_vid.AddItem = RS!vh_id
txt_comp.Text = RS!vh_compnam
txt_model.Text = RS!vh_model
txt_regno.Text = RS!regi_no
txt_cstatus.Text = RS!Status
txt_uid.Text = RS!u_id
End If
End Sub
Public Sub fill4()
If gridview.Rows > 1 Then
STRSQL = "select a.u_id,a.vh_id,a.status,b.vh_id,b.fulbdy,b.chase,b.others,b.f_date,b.status,b.vh_compnam,b.vh_model,b.regi_no from tbl_wstatus a inner join tbl_water b on a.vh_id=b.vh_id where a.u_id='" & mech_id & "'"
Set RS = adocn.Execute(STRSQL)
txt_comp.Text = RS!vh_compnam
txt_model.Text = RS!vh_model
txt_regno.Text = RS!regi_no
txt_cstatus.Text = RS!Status
txt_uid.Text = RS!u_id
If RS!fulbdy = "SELECTED" Then
Check9.Value = 1
Else
Check9.Value = 0
End If
If RS!chase = "SELECTED" Then
Check10.Value = 1
Else
Check10.Value = 0
End If
If RS!others = "SELECTED" Then
Check11.Value = 1
Else
Check11.Value = 1
End If
End If
End Sub
Public Function validation()
Dim ok As Boolean
If (combo_stwork.Text = "----SELECT ONE------") Then
   lbl_wtyp.Visible = True
ok = False
Else
If (comb_change.Text = "----SELECT ONE------") Then
lbl_cstat.Visible = True
ok = False
Else
ok = True
End If
End If
validation = ok
End Function


