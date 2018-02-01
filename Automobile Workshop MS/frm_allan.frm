VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_allan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wheel Allignent"
   ClientHeight    =   6555
   ClientLeft      =   4215
   ClientTop       =   2085
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fram_allan 
      BackColor       =   &H80000016&
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin MSFlexGridLib.MSFlexGrid gridvhservice 
         Height          =   3375
         Left            =   6000
         TabIndex        =   16
         Top             =   1440
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5953
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   4440
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110493697
         CurrentDate     =   42195
      End
      Begin VB.CommandButton cmd_submit 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cancal 
         Caption         =   "CANCAL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   5400
         Width           =   1335
      End
      Begin VB.TextBox txt_fdate 
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
         TabIndex        =   11
         Top             =   3720
         Width           =   2175
      End
      Begin VB.ComboBox comb_vid 
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
         Left            =   2880
         TabIndex        =   9
         Text            =   "__Select One__"
         Top             =   1440
         Width           =   2175
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
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2520
         Width           =   2175
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
         Left            =   2880
         TabIndex        =   5
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txt_rno 
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
         TabIndex        =   3
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VEHICLE ON SERVICE"
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
         Left            =   6480
         TabIndex        =   17
         Top             =   600
         Width           =   3150
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         Left            =   720
         TabIndex        =   12
         Top             =   4560
         Width           =   750
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Left            =   720
         TabIndex        =   10
         Top             =   3840
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   720
         TabIndex        =   8
         Top             =   3120
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   840
         TabIndex        =   6
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   840
         TabIndex        =   4
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         Left            =   840
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WHEEL ALLAINMENT"
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
         Left            =   2055
         TabIndex        =   1
         Top             =   600
         Width           =   3105
      End
   End
End
Attribute VB_Name = "frm_allan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Dim STRSQL1 As String
Dim RS1 As ADODB.Recordset
Dim tdate As Date
Dim a As String
Dim e As Date
Public Sub combofill()
tdate = DateValue(Now)
STRSQL = "select * from tbl_vehicleregistration where date = '" & tdate & "' and status='REGISTERED' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb_vid.AddItem RS!vh_id
        RS.MoveNext
    Loop
End Sub
Private Sub cmd_cancal_Click()
subcleardata
MsgBox "Canseld. . ."
End Sub
Public Sub subcleardata()
txt_comp.Text = " "
txt_model.Text = " "
txt_rno.Text = " "
End Sub

Private Sub cmd_submit_Click()
If validation Then
STRSQL1 = "select * from tbl_wheel where vh_id='" & comb_vid.List(comb_vid.ListIndex) & "'"
Set RS1 = adocn.Execute(STRSQL1)
If RS1.RecordCount < 1 Then
insert
MsgBox "Registration successfull"
cleardata
Unload Me
MDIForm1.Show
Else
MsgBox "Already Registered"
Unload Me
MDIForm1.Show
End If
Else
MsgBox "Field Needed"
End If
End Sub

Private Sub comb_vid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_vid_Click()
subAutofill
End Sub

Private Sub Form_Load()
combofill
txt_fdate = DateValue(Now)
subaddtogrid
'subaddtogrid1
End Sub
Private Sub subAutofill()
    STRSQL = "select vh_compnam,vh_model,regi_no from tbl_vehicleregistration  where vh_id= " _
             & " '" & comb_vid.List(comb_vid.ListIndex) & "'"
   Set RS = adocn.Execute(STRSQL)
    txt_comp = RS!vh_compnam
    txt_model = RS!vh_model
    txt_rno = RS!regi_no

End Sub
Private Sub insert()
e = DTPicker1.Value

STRSQL = "insert into tbl_wheel(t_date,f_date,status,vh_compnam,vh_model,regi_no,vh_id)values " _
        & " ( '" & e & "' , '" & txt_fdate & "' ,' REGISTREAD ' ,'" & txt_comp & "','" & txt_model & "','" & txt_rno & "','" & comb_vid.List(comb_vid.ListIndex) & "')"
    Set RS = adocn.Execute(STRSQL)
End Sub
Private Sub subsetgrid()
gridvhservice.Cols = 3
gridvhservice.Rows = 2
gridvhservice.FixedRows = 1
gridvhservice.TextMatrix(0, 1) = "COMPANY"
gridvhservice.TextMatrix(0, 2) = "REGI. NO"
gridvhservice.ColWidth(0) = 0
gridvhservice.ColWidth(1) = 1000
gridvhservice.ColWidth(2) = 1730
End Sub
Public Sub subaddtogrid()
subsetgrid
STRSQL = "select * from tbl_wheel"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridvhservice.TextMatrix(i, 0) = RS!t_id
gridvhservice.TextMatrix(i, 1) = RS!vh_compnam
gridvhservice.TextMatrix(i, 2) = RS!regi_no
gridvhservice.Rows = gridvhservice.Rows + 1
SLNO = SLNO + 1
RS.MoveNext
i = i + 1
Wend
End If
gridvhservice.Rows = gridvhservice.Rows - 1
End Sub

'Private Sub subsetgrid1()
'griddeleverd.Cols = 3
'griddeleverd.Rows = 2
'griddeleverd.FixedRows = 1
'griddeleverd.TextMatrix(0, 1) = "COMPANY"
'griddeleverd.TextMatrix(0, 2) = "REGI. NO"
'griddeleverd.ColWidth(0) = 0
'griddeleverd.ColWidth(1) = 1000
'griddeleverd.ColWidth(2) = 1730
'End Sub
'Public Sub subaddtogrid1()
'subsetgrid1
'STRSQL = "select * from tbl_wheel where status= 'FINISHED' "
'Set RS = adocn.Execute(STRSQL)
'If RS.RecordCount > 0 Then
'i = 1
'SLNO = 1
'While Not RS.EOF
'griddeleverd.TextMatrix(i, 0) = RS!t_id
'griddeleverd.TextMatrix(i, 1) = RS!vh_compnam
'griddeleverd.TextMatrix(i, 2) = RS!regi_no
'griddeleverd.Rows = griddeleverd.Rows + 1
'SLNO = SLNO + 1
'RS.MoveNext
'i = i + 1
'Wend
'End If
'griddeleverd.Rows = griddeleverd.Rows - 1
'End Sub



Private Sub gridvhservice_Click()
If gridvhservice.Rows > 1 Then
'cmd_update.Enabled = True
'cmd_delete.Enabled = True
'cmd_submit.Enabled = False
STRSQL = "select * from tbl_wheel where  t_id = '" & gridvhservice.TextMatrix(gridvhservice.RowSel, 0) & "'"
Set RS = adocn.Execute(STRSQL)
txt_comp.Text = RS!vh_compnam
txt_model.Text = RS!vh_model
txt_rno.Text = RS!regi_no
txt_fdate.Text = RS!f_date
comb_vid.Text = RS!vh_id
comb_typ.Text = RS!t_type
comb_status.Text = RS!Status
DTPicker1.Value = RS!t_date
        End If
End Sub


'Private Sub txt_comp_Change()
' If Trim(txt_comp.Text) = "" Then
'        lbl_comp.Visible = True
'    Else
'        lbl_comp.Visible = False
'    End If
'End Sub
'Private Sub txt_model_Change()
' If Trim(txt_model.Text) = "" Then
'        lbl_typr.Visible = True
'    Else
'        lbl_typr.Visible = False
'    End If
'End Sub
'Private Sub txt_rno_Change()
' If Trim(txt_rno.Text) = "" Then
'        lbl_regno.Visible = True
'    Else
'        lbl_regno.Visible = False
'    End If
'End Sub
'Public Sub clearlabel()
'    lbl_modl.Visible = False
'    lbl_regno.Visible = False
'    lbl_comp.Visible = False
'   End Sub

'Public Function fnValidation()
'Dim ok As Boolean
'If (Trim(txt_comp.Text) = "") Then
'   lbl_comp.Visible = True
'ok = False
'Else
' If (Trim(txt_model.Text) = "") Then
'  lbl_modl.Visible = True
'    ok = False
'    Else
'    If (Trim(txt_rno.Text) = "") Then
'     lbl_regno.Visible = True
'    ok = False
''    Else
''    If (Trim(txt_vmodel.Text) = "") Then
''     lbl_vhmod.Visible = True
''     ok = False
''     Else
''     If (Trim(txt_rno.Text) = "") Then
''       lbl_regno.Visible = True
''     ok = False
''     Else
''     If (Trim(txt_cheno.Text) = "") Then
''        lbl_chase.Visible = True
''     ok = False
''     Else
''    ok = True
''    End If
''    End If
''    End If
'    End If
'    End If
'    'End If
'    'End If
'    End If
'    fnValidation = ok
'End Function
Public Sub cleardata()
txt_comp.Text = ""
txt_model.Text = ""
txt_rno.Text = ""

End Sub
Public Function validation()
Dim ok As Boolean
If (comb_vid.Text = "__Select One__") Then
ok = False
Else
ok = True
End If
validation = ok
End Function

