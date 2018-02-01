VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_staupdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Status"
   ClientHeight    =   6435
   ClientLeft      =   3525
   ClientTop       =   2205
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   13785
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      Begin VB.CheckBox Check4 
         Caption         =   "Check4"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   12000
         TabIndex        =   21
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Enabled         =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   8880
         TabIndex        =   20
         Top             =   3480
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   12000
         TabIndex        =   19
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8880
         TabIndex        =   18
         Top             =   1320
         Width           =   255
      End
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
         Height          =   615
         Left            =   3600
         TabIndex        =   17
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   16
         Top             =   5040
         Width           =   1215
      End
      Begin VB.ComboBox comb_upstat 
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
         ItemData        =   "frm_staupdate.frx":0000
         Left            =   3000
         List            =   "frm_staupdate.frx":0007
         TabIndex        =   15
         Text            =   "-SELECT ONE-"
         Top             =   3480
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid gridwater 
         Height          =   1215
         Left            =   9720
         TabIndex        =   13
         Top             =   4080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridwheel 
         Height          =   1215
         Left            =   6840
         TabIndex        =   12
         Top             =   4080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridmechanic 
         Height          =   1215
         Left            =   9720
         TabIndex        =   9
         Top             =   1800
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridbdwork 
         Height          =   1215
         Left            =   6840
         TabIndex        =   7
         Top             =   1800
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
         _Version        =   393216
      End
      Begin VB.TextBox txt_cstat 
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
         TabIndex        =   5
         Top             =   2280
         Width           =   2055
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
         Left            =   3000
         TabIndex        =   3
         Text            =   "------SELECT ONE------"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   2280
         TabIndex        =   22
         Top             =   3360
         Width           =   150
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Update Status"
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
         TabIndex        =   14
         Top             =   3480
         Width           =   1290
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "WATER SERVICE"
         Height          =   195
         Left            =   10080
         TabIndex        =   11
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "WHEEL ALLIGNMENT"
         Height          =   195
         Left            =   7080
         TabIndex        =   10
         Top             =   3480
         Width           =   1665
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "MACHANICAL WORK"
         Height          =   195
         Left            =   10080
         TabIndex        =   8
         Top             =   1320
         Width           =   1590
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "BODY WORK"
         Height          =   195
         Left            =   7320
         TabIndex        =   6
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Current status"
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
         Left            =   960
         TabIndex        =   4
         Top             =   2280
         Width           =   1260
      End
      Begin VB.Label Label2 
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
         Left            =   960
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "STATUS UPDATION"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   2760
      End
   End
End
Attribute VB_Name = "frm_staupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Dim STRSQL1 As String
Dim RS1 As ADODB.Recordset
Dim STRSQL2 As String
Dim RS2 As ADODB.Recordset

Private Sub comb_upstat_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_vhid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_vhid_Click()
subAutofill
check
End Sub
Private Sub subAutofill()
    STRSQL = "select status from tbl_vehicleregistration  where vh_id= " _
             & " '" & comb_vhid.List(comb_vhid.ListIndex) & "'"
   Set RS = adocn.Execute(STRSQL)
    txt_cstat = RS!Status
  End Sub

Private Sub Command1_Click()
If validation Then
STRSQL1 = "update tbl_vehicleregistration set status= '" & comb_upstat.List(comb_upstat.ListIndex) & "' where vh_id = '" & comb_vhid.List(comb_vhid.ListIndex) & "' "
Set RS1 = adocn.Execute(STRSQL1)
        MsgBox "User details updated . . ."
        workser
 Else
 MsgBox "Field Required"
 End If
End Sub

Private Sub subsetgrid()
gridbdwork.Cols = 4
gridbdwork.Rows = 2
gridbdwork.FixedRows = 1
gridbdwork.TextMatrix(0, 1) = "VEHICLE ID"
gridbdwork.TextMatrix(0, 2) = "STATUS"
gridbdwork.TextMatrix(0, 3) = "REGI. NO"
gridbdwork.ColWidth(0) = 0
gridbdwork.ColWidth(1) = 1000
gridbdwork.ColWidth(2) = 1730
gridbdwork.ColWidth(2) = 1730
End Sub
Public Sub subaddtogrid()
subsetgrid
STRSQL = "select * from tbl_bodywork where status= 'FINISHED' "
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridbdwork.TextMatrix(i, 0) = RS!cust_id2
gridbdwork.TextMatrix(i, 1) = RS!vh_id
gridbdwork.TextMatrix(i, 2) = RS!Status
gridbdwork.TextMatrix(i, 3) = RS!regi_no
gridbdwork.Rows = gridbdwork.Rows + 1
SLNO = SLNO + 1
RS.MoveNext
i = i + 1
Wend
End If
gridbdwork.Rows = gridbdwork.Rows - 1
End Sub

Private Sub subsetgrid1()
gridmechanic.Cols = 4
gridmechanic.Rows = 2
gridmechanic.FixedRows = 1
gridmechanic.TextMatrix(0, 1) = "VEHICLE ID"
gridmechanic.TextMatrix(0, 2) = "STATUS"
gridmechanic.TextMatrix(0, 3) = "REGI. NO"
gridmechanic.ColWidth(0) = 0
gridmechanic.ColWidth(1) = 1000
gridmechanic.ColWidth(2) = 1730
gridmechanic.ColWidth(2) = 1730
End Sub
Public Sub subaddtogrid1()
subsetgrid1
STRSQL = "select * from tbl_repare where status= 'FINISHED' "
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridmechanic.TextMatrix(i, 0) = RS!rid
gridmechanic.TextMatrix(i, 1) = RS!vh_id
gridmechanic.TextMatrix(i, 2) = RS!Status
gridmechanic.TextMatrix(i, 3) = RS!regi_no
gridmechanic.Rows = gridmechanic.Rows + 1
SLNO = SLNO + 1
RS.MoveNext
i = i + 1
Wend
End If
gridmechanic.Rows = gridmechanic.Rows - 1
End Sub
Private Sub subsetgrid2()
gridwheel.Cols = 4
gridwheel.Rows = 2
gridwheel.FixedRows = 1
gridwheel.TextMatrix(0, 1) = "VEHICLE ID"
gridwheel.TextMatrix(0, 2) = "STATUS"
gridwheel.TextMatrix(0, 3) = "REGI. NO"
gridwheel.ColWidth(0) = 0
gridwheel.ColWidth(1) = 1000
gridwheel.ColWidth(2) = 1730
gridwheel.ColWidth(2) = 1730
End Sub
Public Sub subaddtogrid2()
subsetgrid2
STRSQL = "select * from tbl_wheel where status= 'FINISHED' "
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridwheel.TextMatrix(i, 0) = RS!t_id
gridwheel.TextMatrix(i, 1) = RS!vh_id
gridwheel.TextMatrix(i, 2) = RS!Status
gridwheel.TextMatrix(i, 3) = RS!regi_no
gridwheel.Rows = gridwheel.Rows + 1
SLNO = SLNO + 1
RS.MoveNext
i = i + 1
Wend
End If
gridwheel.Rows = gridwheel.Rows - 1
End Sub
Private Sub subsetgrid3()
gridwater.Cols = 4
gridwater.Rows = 2
gridwater.FixedRows = 1
gridwater.TextMatrix(0, 1) = "VEHICLE ID"
gridwater.TextMatrix(0, 2) = "STATUS"
gridwater.TextMatrix(0, 3) = "REGI. NO"
gridwater.ColWidth(0) = 0
gridwater.ColWidth(1) = 1000
gridwater.ColWidth(2) = 1730
gridwater.ColWidth(2) = 1730
End Sub
Public Sub subaddtogrid3()
subsetgrid3
STRSQL = "select * from tbl_water where status= 'FINISHED' "
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridwater.TextMatrix(i, 0) = RS!w_id
gridwater.TextMatrix(i, 1) = RS!vh_id
gridwater.TextMatrix(i, 2) = RS!Status
gridwater.TextMatrix(i, 3) = RS!regi_no
gridwater.Rows = gridwater.Rows + 1
SLNO = SLNO + 1
RS.MoveNext
i = i + 1
Wend
End If
gridwater.Rows = gridwater.Rows - 1
End Sub

Private Sub Command2_Click()
Unload Me
MDIForm1.Show
End Sub

Private Sub Form_Load()
combofill
subaddtogrid
subaddtogrid1
subaddtogrid2
subaddtogrid3
End Sub
Public Sub combofill()
STRSQL = "select * from tbl_vehicleregistration where status ='ON PROCESS' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb_vhid.AddItem RS!vh_id
        RS.MoveNext
    Loop
End Sub
Public Sub check()
Dim s As String
Dim t As String
Dim r As String
Dim q As String
Dim RS1 As ADODB.Recordset
Dim RS2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
s = "Select * from tbl_bodywork where vh_id= '" & comb_vhid.List(comb_vhid.ListIndex) & "' "
Set RS1 = adocn.Execute(s)
If RS1.RecordCount > 0 Then
Check1.Visible = True
Check1.Value = 1
Else
Check1.Visible = False
End If
RS1.Close
t = "Select * from tbl_repare where vh_id= '" & comb_vhid.List(comb_vhid.ListIndex) & "'"
Set RS2 = adocn.Execute(t)
If RS2.RecordCount > 0 Then
Check2.Visible = True
Check2.Value = 1
Else
Check2.Visible = False
End If
RS2.Close
r = "Select * from tbl_wheel where vh_id= '" & comb_vhid.List(comb_vhid.ListIndex) & "'"
Set rs3 = adocn.Execute(r)
If rs3.RecordCount > 0 Then
Check3.Visible = True
Check3.Value = 1
Else
Check3.Visible = False
End If
rs3.Close
q = "Select * from tbl_water where vh_id= '" & comb_vhid.List(comb_vhid.ListIndex) & "'"
Set rs4 = adocn.Execute(q)
If rs4.RecordCount > 0 Then
Check4.Visible = True
Check4.Value = 1
Else
Check4.Visible = False
End If
rs4.Close
End Sub
Public Sub workser()
STRSQL2 = "update tbl_wstatus set status= 'FINISHED' "
Set RS2 = adocn.Execute(STRSQL1)
End Sub
Public Function validation()
Dim ok As Boolean
If (comb_upstat.Text = "------SELECT ONE------") Then
 
ok = False
Else
If (comb_vhid.Text = "------SELECT ONE------") Then
  
ok = False
Else
ok = True
End If
End If
validation = ok
End Function

