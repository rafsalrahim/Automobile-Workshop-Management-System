VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_bdyrate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rate For Body work"
   ClientHeight    =   6285
   ClientLeft      =   5070
   ClientTop       =   1845
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frm_bdyrate.frx":0000
   ScaleHeight     =   6285
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9015
      Begin VB.CommandButton cmd_cancel 
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
         Left            =   3120
         TabIndex        =   8
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton cmd_update 
         Caption         =   "UPDATE"
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
         Top             =   4800
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid gridrate 
         Height          =   2175
         Left            =   5520
         TabIndex        =   6
         Top             =   2040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3836
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
      Begin VB.TextBox txt_rate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox txt_cname 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label lbl_rate 
         AutoSize        =   -1  'True
         Caption         =   "* Field required"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   225
         Left            =   3960
         TabIndex        =   10
         Top             =   2880
         Width           =   1125
      End
      Begin VB.Label lbl_cat 
         AutoSize        =   -1  'True
         Caption         =   "*Double tap field in grid"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   210
         Left            =   3480
         TabIndex        =   9
         Top             =   2040
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
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
         TabIndex        =   4
         Top             =   3240
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Category Name"
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
         TabIndex        =   2
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RATE FOR BODY WORK"
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
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   3390
      End
   End
End
Attribute VB_Name = "frm_bdyrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub subsetgrid()
gridrate.Cols = 3
gridrate.Rows = 2
gridrate.FixedRows = 1
gridrate.TextMatrix(0, 1) = "CATEGORY"
gridrate.TextMatrix(0, 2) = "RATE"
gridrate.ColWidth(0) = 0
gridrate.ColWidth(1) = 1000
gridrate.ColWidth(2) = 1730
End Sub
Public Sub subaddtogrid()
subsetgrid
STRSQL = "select * from tbl_bdwrkrate"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridrate.TextMatrix(i, 0) = RS!b_id
gridrate.TextMatrix(i, 1) = RS!Type
gridrate.TextMatrix(i, 2) = RS!rate
gridrate.Rows = gridrate.Rows + 1
SLNO = SLNO + 1
RS.MoveNext
i = i + 1
Wend
End If
gridrate.Rows = gridrate.Rows - 1
End Sub

Private Sub cmd_cancel_Click()
cleardata
Form_Load
End Sub

Private Sub cmd_update_Click()
If fnValidation Then
STRSQL = "update tbl_bdwrkrate set Type='" & txt_cname.Text & "',Rate='" & txt_rate.Text & "' where b_id= '" & gridrate.TextMatrix(gridrate.RowSel, 0) & "'"
        adocn.Execute (STRSQL)
        MsgBox " details updated . . ."
        subaddtogrid
        cleardata
       Unload Me
MDIForm1.Show
Else
MsgBox "Field Required"
End If
End Sub

Private Sub Form_Load()
subaddtogrid
clearlabel
cleardata
End Sub

Private Sub gridrate_Click()
If gridrate.Rows > 1 Then
STRSQL = "select * from tbl_bdwrkrate where  b_id = '" & gridrate.TextMatrix(gridrate.RowSel, 0) & "'"
Set RS = adocn.Execute(STRSQL)
txt_cname.Text = RS!Type
txt_rate.Text = RS!rate
End If
End Sub
Public Sub cleardata()
txt_cname.Text = " "
txt_rate.Text = " "
End Sub

Private Sub txt_cname_Change()
 If Trim(txt_cname.Text) = "" Then
        lbl_cat.Visible = True
    Else
        lbl_cat.Visible = False
    End If
End Sub
Private Sub txt_rate_Change()
 If Trim(txt_rate.Text) = "" Then
        lbl_rate.Visible = True
    Else
        lbl_rate.Visible = False
    End If
End Sub
Public Sub clearlabel()
    lbl_rate.Visible = False
    lbl_cat.Visible = False
      End Sub
Public Function fnValidation()
Dim ok As Boolean
If (Trim(txt_rate.Text) = "") Then
   lbl_rate.Visible = True
ok = False
ElseIf (Trim(txt_cname.Text) = "") Then
  lbl_cat.Visible = True
    ok = False
    Else
    ok = True
    End If
    fnValidation = ok
End Function
