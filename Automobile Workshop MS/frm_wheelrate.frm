VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_wheelrate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set  Wheel Allignment Rate"
   ClientHeight    =   5955
   ClientLeft      =   2790
   ClientTop       =   2205
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton cmd_cancel 
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
         Left            =   3840
         TabIndex        =   8
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmd_update 
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
         Left            =   1920
         TabIndex        =   7
         Top             =   3960
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid gridrate 
         Height          =   1935
         Left            =   5640
         TabIndex        =   6
         Top             =   1920
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3413
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
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
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txt_cname 
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
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lbl_check 
         AutoSize        =   -1  'True
         Caption         =   "Tap Grid"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   4440
         TabIndex        =   9
         Top             =   1680
         Width           =   690
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
         Left            =   840
         TabIndex        =   4
         Top             =   2760
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
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RATE FOR WHEEL ALLIGNMENT"
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
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   4665
      End
   End
End
Attribute VB_Name = "FRM_WHEELRATE"
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
gridrate.ColWidth(2) = 1000
End Sub
Public Sub subaddtogrid()
subsetgrid
STRSQL = "select * from tbl_wheelrate"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridrate.TextMatrix(i, 0) = RS!tr_id
gridrate.TextMatrix(i, 1) = RS!service
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
End Sub

Private Sub cmd_update_Click()
If fnValidation Then
STRSQL = "update tbl_wheelrate set service='" & txt_cname.Text & "',Rate='" & txt_rate.Text & "' where  tr_id = '" & gridrate.TextMatrix(gridrate.RowSel, 0) & "'"
        adocn.Execute (STRSQL)
        MsgBox "User details updated . . ."
        subaddtogrid
        cleardata
        Else
        MsgBox "Field Required"
        End If
End Sub

Private Sub Form_Load()
subaddtogrid
lbl_check.Visible = False
End Sub

Private Sub gridrate_Click()
If gridrate.Rows > 1 Then
STRSQL = "select * from tbl_wheelrate where  tr_id = '" & gridrate.TextMatrix(gridrate.RowSel, 0) & "'"
Set RS = adocn.Execute(STRSQL)
txt_cname.Text = RS!service
txt_rate.Text = RS!rate
End If
End Sub
Public Sub cleardata()
txt_cname.Text = " "
txt_rate.Text = " "
End Sub

Private Sub txt_cname_LostFocus()
lbl_check.Visible = True
End Sub
Public Function fnValidation()
Dim ok1 As Boolean
If (Trim(txt_cname.Text) = "") Then
ok1 = False
ElseIf (Trim(txt_rate.Text) = "") Then
  ok1 = False
  Else
  ok1 = True
  End If
  fnValidation = ok1
End Function

