VERSION 5.00
Begin VB.Form frm_waterbill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Water Bill"
   ClientHeight    =   9645
   ClientLeft      =   4335
   ClientTop       =   870
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   11190
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   8400
      TabIndex        =   15
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10575
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HOTE WHEELS"
         BeginProperty Font 
            Name            =   "FoughtKnight"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   3360
         TabIndex        =   16
         Top             =   240
         Width           =   3285
      End
      Begin VB.Label lbl_tot 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <total amt>"
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
         Left            =   8160
         TabIndex        =   14
         Top             =   7440
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOTAL"
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
         Left            =   240
         TabIndex        =   13
         Top             =   7440
         Width           =   930
      End
      Begin VB.Label lbl_other 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <other amt>"
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
         Left            =   8040
         TabIndex        =   12
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label lbl_chase 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <chase amt>"
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
         Left            =   8040
         TabIndex        =   11
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label lbl_full 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <full amt>"
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
         Left            =   8040
         TabIndex        =   10
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "OTHERS"
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
         TabIndex        =   9
         Top             =   5160
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   7680
         X2              =   7680
         Y1              =   2520
         Y2              =   6960
      End
      Begin VB.Line Line3 
         X1              =   2280
         X2              =   2280
         Y1              =   2520
         Y2              =   6960
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   10560
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CHASE"
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
         TabIndex        =   8
         Top             =   4320
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "FULL BODY"
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
         TabIndex        =   7
         Top             =   3240
         Width           =   1290
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10560
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lbl_compny 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <company>"
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
         Left            =   8160
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lbl_regno 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <reg no>"
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
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lbl_addr 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <address>"
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
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lbl_name 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <name>"
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
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lbl_date 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <date>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8160
         TabIndex        =   2
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "BILL FOR WATER SERVICE"
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
         Left            =   3120
         TabIndex        =   1
         Top             =   1320
         Width           =   3570
      End
   End
End
Attribute VB_Name = "frm_waterbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Dim v3 As Integer

Private Sub Command1_Click()
frm_waterbill.PrintForm
End Sub

Private Sub Form_Load()
LBL_DATE.Caption = DateValue(Now)
STRSQL = "select a.fulbdy,a.chase,a.others,a.tot_amt,b.vh_id,b.name,b.status,b.address,b.mob_no,b.vh_compnam,b.regi_no,b.vh_id,b.vh_model from tbl_watere a inner join tbl_vehicleregistration b on a.vh_id=b.vh_id where b.vh_id='" & VECH_ID & "' and b.status='FINISHED'"
'Set RS = adocn.Execute(STRSQL)
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
lbl_name.Caption = RS!Name
lbl_addr.Caption = RS!address
lbl_regno.Caption = RS!regi_no
lbl_compny.Caption = RS!vh_compnam
lbl_tot.Caption = RS!tot_amt
fill3
End If
End Sub
Private Sub fill3()
s = 0
STRSQL = "select * from tbl_water where vh_id= '" & VECH_ID & "' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
If RS!fulbdy = "SELECTED" Then
v3 = 1
lbl_full.Caption = rate3(v3)
Else
lbl_full.Caption = " "
End If
If RS!chase = "SELECTED" Then
v3 = 2
lbl_chase.Caption = rate3(v3)
Else
lbl_chase.Caption = " "
End If
If RS!others = "SELECTED" Then
v3 = 3
lbl_other.Caption = rate3(v3)
Else
lbl_other.Caption = " "
End If
        RS.MoveNext
    Loop
End Sub
Public Function rate3(v3 As Integer)
Dim s1 As Integer
Dim str As String
Dim RS1 As ADODB.Recordset
str = "select rate from tbl_waterrate where wr_id = '" & v3 & "' "
Set RS1 = adocn.Execute(str)
s1 = RS1!rate
rate3 = s1
RS1.Close
End Function

