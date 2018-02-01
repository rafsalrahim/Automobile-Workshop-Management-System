VERSION 5.00
Begin VB.Form frm_wwheelbill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wheel Bill"
   ClientHeight    =   8865
   ClientLeft      =   3705
   ClientTop       =   2025
   ClientWidth     =   8085
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
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
      Left            =   6000
      TabIndex        =   6
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.Label Label3 
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
         Left            =   2040
         TabIndex        =   11
         Top             =   360
         Width           =   3285
      End
      Begin VB.Label lbl_company 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <company>"
         Height          =   495
         Left            =   5640
         TabIndex        =   10
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lbl_model 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <model>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2640
         TabIndex        =   9
         Top             =   3840
         Width           =   2280
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   7800
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Label lbl_tot 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <totamt>"
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOTAL AMT"
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   6720
         Width           =   1320
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7800
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label LBL_DATE 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <date>"
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lbl_regno 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <regno>"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbl_addr 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <addr>"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lbl_name 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <name>"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "WHEEL ALLIGNMENT"
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
         Left            =   2040
         TabIndex        =   1
         Top             =   1560
         Width           =   2835
      End
   End
End
Attribute VB_Name = "frm_wwheelbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset

Private Sub Command1_Click()
frm_wwheelbill.PrintForm
End Sub

Private Sub Form_Load()
LBL_DATE.Caption = DateValue(Now)
STRSQL = "select a.amt,b.vh_id,b.name,b.status,b.address,b.mob_no,b.vh_compnam,b.vh_model,b.regi_no from tbl_wheelbill a inner join tbl_vehicleregistration b on a.vh_id=b.vh_id where b.vh_id='" & VECH_ID & "' and b.status='FINISHED'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
lbl_name.Caption = RS!Name
lbl_addr.Caption = RS!address
lbl_regno.Caption = RS!regi_no
lbl_company.Caption = RS!vh_compnam
lbl_model.Caption = RS!vh_model
lbl_tot.Caption = RS!amt
End If
End Sub
