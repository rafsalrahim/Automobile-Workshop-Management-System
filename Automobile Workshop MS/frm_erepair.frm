VERSION 5.00
Begin VB.Form frm_erepair 
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   7725
   Begin VB.Frame fram_erepair 
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txt_vmodel 
         Height          =   375
         Left            =   2880
         TabIndex        =   19
         Top             =   3000
         Width           =   3015
      End
      Begin VB.TextBox txt_vcomp 
         Height          =   405
         Left            =   2880
         TabIndex        =   17
         Top             =   2520
         Width           =   3015
      End
      Begin VB.CommandButton cmd_submit 
         Caption         =   "SUBMIT"
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
         Left            =   4080
         TabIndex        =   15
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancal 
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
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   6720
         Width           =   1335
      End
      Begin VB.TextBox txt_tdate 
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   6000
         Width           =   3015
      End
      Begin VB.TextBox txt_fdate 
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   5400
         Width           =   3015
      End
      Begin VB.TextBox txt_work 
         Height          =   975
         Left            =   2880
         TabIndex        =   9
         Top             =   4080
         Width           =   3015
      End
      Begin VB.TextBox txt_rno 
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox txt_name 
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   1920
         Width           =   3015
      End
      Begin VB.ComboBox comb_vid 
         Height          =   315
         Left            =   2880
         TabIndex        =   3
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle model"
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
         TabIndex        =   18
         Top             =   3120
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle company"
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
         TabIndex        =   16
         Top             =   2520
         Width           =   1560
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Left            =   960
         TabIndex        =   12
         Top             =   6000
         Width           =   750
      End
      Begin VB.Label Label7 
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
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Top             =   5400
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Left            =   960
         TabIndex        =   8
         Top             =   4320
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Registration No."
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
         TabIndex        =   6
         Top             =   3600
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "ELECTRICAL WORKS"
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
         Left            =   2280
         TabIndex        =   1
         Top             =   480
         Width           =   3045
      End
   End
End
Attribute VB_Name = "frm_erepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

