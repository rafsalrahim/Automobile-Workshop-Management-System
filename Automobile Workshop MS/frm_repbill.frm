VERSION 5.00
Begin VB.Form frm_repbill 
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   375
         Left            =   4800
         TabIndex        =   10
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lbl_enginant 
         Caption         =   "Label3"
         Height          =   375
         Left            =   4920
         TabIndex        =   8
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Engine"
         Height          =   315
         Left            =   600
         TabIndex        =   7
         Top             =   3360
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10440
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label lbl_vid 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label lbl_addr 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lbl_name 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lbl_mod 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   495
         Left            =   8400
         TabIndex        =   3
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lbl_date 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   495
         Left            =   8400
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "BILL FOR MECHANICAL WORK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   4620
      End
   End
End
Attribute VB_Name = "frm_repbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
