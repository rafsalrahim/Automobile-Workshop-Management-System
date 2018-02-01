VERSION 5.00
Begin VB.Form frm_login 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fram_login 
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.CommandButton cmd_login 
         Caption         =   "Login"
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
         Left            =   1560
         TabIndex        =   6
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txt_pword 
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txt_uname 
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
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
         TabIndex        =   4
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Username"
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
         TabIndex        =   2
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "   LOGIN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

