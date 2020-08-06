VERSION 5.00
Begin VB.Form frmDtaAccess 
   Caption         =   "Demo for Data & Access Link"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPDCS 
      DataField       =   "SrNo"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   16
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "RNo"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   15
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Name_M"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   14
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Name_E"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   13
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "DOB"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   12
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Degree"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   5
      Left            =   4200
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Major"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   6
      Left            =   960
      TabIndex        =   10
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Desg"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   7
      Left            =   4200
      TabIndex        =   9
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Dept"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   8
      Left            =   960
      TabIndex        =   8
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Sdate"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   9
      Left            =   4200
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Institute"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   10
      Left            =   960
      TabIndex        =   6
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   6480
      Width           =   855
   End
   Begin VB.Data dtaPDCS 
      Caption         =   "DtaPDCS"
      Connect         =   "Access"
      DatabaseName    =   "E:\DCSsem2Exercise\ex5.1.3\CompSc97.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "PDCS$"
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "No"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Name in English"
      Height          =   495
      Left            =   3120
      TabIndex        =   25
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Roll No"
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "DOB"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Degree"
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Major"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Desg"
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Dept"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "SDate"
      Height          =   255
      Left            =   3120
      TabIndex        =   18
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "Institute"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   735
   End
End
Attribute VB_Name = "frmDtaAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
