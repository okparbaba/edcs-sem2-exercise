VERSION 5.00
Begin VB.Form MSc 
   Caption         =   "MSc Computer Science"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dtaMSc 
      Caption         =   "dtaMSc"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "C:\AssignmentDb\CompSc.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "MSc$"
      Top             =   5880
      Width           =   2655
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "Institute"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   10
      Left            =   3120
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "Sdate"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   9
      Left            =   3120
      TabIndex        =   9
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "Dept"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "Desg"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   7
      Left            =   3120
      TabIndex        =   7
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "Major"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "Degree"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   5
      Left            =   3120
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "DOB"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "Name_E"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "Name_M"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "Rno"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtMSc 
      DataField       =   "SrNo"
      DataSource      =   "dtaMSc"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "MSc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
