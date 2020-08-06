VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cancel"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Data dtaPDCS 
      Caption         =   "DtaPDCS"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   "E:\Assignment\Assignment1\CompSc.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "PDCS$"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Institute"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   10
      Left            =   1320
      TabIndex        =   10
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Sdate"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   9
      Left            =   4560
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Dept"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Desg"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   7
      Left            =   4560
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Major"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Degree"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   5
      Left            =   4560
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "DOB"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Name_E"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   3
      Left            =   4560
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "Name_M"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "RNo"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtPDCS 
      DataField       =   "SrNo"
      DataSource      =   "dtaPDCS"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Institute"
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "SDate"
      Height          =   255
      Left            =   3480
      TabIndex        =   26
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Dept"
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Desg"
      Height          =   255
      Left            =   3480
      TabIndex        =   24
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Major"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Degree"
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "DOB"
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Roll No"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Name in English"
      Height          =   495
      Left            =   3480
      TabIndex        =   19
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "No"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recCount As Integer
Const PassValue = "PDCS-521"

Private Sub LockTextBox()
    For i = 0 To 10
        txtPDCS(i).Locked = True
        txtPDCS(i).Appearance = 0
    Next i
        cmdEdit.Enabled = True
        cmdFind.Enabled = True
        cmdSave.Enabled = False
        cmdCancel.Enabled = False

End Sub

Private Sub UnLockTextBox()
        For i = 0 To 10
                txtPDCS(i).Locked = False
                txtPDCS(i).Appearance = 1
        txtPDCS(i).ForeColor = vbRed
        Next i
        cmdEdit.Enabled = False
        cmdFind.Enabled = False
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
End Sub

Private Sub Form_Load()
    LockTextBox
End Sub

Private Sub cmdCancel_Click()
        dtaPDCS.Recordset.CancelUpdate
        LockTextBox
End Sub
Private Sub cmdSave_Click()
    dtaPDCS.Recordset.Update
    LockTextBox
End Sub

Private Sub cmdClose_Click()
        Dim k As Integer
        k = MsgBox("Do u really want to Quit ? ", vbQuestion + vbYesNo, "Quiting ?")
        If k = vbYes Then
                Unload Me
        Else
                Exit Sub
        End If
End Sub

Private Sub cmdEdit_Click()
        Dim pw As String
        pw = InputBox("Enter Password to EDIT", "Security")
        If pw <> "" Then
                If pw = PassValue Then
                    UnLockTextBox
                    dtaPDCS.Recordset.Edit
                Else
                    MsgBox " Invalid Password", vbExclamation
                End If
        End If
End Sub

Private Sub cmdFind_Click()
        Dim FN As String
        FN = InputBox("Enter the Name You want to find", "Finding")

        If FN <> "" Then
        dtaPDCS.RecordSource = "Select * from [PDCS$]"
        dtaPDCS.Refresh
        dtaPDCS.Recordset.Filter = "Name_E LIKE '*" & FN & "*'"
        Set dtaPDCS.Recordset = dtaPDCS.Recordset.OpenRecordset
        If dtaPDCS.Recordset.EOF Then
                    MsgBox "Not found"
                    dtaPDCS.Recordset.Filter = ""
                End If
        Else
                dtaPDCS.Recordset.Filter = ""
        End If
End Sub


