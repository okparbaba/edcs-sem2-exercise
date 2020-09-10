VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dcs Computer Science"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.Data adocmp 
      Caption         =   "adoCompSc"
      Connect         =   "Excel 4.0;"
      DatabaseName    =   "E:\DCSsem2Exercise\ex5.2.2\CompSc.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "Sheet1$"
      Top             =   7320
      Width           =   2895
   End
   Begin VB.PictureBox adoCompSc 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      ScaleHeight     =   315
      ScaleWidth      =   2835
      TabIndex        =   30
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton cmdGoto 
      Caption         =   "Go To"
      Height          =   375
      Left            =   7560
      TabIndex        =   29
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7560
      TabIndex        =   28
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   7560
      TabIndex        =   27
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7560
      TabIndex        =   26
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   7560
      TabIndex        =   25
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   7560
      TabIndex        =   24
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   7560
      TabIndex        =   23
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   7560
      TabIndex        =   22
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txt 
      DataField       =   "SrNo"
      DataSource      =   "adocmp"
      Height          =   375
      Index           =   9
      Left            =   5400
      TabIndex        =   9
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Degree"
      DataSource      =   "adocmp"
      Height          =   375
      Index           =   8
      Left            =   1920
      TabIndex        =   8
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Sdate"
      DataSource      =   "adocmp"
      Height          =   375
      Index           =   7
      Left            =   5400
      TabIndex        =   7
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "DOB"
      DataSource      =   "adocmp"
      Height          =   375
      Index           =   6
      Left            =   1920
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Institute"
      DataSource      =   "adocmp"
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Dept"
      DataSource      =   "adocmp"
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Desig"
      DataSource      =   "adocmp"
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Name_E"
      DataSource      =   "adocmp"
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txt 
      DataField       =   "Name_M"
      DataSource      =   "adocmp"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txt 
      DataField       =   "RN"
      DataSource      =   "adocmp"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "rSwfykHwif"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "tvkyfp0ifaeU"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   960
      X2              =   7080
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   960
      X2              =   7080
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   960
      X2              =   7080
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label10 
      Caption         =   "University of Yangon"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label9 
      Caption         =   "bGJU"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "arG;aeY"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "wuåodkvf"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Xme"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "&&mxl;"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "(English)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "(jrefrm)"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "trnf"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ckHtrSwf"
      BeginProperty Font 
         Name            =   "Win_Researcher"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PassVal = "PDCS-521"



Private Sub cmdCancel_Click()
    adocmp.Recordset.CancelUpdate
    adocmp.Refresh
    Call LockText
    Call OnButtons
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("R U Sure to DELETE?", vbQuestion + vbYesNo, "Confirm") Then
            adocmp.Recordset.Delete
            adocmp.Refresh
    End If
End Sub

Private Sub cmdEdit_Click()
    Dim st As String
    st = InputBox("Enter Password to Edit", "Security")
    If st <> "" Then
        If st = PassVal Then
            Call UnlockText
            Call OffButtons
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdFind_Click()
    Dim Str As String
    Str = InputBox("Enter the Name U want to find", "Searching")
    If Str <> "" Then
            adocmp.Recordset.Filter = "Name_E LIKE '*" & Str & "*'"
            If adocmp.Recordset.EOF Then
                    MsgBox "Not Found"
                    adocmp.Recordset.Filter = ""
            End If
    Else
            adocmp.Recordset.Filter = ""
    End If
End Sub

Private Sub cmdGoto_Click()
    Dim rno
        rno = Val(InputBox("Enter recno to go"))
        If rno >= 1 And rno <= adocmp.Recordset.RecordCount Then
                adocmp.Recordset.MoveFirst
                adocmp.Recordset.Move rno - 1
        Else
                MsgBox "Invalid record number"
        End If

End Sub

Private Sub cmdInsert_Click()
    Call UnlockText
    Call OffButtons
    adocmp.Recordset.AddNew

End Sub

Private Sub cmdSave_Click()
    adocmp.Recordset.UpdateBatch
    adocmp.Refresh
    Call LockText
    Call OnButtons
End Sub

Private Sub Form_Load()
    LockText
End Sub

Private Sub LockText()
    Dim i As Integer
    For i = 0 To 9
        txt(i).Appearance = 0
        txt(i).Locked = True
    Next i
End Sub

Private Sub UnlockText()
    Dim i As Integer
    For i = 0 To 9
        txt(i).Appearance = 1
        txt(i).Locked = False
    Next i
End Sub

Private Sub OffButtons()
    cmdInsert.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdExit.Enabled = False
    cmdFind.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub OnButtons()
    cmdInsert.Enabled = True
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
End Sub

