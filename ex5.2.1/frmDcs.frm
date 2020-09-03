VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Dcs Computer Science"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc adoCompSc 
      Height          =   375
      Left            =   1920
      Top             =   6480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\DCSsem2Exercise\ex5.2.1\CompSc97.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\DCSsem2Exercise\ex5.2.1\CompSc97.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "PDCS"
      Caption         =   "adoCompSc"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      DataSource      =   "adoCompSc"
      Height          =   375
      Index           =   9
      Left            =   5400
      TabIndex        =   9
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Degree"
      DataSource      =   "adoCompSc"
      Height          =   375
      Index           =   8
      Left            =   1920
      TabIndex        =   8
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Sdate"
      DataSource      =   "adoCompSc"
      Height          =   375
      Index           =   7
      Left            =   5400
      TabIndex        =   7
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "DOB"
      DataSource      =   "adoCompSc"
      Height          =   375
      Index           =   6
      Left            =   1920
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Institute"
      DataSource      =   "adoCompSc"
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Dept"
      DataSource      =   "adoCompSc"
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Desig"
      DataSource      =   "adoCompSc"
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txt 
      DataField       =   "Name_E"
      DataSource      =   "adoCompSc"
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txt 
      DataField       =   "Name_M"
      DataSource      =   "adoCompSc"
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
      DataField       =   "RNo"
      DataSource      =   "adoCompSc"
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

Private Sub AdoCompSc_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordSet As ADODB.Recordset)
    adoCompSc.Caption = adoCompSc.Recordset.AbsolutePosition & " of" & adoCompSc.Recordset.RecordCount
    
End Sub

Private Sub cmdCancel_Click()
    adoCompSc.Recordset.CancelUpdate
    adoCompSc.Refresh
    Call LockText
    Call OnButtons
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("R U Sure to DELETE?", vbQuestion + vbYesNo, "Confirm") Then
            adoCompSc.Recordset.Delete
            adoCompSc.Refresh
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
            adoCompSc.Recordset.Filter = "Name_E LIKE '*" & Str & "*'"
            If adoCompSc.Recordset.EOF Then
                    MsgBox "Not Found"
                    adoCompSc.Recordset.Filter = ""
            End If
    Else
            adoCompSc.Recordset.Filter = ""
    End If
End Sub

Private Sub cmdGoto_Click()
    Dim rno
        rno = Val(InputBox("Enter recno to go"))
        If rno >= 1 And rno <= adoCompSc.Recordset.RecordCount Then
                adoCompSc.Recordset.MoveFirst
                adoCompSc.Recordset.Move rno - 1
        Else
                MsgBox "Invalid record number"
        End If

End Sub

Private Sub cmdInsert_Click()
    Call UnlockText
    Call OffButtons
    adoCompSc.Recordset.AddNew

End Sub

Private Sub cmdSave_Click()
    adoCompSc.Recordset.UpdateBatch
    adoCompSc.Refresh
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

