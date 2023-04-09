VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   12120
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Studentdb 
      Height          =   855
      Left            =   840
      Top             =   6960
      Visible         =   0   'False
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1508
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
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":00BA
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "student_info"
      Caption         =   "Student Database"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Fax"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtphone 
      DataField       =   "Phone:"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   2760
      TabIndex        =   20
      Top             =   5040
      Width           =   4935
   End
   Begin VB.CommandButton lastbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "LAST"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton nextbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton prevbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton firstbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "FIRST"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton updatebtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "UPDATE RECORD"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cancelbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "CLEAR RECORD"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton delbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "DELETE RECORD"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton addbtn 
      BackColor       =   &H0080C0FF&
      Caption         =   "ADD RECORD"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtemail 
      DataField       =   "Email:"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   2760
      TabIndex        =   10
      Top             =   4320
      Width           =   4935
   End
   Begin VB.TextBox txtadd 
      DataField       =   "Address:"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2760
      TabIndex        =   9
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox txtclass 
      DataField       =   "Class:"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2760
      TabIndex        =   8
      Top             =   2880
      Width           =   4935
   End
   Begin VB.TextBox txtname 
      DataField       =   "Name:"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   2160
      Width           =   4935
   End
   Begin VB.TextBox txtroll 
      DataField       =   "Roll No:"
      DataSource      =   "Studentdb"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   495
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Class:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Roll no:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT REGISTRATION SYSTEM"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   9855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderWidth     =   4
      Height          =   8175
      Left            =   0
      Top             =   0
      Width           =   12135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub addbtn_Click(Index As Integer)
Studentdb.Recordset.AddNew
End Sub

Private Sub cancelbtn_Click(Index As Integer)
txtroll.Text = ""
txtname.Text = ""
txtclass.Text = ""
txtadd.Text = ""
txtemail.Text = ""
txtphone.Text = ""
End Sub

Private Sub firstbtn_Click(Index As Integer)
Studentdb.Recordset.MoveFirst
End Sub

Private Sub lastbtn_Click(Index As Integer)
Studentdb.Recordset.MoveLast
End Sub

Private Sub nextbtn_Click(Index As Integer)
Studentdb.Recordset.MoveNext
End Sub

Private Sub prevbtn_Click(Index As Integer)
Studentdb.Recordset.MovePrevious
End Sub

Private Sub updatebtn_Click(Index As Integer)
Studentdb.Recordset.Update
MsgBox "Data is saved successfully", vbInformation, "Message"
End Sub
