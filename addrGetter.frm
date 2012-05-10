VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Address"
   ClientHeight    =   5085
   ClientLeft      =   10860
   ClientTop       =   7365
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cities 
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Text            =   "(Select)"
      Top             =   2880
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3000
      Top             =   4440
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   $"addrGetter.frx":0000
      OLEDBString     =   $"addrGetter.frx":019E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.ComboBox states 
      Height          =   315
      Left            =   2280
      TabIndex        =   12
      Text            =   "(Select)"
      Top             =   3360
      Width           =   2055
   End
   Begin MSMask.MaskEdBox pin 
      Height          =   285
      Left            =   2280
      TabIndex        =   10
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1793
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox locality 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox street 
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox house 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   120
      Picture         =   "addrGetter.frx":033C
      ScaleHeight     =   4635
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "State"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Pin"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "City"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Locality"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Street"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "House"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Hide
Form2.address.Text = house.Text & " " & street.Text & "," & locality.Text & "," & cities.Text & " - " & pin.Text & ", " & states.Text
End Sub

Private Sub Form_Load()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT State FROM Cities GROUP BY State ORDER BY State"
Adodc1.Refresh
Dim i As Integer
For i = 0 To Adodc1.Recordset.RecordCount - 1
states.AddItem Adodc1.Recordset.Fields("State")
Adodc1.Recordset.Move (1)
Next i
Adodc1.Recordset.MoveFirst
End Sub

Private Sub states_Click()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT City FROM Cities WHERE State = '" & states.Text & "' ORDER BY City"
Adodc1.Refresh
cities.Clear
Dim j As Integer
For j = 0 To Adodc1.Recordset.RecordCount - 1
cities.AddItem Adodc1.Recordset.Fields("City")
Adodc1.Recordset.Move (1)
Next j
Adodc1.Recordset.MoveFirst
End Sub
