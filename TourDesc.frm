VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Describe your Tour"
   ClientHeight    =   5085
   ClientLeft      =   8295
   ClientTop       =   4710
   ClientWidth     =   7890
   Icon            =   "TourDesc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   150
   ScaleMode       =   0  'User
   ScaleWidth      =   200
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "TourDesc.frx":000C
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4440
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   $"TourDesc.frx":0021
      OLEDBString     =   $"TourDesc.frx":01BF
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Tours"
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
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      ToolTipText     =   "List all travellers"
      Top             =   2040
      Width           =   255
   End
   Begin VB.CheckBox fdCheck 
      Caption         =   "Food"
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.CheckBox ssCheck 
      Caption         =   "Sightseeing"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Proceed to Booking >>"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   4200
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   120
      Picture         =   "TourDesc.frx":035D
      ScaleHeight     =   4635
      ScaleWidth      =   915
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox chilCombo 
      Height          =   315
      Left            =   2400
      TabIndex        =   11
      Top             =   2280
      Width           =   855
   End
   Begin VB.ComboBox adulCombo 
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox yearCombo 
      Height          =   315
      Left            =   4440
      TabIndex        =   9
      Text            =   "Year"
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox monthCombo 
      Height          =   315
      Left            =   3360
      TabIndex        =   8
      Text            =   "Month"
      Top             =   1320
      Width           =   975
   End
   Begin VB.ComboBox dateCombo 
      Height          =   315
      ItemData        =   "TourDesc.frx":72D2
      Left            =   2400
      List            =   "TourDesc.frx":72D4
      TabIndex        =   7
      Text            =   "Date"
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox tourList 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   4560
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label nightLabel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label descLabel 
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Children"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Adults"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Departure"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Tour"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Form5.Show
End Sub

Sub loadTours()
Dim i As Integer
For i = 0 To Adodc1.Recordset.RecordCount - 1
tourList.AddItem Adodc1.Recordset.Fields("Tour")
Adodc1.Recordset.Move (1)
Next i
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Form_Load()
loadTours
For i = 1 To 31
dateCombo.AddItem (i)

Next i

Dim monArr As Variant
monArr = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
For m = 0 To 11
monthCombo.AddItem (monArr(m))
Next m

For y = 2012 To 2015
yearCombo.AddItem (y)
Next y

For c = 1 To 4
adulCombo.AddItem (c)
chilCombo.AddItem (c - 1)
Next c
End Sub

Private Sub tourList_Click()
Dim k As Integer
For k = 0 To Adodc1.Recordset.RecordCount - 1
If Adodc1.Recordset.Fields("Tour") = tourList.Text Then
descLabel.Caption = Adodc1.Recordset.Fields("Description")
nightLabel.Caption = Adodc1.Recordset.Fields("Nights") & " nights"
Image1.Picture = LoadPicture(Adodc1.Recordset.Fields("PictureLocation"))
End If
Adodc1.Recordset.Move (1)
Next k
Adodc1.Recordset.MoveFirst
End Sub
