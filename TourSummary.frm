VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Tour Summary"
   ClientHeight    =   5085
   ClientLeft      =   10500
   ClientTop       =   6675
   ClientWidth     =   7890
   LinkTopic       =   "Form3"
   ScaleHeight     =   5085
   ScaleWidth      =   7890
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1560
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   $"TourSummary.frx":0000
      OLEDBString     =   $"TourSummary.frx":019E
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
   Begin VB.CommandButton Command2 
      Caption         =   "View Ticket >>"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   4200
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   120
      Picture         =   "TourSummary.frx":033C
      ScaleHeight     =   4635
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label travelSum 
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label detailsSum 
      Height          =   495
      Left            =   5760
      TabIndex        =   14
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label payMode 
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Payment Status"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label7 
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Gross"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Details"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label departSum 
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label tourSum 
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label nameSum 
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Departure"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Travelers"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Tour"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tn As String
Dim travList As String
Dim bid As String
Dim pic As String

Private Sub Command2_Click()
printTicketToHtml
Form4.Show
End Sub

Sub printTicketToHtml()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
tn = "ticket.html"
FileNumber = FreeFile
Dim html As String
html = "<html><head><title>Your Ticket</title></head><body><table><td><img src='logo.gif' width='64' height='64'></td><td><h1>Tourista</h3></td></table><h2>Ticket #" & bid & "</h2><table>"
html = html & "<tr><td><b>Name</b></td><td>" & nameSum.Caption & "</td><tr> <td><b>Departure<b></td><td><i>" & departSum.Caption & "</i></td></tr>"
html = html & "<tr><td><b>Destination</b></td><td><i>" & "" & Form1.tourList.Text & "</i></td></tr>"
html = html & "<tr><td><b>Travelers</b></td><td>" & travList & "</td></tr>"
html = html & "<tr><td><b>Details</b></td><td>" & detailsSum.Caption & "</td></tr>"
html = html & "</table></body></html>"
Open tn For Output As #FileNumber
Print #FileNumber, html
Close #FileNumber
End Sub


Sub fetchFromDB()
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT Payments.BookID, Payments.Payer, Payments.PaymentMode, BookedTours.Departure, BookedTours.Sightseeing, BookedTours.Food FROM Payments INNER JOIN BookedTours ON Payments.BookID=BookedTours.BookID ORDER BY Payments.PID"
Adodc1.Refresh
Adodc1.Recordset.MoveLast
nameSum.Caption = Adodc1.Recordset.Fields("Payer")
departSum.Caption = Adodc1.Recordset.Fields("Departure")
bid = Adodc1.Recordset.Fields("BookID")
'pic = Adodc1.Recordset.Fields("PictureLocation")
If Adodc1.Recordset.Fields("PaymentMode") = "Credit Card" Or Adodc1.Recordset.Fields("PaymentMode") = "Debit Card" Then
payMode.ForeColor = RGB(0, 255, 0)
payMode.Caption = "Paid"
Else
payMode.ForeColor = RGB(255, 0, 0)
payMode.Caption = "Pending"
End If
If Adodc1.Recordset.Fields("Sightseeing") = True Then
detailsSum.Caption = detailsSum.Caption & "Sightseeing included "
Else
detailsSum.Caption = detailsSum.Caption & "Sightseeing excluded "
End If
If Adodc1.Recordset.Fields("Food") = True Then
detailsSum.Caption = detailsSum.Caption & "Food included"
Else
detailsSum.Caption = detailsSum.Caption & "Food excluded"
End If
'Generate list of travellers
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT Individual FROM Travelers WHERE BookID = '" & bid & "' ORDER BY Individual"
Adodc1.Refresh
'Adodc1.Recordset.MoveFirst
For t = 0 To Adodc1.Recordset.RecordCount - 1
travList = travList & Adodc1.Recordset.Fields("Individual") & " , "
Adodc1.Recordset.Move (1)
Next t
travelSum.Caption = Adodc1.Recordset.RecordCount & " ~ " & travList
End Sub

Private Sub Form_Load()
fetchFromDB
End Sub
