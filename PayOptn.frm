VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Payment details"
   ClientHeight    =   5085
   ClientLeft      =   9450
   ClientTop       =   5640
   ClientWidth     =   7890
   LinkTopic       =   "Form2"
   ScaleHeight     =   5085
   ScaleWidth      =   7890
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      ToolTipText     =   "Set Address"
      Top             =   3720
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   4680
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
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
      Connect         =   $"PayOptn.frx":0000
      OLEDBString     =   $"PayOptn.frx":019E
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
   Begin MSMask.MaskEdBox cvvNum 
      Height          =   285
      Left            =   2520
      TabIndex        =   16
      Top             =   3120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox cardNum 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Top             =   2640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   19
      Mask            =   "#### #### #### ####"
      PromptChar      =   "_"
   End
   Begin VB.OptionButton optDr 
      Caption         =   "Draft"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.OptionButton optCq 
      Caption         =   "Cheque"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   5640
      TabIndex        =   25
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox phone 
      Height          =   285
      Left            =   2520
      TabIndex        =   23
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox address 
      Height          =   285
      Left            =   2520
      TabIndex        =   21
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox payerName 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pay and View Ticket >>"
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   4200
      Width           =   1935
   End
   Begin VB.OptionButton optDC 
      Caption         =   "Debit Card"
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton optCC 
      Caption         =   "Credit Card"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1320
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.ComboBox bankName 
      Height          =   315
      Left            =   2520
      TabIndex        =   14
      Top             =   2160
      Width           =   3615
   End
   Begin VB.ComboBox cardYearVal 
      Height          =   315
      Left            =   6720
      TabIndex        =   18
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox cardMonthVal 
      Height          =   315
      Left            =   6000
      TabIndex        =   17
      Top             =   2640
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   120
      Picture         =   "PayOptn.frx":033C
      ScaleHeight     =   4635
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSMask.MaskEdBox dcNum 
      Height          =   285
      Left            =   6360
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   "_"
   End
   Begin VB.Label Label9 
      Caption         =   "Draft/Cheque"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   3120
      Width           =   975
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000000&
      X1              =   1200
      X2              =   1200
      Y1              =   2040
      Y2              =   3480
   End
   Begin VB.Line Line3 
      X1              =   7560
      X2              =   7560
      Y1              =   2040
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   7560
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label8 
      Caption         =   "Phone"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Address"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   3720
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   1200
      X2              =   7560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label6 
      Caption         =   "CVV"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Valid upto"
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Card number"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Bank"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Payment mode"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Payer"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
writeToDB
Form3.Show
End Sub

Sub writeToDB()
'Generate BookID first
Dim bid As String
Dim dStr As String
bid = Mid$(payerName.Text, 1, 3)
bid = bid & ((Form1.dateCombo.ListIndex + 1)) & ((Form1.monthCombo.ListIndex + 1)) & ((Form1.yearCombo.ListIndex + 12))
If optCC.Value = True Or optDC.Value = True Then
bid = bid & Right$(cardNum.Text, 4)
Else
bid = bid & Right$(dcNum.Text, 4)
End If
'BookID bid generated
'Generate departure date
dStr = ((Form1.monthCombo.ListIndex + 1)) & "/" & ((Form1.dateCombo.ListIndex + 1)) & "/" & ((Form1.yearCombo.ListIndex + 12))
'Departure date generated
Adodc1.CommandType = adCmdText
'Store tour data first
Adodc1.RecordSource = "SELECT * FROM BookedTours"
Adodc1.Refresh
Adodc1.Recordset.MoveLast
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("BookID") = bid
Adodc1.Recordset.Fields("TourID") = (Form1.tourList.ListIndex + 1)
Adodc1.Recordset.Fields("Departure") = dStr
Adodc1.Recordset.Fields("Adults") = Form1.adulCombo.Text
Adodc1.Recordset.Fields("Children") = Form1.chilCombo.Text
If Form1.ssCheck.Value = True Then
Adodc1.Recordset.Fields("Sightseeing") = True
Else
Adodc1.Recordset.Fields("Sightseeing") = False
End If
If Form1.fdCheck.Value = True Then
Adodc1.Recordset.Fields("Food") = True
Else
Adodc1.Recordset.Fields("Food") = False
End If
Adodc1.Recordset.Update
'Store payment data
Adodc1.RecordSource = "SELECT * FROM Payments"
Adodc1.Refresh
Adodc1.Recordset.MoveLast
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("BookID") = bid
Adodc1.Recordset.Fields("Payer") = payerName.Text
Adodc1.Recordset.Fields("Bank") = bankName.Text
Adodc1.Recordset.Fields("CardNumber") = cardNum.Text
Adodc1.Recordset.Fields("CVV") = cvvNum.Text
'Adodc1.Recordset.Fields("Validity") = cardMonthVal.Text & "/" & cardYearVal.Text
Adodc1.Recordset.Fields("ChequeDraftNum") = dcNum.Text
Adodc1.Recordset.Fields("Address") = address.Text
Adodc1.Recordset.Fields("Phone") = phone.Text
'Handle payment type
If optCC.Value = True Then
Adodc1.Recordset.Fields("PaymentMode") = "Credit Card"
Adodc1.Recordset.Fields("Validity") = cardMonthVal.Text & "/" & cardYearVal.Text
GoTo PayUpdate
End If

If optDC.Value = True Then
Adodc1.Recordset.Fields("PaymentMode") = "Debit Card"
Adodc1.Recordset.Fields("Validity") = cardMonthVal.Text & "/" & cardYearVal.Text
GoTo PayUpdate
End If

If optCq.Value = True Then
Adodc1.Recordset.Fields("PaymentMode") = "Cheque"
Else
Adodc1.Recordset.Fields("PaymentMode") = "Draft"
End If

PayUpdate:
 Adodc1.Recordset.Update
'add list of travellers
Adodc1.RecordSource = "SELECT BookID, Individual FROM Travelers"
Adodc1.Refresh
Adodc1.Recordset.MoveLast
For t = 0 To Form5.travList.ListCount - 1
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("BookID") = bid
Adodc1.Recordset.Fields("Individual") = Form5.travList.List(t)
Adodc1.Recordset.Update
Next t
End Sub

Private Sub Command3_Click()
Form6.Show
End Sub

Private Sub Form_Load()
For i = 1 To 12
cardMonthVal.AddItem (i)
Next i
For j = 2012 To 2015
cardYearVal.AddItem (j)
Next j
End Sub

Private Sub Option1_Click()
enableCardDetails
End Sub

Private Sub Option2_Click()
enableCardDetails
End Sub

Private Sub Option3_Click()
disableCardDetails
End Sub

Sub disableCardDetails()

cardNum.Enabled = False
cvvNum.Enabled = False
cardMonthVal.Enabled = False
cardYearVal.Enabled = False
dcNum.Enabled = True

Label4.Enabled = False
Label5.Enabled = False
Label6.Enabled = False
Label9.Enabled = True
End Sub

Sub enableCardDetails()

cardNum.Enabled = True
cvvNum.Enabled = True
cardMonthVal.Enabled = True
cardYearVal.Enabled = True
dcNum.Enabled = False

Label4.Enabled = True
Label5.Enabled = True
Label6.Enabled = True
Label9.Enabled = False
End Sub

Private Sub Option4_Click()
disableCardDetails
End Sub

Private Sub optCC_Click()
enableCardDetails
End Sub

Private Sub optCq_Click()
disableCardDetails
End Sub

Private Sub optDC_Click()
enableCardDetails
End Sub

Private Sub optDr_Click()
disableCardDetails
End Sub
