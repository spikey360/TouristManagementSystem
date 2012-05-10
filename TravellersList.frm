VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Travellers"
   ClientHeight    =   5085
   ClientLeft      =   8835
   ClientTop       =   4590
   ClientWidth     =   6240
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Finished"
      Height          =   375
      Left            =   2213
      TabIndex        =   10
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add >>"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.ListBox travList 
      Height          =   2010
      ItemData        =   "TravellersList.frx":0000
      Left            =   2880
      List            =   "TravellersList.frx":0002
      TabIndex        =   11
      Top             =   2280
      Width           =   3015
   End
   Begin MSMask.MaskEdBox ageText 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   2
      Mask            =   "##"
      PromptChar      =   "_"
   End
   Begin VB.OptionButton isFemale 
      Caption         =   "Female"
      Height          =   375
      Left            =   3120
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.OptionButton isMale 
      Caption         =   "Male"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   1560
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox nameText 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   120
      Picture         =   "TravellersList.frx":0004
      ScaleHeight     =   4635
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Sex"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Age"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim txt As String
txt = txt & nameText.Text & " : " & ageText.Text
If isMale.Value = True Then
txt = txt & " : Male"
End If
If isFemale.Value = True Then
txt = txt & " : Female"
End If
travList.AddItem (txt)
nameText.Text = ""

End Sub

Private Sub Command2_Click()
If travList.ListCount > 0 Then
travList.RemoveItem (travList.ListIndex)
Command2.Enabled = False
End If
End Sub

Private Sub Command3_Click()
Form5.Hide
End Sub

Private Sub travList_Click()
Command2.Enabled = True
End Sub
