VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   1770
   ClientLeft      =   3120
   ClientTop       =   3360
   ClientWidth     =   4395
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form2.frx":0D2A
      Left            =   1080
      List            =   "Form2.frx":0D2C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0D2E
      Left            =   1080
      List            =   "Form2.frx":0D30
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "&Colors:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "&Dimensions:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Select Case Combo2.ListIndex
Case 0
Colors = 2
Case 1
Colors = 16
Case 2
Colors = 256
Case 3
Colors = 24
End Select
Select Case Combo1.ListIndex
Case 0
Dime1 = 16
Case 1
Dime1 = 32
Case 2
Dime1 = 48
End Select
Hide
End Sub
Private Sub Command2_Click()
Hide
End Sub
Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Form_Load()
Combo1.Clear
Combo2.Clear
For i = 16 To 48 Step 16
Combo1.AddItem i & "X" & i & " Pixels"
Next
Combo2.AddItem "Black and White (2 Colors Only)"
Combo2.AddItem "16 Colors"
Combo2.AddItem "256 Colors"
Combo2.AddItem "24 Bits"
Combo1.ListIndex = 1: Combo2.ListIndex = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = CBool(1)
Hide
End Sub
