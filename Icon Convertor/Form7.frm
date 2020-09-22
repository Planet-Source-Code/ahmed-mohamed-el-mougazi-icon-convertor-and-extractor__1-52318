VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image to Icon Convertor"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2880
      Top             =   3360
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Finish"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "All Picture(s) Successfully Converted"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   120
      Picture         =   "Form7.frx":08CA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   4560
      Y2              =   4560
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command3_Click()
Me.Hide
Form1.Show
End Sub
Private Sub Command4_Click()
Cancel1
End Sub
Private Sub Command6_Click()
frmAbout.Show
End Sub
Private Sub Form_Initialize()
On Error Resume Next
InitCommonControls
End Sub
Private Sub Form_Load()
Sleep 100
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel1
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
If Label1.ForeColor <> vbBlack Then Label1.ForeColor = vbBlack Else Label1.ForeColor = vbRed
End Sub
