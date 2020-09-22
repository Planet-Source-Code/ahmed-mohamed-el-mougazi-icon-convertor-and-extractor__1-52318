VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image to Icon Convertor"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   1920
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   4680
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "From &File"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3840
         Top             =   1560
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   7
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.ListBox List1 
         Height          =   1230
         ItemData        =   "Form1.frx":08CA
         Left            =   120
         List            =   "Form1.frx":08CC
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         ToolTipText     =   "Drag file(s) to me or click Add"
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "From &Clipboard"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label2 
      Caption         =   "Please Selct your Picture(s) source here:"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":08CE
      Height          =   855
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2130
      Left            =   120
      Picture         =   "Form1.frx":096C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ext$, Point1$, Flag&
Private Sub Command1_Click()
On Error Resume Next
Cd1.Flags = 5
Cd1.DialogTitle = "Get Pictures"
Cd1.Filter = "Bitmap Files|**.bmp;**.dib|JPEG Files|**.jpg|Gif Files|**.gif|Others|**.wmf;**.emf|All Picture Files|**bmp;**.jpg;**.gif;**.emf;**.wmf"
Cd1.FilterIndex = 5
Cd1.ShowOpen
If Cd1.Filename = "" Then Exit Sub
 If Cd1.Filename = List1.List(i) Then
    MsgBox "There is other picture with this name", vbCritical, "Error"
    Exit Sub
 End If
List1.AddItem Cd1.Filename
End Sub
Private Sub Command2_Click()
On Error Resume Next
Dim ind As Integer
If List1.ListIndex > -1 Then
ind = List1.ListIndex
List1.RemoveItem ind
Else
List1.RemoveItem 0
End If
List1.Refresh
End Sub
Private Sub Command4_Click()
On Error Resume Next
If Option2.Value = True Then
ReDim dire(0 To List1.ListCount - 1)
For i = 0 To List1.ListCount - 1
dire(i) = List1.List(i)
Next
Form3.Show
Form3.Top = Top
Form3.Left = Left
Form3.Option1.Value = True
Form3.Option2.Value = True
If Option1.Value = True Then
Many = False
ElseIf Option2.Value = True And List1.ListCount = 1 Then
Many = False
Else
Many = True
End If
ElseIf Option1.Value = True Then
SavePicture Clipboard.GetData, App.Path & "\Clip.bmp"
ReDim dire(0)
dire(0) = App.Path & "\Clip.bmp"
Form3.Show
Form3.Top = Top
Form3.Left = Left
Many = False
End If
Me.Hide
End Sub
Private Sub Command5_Click()
On Error Resume Next
Cancel1
End Sub
Private Sub Command6_Click()
On Error Resume Next
frmAbout.Show
End Sub
Private Sub Form_Initialize()
On Error Resume Next
InitCommonControls
End Sub
Private Sub Form_Paint()
On Error Resume Next
If Clipboard.GetData(vbCFBitmap) = 0 Then Option1.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Cancel1
End Sub
Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim hh, Na$
If Button = 0 Then
For Each hh In Data.Files
For i = 0 To List1.ListCount - 1
 If hh = List1.List(i) Then
    MsgBox "There is other picture with this name", vbCritical, "Error"
    GoTo 100
 End If
Next
 For i = Len(hh) To 0 Step -1
 Point1 = Mid(hh, i, 1)
      If Point1 = "." Then
      Ext = Right(hh, Len(hh) - i)
      Exit For
       End If
Next
Select Case UCase(Ext)
Case UCase("bmp"), UCase("gif"), UCase("Jpg"), UCase("dib"), UCase("Wmf"), UCase("emf")
Na = hh
End Select
If Na <> "" Then List1.AddItem Na
100:
Next
End If
End Sub
Private Sub Option1_Click()
On Error Resume Next
Frame1.Visible = False
End Sub
Private Sub Option2_Click()
On Error Resume Next
Frame1.Visible = True
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
If List1.ListCount = 0 Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If
If Option1.Value = True Or (Option2.Value = True And List1.ListCount > 0) Then
Command4.Enabled = True
Exit Sub
ElseIf Option2.Value = True And List1.ListCount = 0 Then
Command4.Enabled = False
End If
End Sub
