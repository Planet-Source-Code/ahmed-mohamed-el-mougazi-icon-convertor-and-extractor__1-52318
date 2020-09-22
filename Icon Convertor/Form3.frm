VERSION 5.00
Begin VB.Form MskColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Masking"
   ClientHeight    =   5250
   ClientLeft      =   1665
   ClientTop       =   2040
   ClientWidth     =   7725
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Masking"
      Height          =   2295
      Left            =   5520
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
      Begin VB.OptionButton Option1 
         Caption         =   "&Auto-mask"
         Height          =   375
         Index           =   2
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By &Color"
         Height          =   375
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By &Pixel"
         Height          =   375
         Index           =   0
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.PictureBox myoP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6240
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   480
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   120
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   4560
      Width           =   1935
   End
End
Attribute VB_Name = "MskColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MaskFl As Byte
Dim SelCol&
Dim FirstC&
Dim MaK&
Dim ratio&
Private Sub Command1_Click()
On Error Resume Next
Me.Hide
Form5.Enabled = True
Form5.SetFocus
Set Pic1 = myoP.Image
Unload Me
End Sub
Private Sub Command2_Click()
Form_Load
End Sub
Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim F&
Dim IW2&, IH2&
MaK = RGB(150, 150, 150)
myoP.Cls
myoP.Picture = LoadPicture()
myoP.Width = Form3.Picture1.Width: myoP.Height = Form3.Picture1.Height
myoP.Picture = Form3.Picture1.Image
Picture3.Cls
If Dimension <> 48 Then
Picture3.Width = 321: Picture3.Height = 321
Picture3.PaintPicture myoP.Picture, 0, 0, 321, 321
For F = 0 To Picture3.ScaleWidth Step Picture3.ScaleHeight / Dimension
Picture3.Line (F, 0)-(F, Picture3.ScaleHeight), &H4040&
Next F
For F = 0 To Picture3.ScaleHeight Step Picture3.ScaleHeight / Dimension
Picture3.Line (0, F)-(Picture3.ScaleWidth, F), &H4040&
Next F
Else
Call E48
End If
ratio = 321 / myoP.Width
End Sub
Private Sub Option1_Click(Index As Integer)
On Error Resume Next
If Index <> 2 Then
MaskFl = Index
Else
FirstC = Picture3.Point(1, 1)
If FirstC <> &H4040& Then
Enabled = False
DoEvents
Wait.Show: Wait.SetFocus
 For i = 0 To Picture3.ScaleWidth Step Picture3.ScaleHeight / Dimension
   For J = 0 To Picture3.ScaleWidth Step Picture3.ScaleHeight / Dimension
     If Picture3.Point(i + 1, J + 1) = FirstC Or Picture3.Point(i + 1, J + 1) = MaK Then
       Picture3.Line (i, J)-(i + (Picture3.ScaleHeight / Dimension), J + (Picture3.ScaleHeight / Dimension)), MaK, BF
       Picture3.Line (i, J)-(i + (Picture3.ScaleHeight / Dimension), J + (Picture3.ScaleHeight / Dimension)), vbBlack, B
    myoP.Line (i / ratio, J / ratio)-((i / ratio) + 0.25, (J / ratio) + 0.25), MaK, BF
        End If
    Next
   Next
Enabled = True
Unload Wait
Option1(0).Value = True: Option1(0).SetFocus
End If
End If
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim Helper1&, Helper2&
If Button = 1 Then
Select Case MaskFl
'By Pixel==============================
Case 0
For i = 0 To Picture3.ScaleWidth Step Picture3.ScaleHeight / Dimension
If x >= i And x < i + (Picture3.ScaleHeight / Dimension) Then
  Helper1 = i
End If
Next
For i = 0 To Picture3.ScaleWidth Step Picture3.ScaleHeight / Dimension
If y >= i And y < i + (Picture3.ScaleHeight / Dimension) Then
  Helper2 = i
End If
Next
Picture3.Line (Helper1, Helper2)-(Helper1 + (Picture3.ScaleHeight / Dimension), Helper2 + (Picture3.ScaleHeight / Dimension)), MaK, BF
Picture3.Line (Helper1, Helper2)-(Helper1 + (Picture3.ScaleHeight / Dimension), Helper2 + (Picture3.ScaleHeight / Dimension)), vbBlack, B
      myoP.Line (Helper1 / ratio, Helper2 / ratio)-((Helper1 / ratio) + 0.25, (Helper2 / ratio) + 0.25), MaK, BF
'By Color===============================
Case 1
Dim SelCol&
SelCol = Picture3.Point(x, y)
If SelCol <> &H4040& Then
Enabled = False
 For i = 0 To Picture3.ScaleWidth Step Picture3.ScaleHeight / Dimension
   For J = 0 To Picture3.ScaleWidth Step Picture3.ScaleHeight / Dimension
     If Picture3.Point(i + 1, J + 1) = SelCol Or Picture3.Point(i + 1, J + 1) = MaK Then
       Picture3.Line (i, J)-(i + (Picture3.ScaleHeight / Dimension), J + (Picture3.ScaleHeight / Dimension)), MaK, BF
       Picture3.Line (i, J)-(i + (Picture3.ScaleHeight / Dimension), J + (Picture3.ScaleHeight / Dimension)), vbBlack, B
       myoP.Line (i / ratio, J / ratio)-((i / ratio) + 0.25, (J / ratio) + 0.25), MaK, BF
       End If
    Next
   Next
Enabled = True
End If
End Select
End If
End Sub
Private Sub E48()
On Error Resume Next
Picture3.Width = 336: Picture3.Height = 336
Picture3.PaintPicture myoP.Picture, 0, 0, 336, 336
For J = 0 To Picture3.ScaleWidth Step 7
Picture3.Line (J, 0)-(J, Picture3.ScaleHeight), &H4040&
Next J
For J = 0 To Picture3.ScaleHeight Step 7
Picture3.Line (0, J)-(Picture3.ScaleWidth, J), &H4040&
Next J
End Sub
