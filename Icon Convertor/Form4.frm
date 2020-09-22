VERSION 5.00
Begin VB.Form Crop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crop Options"
   ClientHeight    =   3750
   ClientLeft      =   1665
   ClientTop       =   2040
   ClientWidth     =   7935
   ControlBox      =   0   'False
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   3390
      Left            =   120
      ScaleHeight     =   3330
      ScaleWidth      =   5430
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   5490
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   0
         ScaleHeight     =   1335
         ScaleWidth      =   1815
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   1815
         Begin VB.Shape Shape1 
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   3340
         LargeChange     =   100
         Left            =   5130
         SmallChange     =   10
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   300
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   300
         LargeChange     =   100
         Left            =   0
         SmallChange     =   10
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3030
         Width           =   5130
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Corp"
      Height          =   2295
      Left            =   5640
      TabIndex        =   2
      Top             =   600
      Width           =   2175
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Height          =   615
         Left            =   720
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "&Top"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "&Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   375
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      Caption         =   "Icon:"
      Height          =   255
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Crop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
On Error Resume Next
Me.Hide
myX = Shape1.Left \ 15
myY = Shape1.Top \ 15
Form3.Enabled = True
Form3.SetFocus
Set Crop = Nothing
Unload Me
End Sub
Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Form_Load()
On Error Resume Next
Label2.Caption = Dime1
Picture3.Picture = pic


 If Picture3.Width > (Picture2.ScaleWidth - (VScroll1.Width + 30)) Then
        HScroll1.Value = 0
        HScroll1.Enabled = True
        HScroll1.Max = Picture3.Width - (Picture2.Width - VScroll1.Width - 60)
    Else
        HScroll1.Enabled = False
    End If
    If Picture3.Height > (Picture2.ScaleHeight - (HScroll1.Height + 30)) Then
        VScroll1.Value = 0
        VScroll1.Enabled = True
        VScroll1.Max = Picture3.Height - (Picture2.Height - HScroll1.Height - 60)
    Else
        VScroll1.Enabled = False
    End If
Dim x As Byte
With Form3
If .Option1.Value = True Then
x = 16
ElseIf .Option2.Value = True Then
x = 32
Else
x = 48
End If
End With
Shape1.Width = x * 15
Shape1.Height = x * 15
Picture1.Width = x * 15
Picture1.Height = x * 15
Picture1.Left = Frame1.Width / 2 - (0.5 * Picture1.Width)
HScroll1.ZOrder
VScroll1.ZOrder
Show
Picture3_MouseMove 1, 0, 0, 0
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 1 Then
If x >= 0 Then Shape1.Left = x: Text1 = x \ 15
If y >= 0 Then Shape1.Top = y: Text2 = y \ 15
Picture1.Cls
Picture1.Refresh
Picture1.PaintPicture Picture3.Picture, 0, 0, Picture1.Width, Picture1.Height, Shape1.Left, Shape1.Top, Picture1.Width, Picture1.Height, vbSrcCopy
End If
End Sub
Private Sub HScroll1_Scroll()
On Error Resume Next
HScroll1_Change
End Sub
Private Sub VScroll1_Change()
On Error Resume Next
        Picture3.Top = -VScroll1.Value
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
    Picture3.Left = -HScroll1.Value
End Sub
Private Sub VScroll1_Scroll()
On Error Resume Next
VScroll1_Change
End Sub
