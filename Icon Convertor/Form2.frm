VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image to Icon Convertor"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Colors"
      Height          =   1335
      Left            =   2400
      TabIndex        =   20
      Top             =   1440
      Width           =   3855
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   3375
         TabIndex        =   21
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton Option9 
            Caption         =   "White and Black (2 Colors Only)"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2655
         End
         Begin VB.OptionButton Option8 
            Caption         =   "24 Bit"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   720
            Width           =   2655
         End
         Begin VB.OptionButton Option4 
            Caption         =   "16 Colors"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   240
            Width           =   2655
         End
         Begin VB.OptionButton Option5 
            Caption         =   "256 Colors"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   480
            Value           =   -1  'True
            Width           =   2655
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Size"
      Height          =   1095
      Left            =   2400
      TabIndex        =   19
      Top             =   360
      Width           =   3855
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   3375
         TabIndex        =   22
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton Option1 
            Caption         =   "16 X 16"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Width           =   2655
         End
         Begin VB.OptionButton Option2 
            Caption         =   "32 X 32"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton Option3 
            Caption         =   "48 X 48"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Top             =   480
            Width           =   2655
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2280
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2760
      Top             =   3720
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   1095
      Left            =   4320
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1335
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   360
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&rop options"
      Enabled         =   0   'False
      Height          =   435
      Left            =   240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1335
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Cr&op it"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
   Begin VB.OptionButton Option6 
      Caption         =   "&Scale it"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Back"
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "What you Want you to do to resize the picture"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Select Some properties:"
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   120
      Picture         =   "Form2.frx":08CA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   424
      Y1              =   304
      Y2              =   304
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
On Error Resume Next
Set pic = Picture2.Picture
If Option1.value = True Then
Dime1 = Option1.Caption
ElseIf Option2.value = True Then
Dime1 = Option2.Caption
Else
Dime1 = Option3.Caption
End If
Me.Enabled = False
Crop.Show
End Sub
Private Sub Command3_Click()
On Error Resume Next
Me.Hide
Form1.Show
Form1.Left = Left
Form1.Top = Top
End Sub
Private Sub Command4_Click()
On Error Resume Next
If Option1.value = True Then
  Dimension = 16
ElseIf Option2.value = True Then
Dimension = 32
Else
Dimension = 48
End If
If Option4.value = True Then
Dim MessageR
MessageR = MsgBox("Ther is 2 Methods for 16 Colors Icon Please chose Yes for first method or no for second method", vbInformation + vbYesNo, "Choose")
If MessageR = vbYes Then
Colors = 16
ChooseT = 0
Else
ChooseT = 1
End If
ElseIf Option5.value = True Then
Colors = 256
ElseIf Option9.value = True Then
Colors = 2
Else
Colors = 24
End If
Set Pic1 = Picture1.Image
Form5.Show
Form5.Top = Top
Form5.Left = Left
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
Private Sub Form_Load()
Sleep 100
End Sub
Private Sub Form_Paint()
Picture1_Paint
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Cancel1
End Sub
Private Sub Label3_Click()
Option1.value = True
End Sub
Private Sub Label7_Click()
Option5.value = True
End Sub
Private Sub Option1_Click()
    lWidth1 = 16
    lHeight1 = 16
    Picture1_Paint
    End Sub
Private Sub Option2_Click()
    lWidth1 = 32 '200
    lHeight1 = 32
    Picture1_Paint
    End Sub
Private Sub Option3_Click()
    lWidth1 = 48 '200
    lHeight1 = 48
   Picture1_Paint
End Sub

Private Sub Option6_Click()
Picture1_Paint
Command1.Enabled = False
End Sub
Private Sub Option7_Click()
Picture1_Paint
Command1.Enabled = True
End Sub
Private Sub Picture1_Paint()
On Error Resume Next
If lWidth1 = 0 Then lWidth1 = 32
If lHeight1 = 0 Then lHeight1 = 32
Picture1.Width = lWidth1
Picture1.Height = lHeight1
Set Picture2.Picture = LoadPicture(dire(0))
Picture1.Cls
Picture2.ScaleMode = 1
Picture1.ScaleMode = 3
If Option6.value = True Then
Dim x&, y&
x = ScaleX(Picture2.Picture.Width, vbHimetric, vbPixels)
y = ScaleY(Picture2.Picture.Height, vbHimetric, vbPixels)
Picture1.PaintPicture Picture2.Picture, 0, 0, lWidth1, lHeight1, 0, 0, x, y 'Picture2.Picture.Height, Picture2.Picture.Width, vbSrcCopy
Else
    Picture1.PaintPicture Picture2.Picture, 0, 0, lWidth1, lHeight1, myX, myY, lWidth1, lHeight1, vbSrcCopy
End If
End Sub
Private Sub Timer1_Timer()
Picture1.Top = (Picture3.Height - Picture1.Height) / 2
Picture1.Left = (Picture3.Width - Picture1.Width) / 2
End Sub
Private Sub drawing()
On Error Resume Next
Dim x&, y&
   x = ScaleX(Picture2.Picture.Width, vbHimetric, vbPixels)
    y = ScaleX(Picture2.Picture.Height, vbHimetric, vbPixels)
Picture1.PaintPicture Picture1.Image, 0, 0, lWidth1, lHeight1, 0, 0, x + 15, y + 15, vbSrcCopy
End Sub

