VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image to Icon Convertor"
   ClientHeight    =   5220
   ClientLeft      =   2280
   ClientTop       =   1860
   ClientWidth     =   6555
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   6555
   Begin VB.PictureBox myoP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1080
      Width           =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Masking"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2055
   End
   Begin MSComctlLib.ImageList IL 
      Left            =   5280
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   5880
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "B&rowse"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Width           =   3015
   End
   Begin VB.OptionButton Option2 
      Caption         =   "&BMP"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&ICO"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   840
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Back"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
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
   Begin VB.Label Label4 
      Caption         =   "On Saving With Bitmap The picture Will be 24Bit only."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4080
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Select Output Directory:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "This Picture Will Replace any other one in the same directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Please select the Type of Pictures you Want to Save"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   120
      Picture         =   "Form6.frx":08CA
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
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nam$
Private Type tDeviceImage
   iSizeX As Long
   iSizeY As Long
   cDepth As Long
   cPal As cPalette
End Type
Dim m_tDeviceImage As tDeviceImage
Private Sub Command1_Click()
On Error Resume Next
  Cd1.Flags = 5
  If Option1.value = True Then
     Cd1.Filter = "ICO Files|**.ico"
Else
  Cd1.Filter = "Bmp Files|**.bmp"
  End If
 Cd1.ShowSave
If Cd1.Filename = "" Then Exit Sub
Text1 = Cd1.Filename
End Sub
Private Sub Command2_Click()
MskColor.Show
Me.Enabled = False
End Sub
Private Sub Command3_Click()
On Error Resume Next
Me.Hide
Form3.Show
Form3.Top = Top
Form3.Left = Left
End Sub
Private Sub Command4_Click()
On Error Resume Next
If Option1.value = True Then
   If ChooseT <> 0 Then
              IL.ListImages.Clear
              IL.UseMaskColor = True
              IL.MaskColor = RGB(150, 150, 150)
               IL.ImageHeight = Dimension
               IL.ImageWidth = Dimension
               IL.ListImages.Add 1, , Pic1
               Set Pic1 = IL.ListImages.Item(1).ExtractIcon
               If Option1.value = True Then
               SavePicture Pic1, Text1.Text
               Else
              SavePicture Pic1, Text1
              End If
'256
  Else
'==========================================
Dim oIcon As New cFileIcon
         Set myoP.Picture = Pic1
        oIcon.AddImage Dimension, Dimension, Colors
        m_tDeviceImage.cPal.SetPaletteToIcon oIcon, 1
         oIcon.SetIconFromBitmap myoP.hdc, 1, 0, 0, True, RGB(150, 150, 150)
        oIcon.SaveIcon Text1.Text
'=================================================
  End If
Else
        SavePicture myoP.Picture, Text1
End If
If Many = False Then
100
Me.Hide
Form7.Show
Form7.Top = Top
Form7.Left = Left
Else
For i = 1 To Form1.List1.ListCount - 1
  dire(i - 1) = dire(i)
  dire(Form1.List1.ListCount - 1) = ""
Next
If dire(0) = "" Then GoTo 100
Me.Hide
Form3.Show
Form3.Top = Top
Form3.Left = Left
End If
End Sub
Private Sub Command5_Click()
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
On Error Resume Next
Dim nam$
Option1.value = True
Dim oIcon As cFileIcon
With m_tDeviceImage
 .iSizeX = Dimension
 .iSizeY = Dimension
 .cDepth = 256
 Set .cPal = New cPalette
 .cPal.CreateWebSafe
Text1 = Left(dire(0), Len(dire(0)) - 3) & "Ico"
              myoP.Width = Dimension * 15: myoP.Height = Dimension * 15
              myoP.Picture = Pic1
              End With
              Sleep 100
End Sub
Private Sub Form_Paint()
Option1.value = True
              myoP.Width = Dimension * 15: myoP.Height = Dimension * 15
              myoP.Picture = Pic1
              myoP.Left = Me.Width + myoP.Width + 1500
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel1
End Sub

Private Sub Option1_Click()
On Error Resume Next
Text1 = Left(Text1, Len(Text1) - 3) & "Ico"
End Sub
Private Sub Option2_Click()
On Error Resume Next
Text1 = Left(Text1, Len(Text1) - 3) & "Bmp"
End Sub
