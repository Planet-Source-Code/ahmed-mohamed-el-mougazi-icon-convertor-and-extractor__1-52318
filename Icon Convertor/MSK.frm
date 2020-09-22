VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Msk 
   AutoRedraw      =   -1  'True
   Caption         =   "IconMaker"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6345
   Icon            =   "MSK.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   423
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   556
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "save"
            Object.ToolTipText     =   "Save As"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "fill"
            Object.ToolTipText     =   "Fill"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "paint"
            Object.ToolTipText     =   "Paint"
            Object.Tag             =   ""
            ImageIndex      =   4
            Value           =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "change"
            Object.ToolTipText     =   "Change Color"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "undo"
            Object.ToolTipText     =   "Undo"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "grate"
            Object.ToolTipText     =   "Grid Color"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "up"
            Object.ToolTipText     =   "Move"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "down"
            Object.ToolTipText     =   "Move"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "left"
            Object.ToolTipText     =   "Move"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "right"
            Object.ToolTipText     =   "Move"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "rotate"
            Object.ToolTipText     =   "Rotate"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4335
      Left            =   5880
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   19
      ToolTipText     =   "Current color"
      Top             =   1200
      Width           =   150
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5040
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   18
      ToolTipText     =   "Transparent color"
      Top             =   5640
      Width           =   720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      TabIndex        =   17
      Top             =   1200
      Width           =   4815
      Begin VB.PictureBox Picture11 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   4080
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   26
         Top             =   2880
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Picture10 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   2760
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   25
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   24
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture9 
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   1920
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   23
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   22
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox Picture7 
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   720
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   21
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin ComctlLib.ImageList ImageList2 
         Left            =   2400
         Top             =   1560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   12
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":0928
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":0986
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":09E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":0AA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":0AFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":0B5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":0BBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":0C18
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":0C76
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MSK.frx":0CD4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   1320
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         _Version        =   327682
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   15
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   16
      Top             =   5250
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   14
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   15
      Top             =   4980
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   13
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   14
      Top             =   4710
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   12
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   13
      Top             =   4440
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   11
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   12
      Top             =   4170
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   10
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   11
      Top             =   3900
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   9
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   10
      Top             =   3630
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   8
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   9
      Top             =   3360
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   7
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   8
      Top             =   3090
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   6
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   7
      Top             =   2820
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   6
      Top             =   2550
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   5
      Top             =   2280
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   4
      Top             =   2010
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   3
      Top             =   1740
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   2
      Top             =   1470
      Width           =   720
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   720
   End
   Begin VB.PictureBox Picture1 
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
      TabIndex        =   0
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   5520
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu op 
         Caption         =   "&Open"
      End
      Begin VB.Menu sva 
         Caption         =   "&Save As"
      End
      Begin VB.Menu hr 
         Caption         =   "-"
      End
      Begin VB.Menu ex 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu ed 
      Caption         =   "E&dit"
      Begin VB.Menu und 
         Caption         =   "&Undo"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu effx 
      Caption         =   "&Effects"
      Begin VB.Menu flHor 
         Caption         =   "Flip &Horizontal"
      End
      Begin VB.Menu flVer 
         Caption         =   "Flip &Vertikal"
      End
      Begin VB.Menu hr3 
         Caption         =   "-"
      End
      Begin VB.Menu Rotat 
         Caption         =   "&Rotate"
      End
      Begin VB.Menu hr2 
         Caption         =   "-"
      End
      Begin VB.Menu chCol 
         Caption         =   "Change Color"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "&Help"
      Begin VB.Menu abo 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Msk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a
Dim b
Dim c
Dim d
Dim w
Dim lin
Dim UNL
Private Sub PaintDown()
Static Pont
Pont = Picture3.Point(0, 0)
Picture3.PaintPicture Picture1.Image, 0, 0, 321, 321
If Pont = &HFFC0FF Then
For F = 0 To Picture3.ScaleHeight Step 10
Picture3.Line (0, F)-(Picture3.ScaleWidth, F), &HFFC0FF
Next F
For F = 0 To Picture3.ScaleWidth Step 10
Picture3.Line (F, 0)-(F, Picture3.ScaleHeight), &HFFC0FF
Next F
Line (Picture1.Left - 1, Picture1.Top - 1)-(Picture1.Left + Picture1.Width, Picture1.Top + Picture1.Height), QBColor(15), B
Else
For F = 0 To Picture3.ScaleHeight Step 10
Picture3.Line (0, F)-(Picture3.ScaleWidth, F), &H4040&
Next F
For F = 0 To Picture3.ScaleWidth Step 10
Picture3.Line (F, 0)-(F, Picture3.ScaleHeight), &H4040&
Next F
Line (Picture1.Left - 1, Picture1.Top - 1)-(Picture1.Left + Picture1.Width, Picture1.Top + Picture1.Height), QBColor(0), B
End If
End Sub
Private Sub Undo1()
Picture6 = Picture1.Image
Picture7 = Picture3.Image

End Sub
Private Sub Undo2()
Picture8 = Picture1.Image
Picture9 = Picture3.Image
Picture10 = Picture1.Image

End Sub
Private Sub Fill()
UNL = 1

w = 0
Toolbar1.Buttons(6).Enabled = True
und.Enabled = True
Undo1
Picture1.Line (0, 0)-(31, 31), a, BF
Picture3.BackColor = a
If a = QBColor(0) Or a = QBColor(1) Then
For F = 0 To Picture3.ScaleHeight Step 10
Picture3.Line (0, F)-(Picture3.ScaleWidth, F), &HFFC0FF
Next F
For F = 0 To Picture3.ScaleWidth Step 10
Picture3.Line (F, 0)-(F, Picture3.ScaleHeight), &HFFC0FF
Next F
Line (Picture1.Left - 1, Picture1.Top - 1)-(Picture1.Left + Picture1.Width, Picture1.Top + Picture1.Height), QBColor(15), B
Else
For F = 0 To Picture3.ScaleHeight Step 10
Picture3.Line (0, F)-(Picture3.ScaleWidth, F), &H4040&
Next F
For F = 0 To Picture3.ScaleWidth Step 10
Picture3.Line (F, 0)-(F, Picture3.ScaleHeight), &H4040&
Next F
Line (Picture1.Left - 1, Picture1.Top - 1)-(Picture1.Left + Picture1.Width, Picture1.Top + Picture1.Height), QBColor(0), B
End If
Undo2
End Sub

Private Sub leftt()
On Error GoTo ex
w = 0
Undo1
Picture1.PaintPicture Picture10, -1, 0
Picture1.PaintPicture Picture10, 31, 0
PaintDown
Undo2
ex:
End Sub

Private Sub rightt()
w = 0
Undo1
Picture1.PaintPicture Picture10, 1, 0
Picture1.PaintPicture Picture10, -31, 0
PaintDown
Undo2

End Sub

Private Sub upp()
w = 0
Undo1
Picture1.PaintPicture Picture10, 0, -1
Picture1.PaintPicture Picture10, 0, 31
PaintDown
Undo2

End Sub

Private Sub downn()
w = 0
Undo1
Picture1.PaintPicture Picture10, 0, 1
Picture1.PaintPicture Picture10, 0, -31
PaintDown
Undo2

End Sub

Private Sub abo_Click()
Form2.Show vbModal, Me
End Sub

Private Sub chCol_Click()
    Toolbar1.Buttons(4).Value = tbrUnpressed
    Toolbar1.Buttons(5).Value = tbrPressed
    Picture3.MousePointer = 10

End Sub




Private Sub ex_Click()
Unload Me
End Sub


Private Sub flHor_Click()
Undo1
Picture1.PaintPicture Picture10, Picture10.ScaleWidth - 1, 0, -Picture10.ScaleWidth, Picture10.ScaleHeight
PaintDown
Undo2

End Sub

Private Sub flVer_Click()
Undo1
Picture1.PaintPicture Picture10, 0, Picture10.ScaleHeight - 1, Picture10.ScaleWidth, -Picture10.ScaleHeight
PaintDown
Undo2

End Sub

Private Sub Form_Load()
Picture1.BackColor = &HFFC0C0
Picture3.BackColor = &HFFC0C0
Picture4.BackColor = &HFFC0C0
ImageList1.MaskColor = &HFFC0C0
For e = 0 To 15
Picture2(e).BackColor = QBColor(e)
Next e
For F = 0 To Picture3.ScaleHeight Step 10
Picture3.Line (0, F)-(Picture3.ScaleWidth, F), &H4040&
Next F
For F = 0 To Picture3.ScaleWidth Step 10
Picture3.Line (F, 0)-(F, Picture3.ScaleHeight), &H4040&
Next F
Line (Picture1.Left - 1, Picture1.Top - 1)-(Picture1.Left + Picture1.Width, Picture1.Top + Picture1.Height), QBColor(0), B
Picture4.Print "    Trpt"

Picture11 = Icon

Icon = Image1
For e = 1 To 15
Line (Picture2(e).Left - 1, Picture2(e).Top - 1)-(Picture2(e).Left + Picture2(e).Width, Picture2(e).Top + Picture2(e).Height), QBColor(0), B
Next e
Line (Picture4.Left - 1, Picture4.Top - 1)-(Picture4.Left + Picture4.Width, Picture4.Top + Picture4.Height), QBColor(0), B
Line (Picture2(0).Left - 1, Picture2(0).Top - 1)-(Picture2(0).Left + Picture2(0).Width, Picture2(0).Top + Picture2(0).Height), &HFFC0C0, B

Picture10 = Picture1.Image

End Sub


Private Sub Form_Unload(Cancel As Integer)
If UNL <> 0 Then
Dim Msg, Style, Resp
Msg = "Save icon?"
Style = vbYesNoCancel + vbExclamation
Resp = MsgBox(Msg, Style)
If Resp = vbYes Then sva_Click: If UNL <> 0 Then Cancel = True
If Resp = vbNo Then Cancel = False
If Resp = vbCancel Then Cancel = True
End If
End Sub

Private Sub op_Click()
CommonDialog1.CancelError = True
On Error GoTo ex
CommonDialog1.Filename = ""
CommonDialog1.Flags = cdlOFNFileMustExist
CommonDialog1.Filter = "Icons (*.ico)|*.ico"
CommonDialog1.ShowOpen
If FileLen(CommonDialog1.Filename) <> 766 Then
MsgBox "Ivalid or unsupported file format.", vbCritical
Exit Sub
End If
MousePointer = 11
w = 0
Toolbar1.Buttons(6).Enabled = True
und.Enabled = True
Undo1
Picture1.BackColor = &HFFC0C0
Picture1 = LoadPicture(CommonDialog1.Filename)
PaintDown
'+Çàëèâêà íèæíåé()
UNL = 1
Undo2
MousePointer = 0
Exit Sub
ex:

End Sub

Private Sub Picture2_Click(Index As Integer)
a = QBColor(Index)
For e = 0 To 15
Line (Picture2(e).Left - 1, Picture2(e).Top - 1)-(Picture2(e).Left + Picture2(e).Width, Picture2(e).Top + Picture2(e).Height), QBColor(0), B
Next e
Line (Picture2(Index).Left - 1, Picture2(Index).Top - 1)-(Picture2(Index).Left + Picture2(Index).Width, Picture2(Index).Top + Picture2(Index).Height), &HFFC0C0, B
Line (Picture4.Left - 1, Picture4.Top - 1)-(Picture4.Left + Picture4.Width, Picture4.Top + Picture4.Height), QBColor(0), B
Picture5.BackColor = a
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
b = Picture3.Point(X, Y)
If b = &H4040& Or b = &HFFC0FF Then Exit Sub
If Button = vbRightButton Then
a = Picture3.Point(X, Y)

If a = QBColor(0) Then
Picture2_Click (0)
Exit Sub
End If

If a = QBColor(1) Then
Picture2_Click (1)
Exit Sub
End If

If a = QBColor(2) Then
Picture2_Click (2)
Exit Sub
End If

If a = QBColor(3) Then
Picture2_Click (3)
Exit Sub
End If

If a = QBColor(4) Then
Picture2_Click (4)
Exit Sub
End If

If a = QBColor(5) Then
Picture2_Click (5)
Exit Sub
End If

If a = QBColor(6) Then
Picture2_Click (6)
Exit Sub
End If

If a = QBColor(7) Then
Picture2_Click (7)
Exit Sub
End If

If a = QBColor(8) Then
Picture2_Click (8)
Exit Sub
End If

If a = QBColor(9) Then
Picture2_Click (9)
Exit Sub
End If

If a = QBColor(10) Then
Picture2_Click (10)
Exit Sub
End If

If a = QBColor(11) Then
Picture2_Click (11)
Exit Sub
End If

If a = QBColor(12) Then
Picture2_Click (12)
Exit Sub
End If

If a = QBColor(13) Then
Picture2_Click (13)
Exit Sub
End If


If a = QBColor(14) Then
Picture2_Click (14)
Exit Sub
End If


If a = QBColor(15) Then
Picture2_Click (15)
Exit Sub
End If


If a = &HFFC0C0 Then
Picture4_Click
Exit Sub
End If

End If





If Button <> vbLeftButton Then Exit Sub
UNL = 1
w = 0
Undo1
Toolbar1.Buttons(6).Enabled = True
und.Enabled = True
If Picture3.MousePointer = 10 Then
b = Picture3.Point(X, Y)
If b = a Then Exit Sub
Picture3.MousePointer = 11
For J = 0 To Picture1.ScaleWidth - 1
For p = 0 To Picture1.ScaleHeight - 1
c = Picture1.Point(J, p)
If c = b Then Picture1.PSet (J, p), a
Next p
Next J
PaintDown
'+Çàëèâêà íèæíåé()
Picture3.MousePointer = 10
Exit Sub
End If
lin = 1
'0000000000000

X1 = 0
Y1 = 0
For J = 0 To 31
For p = 0 To 31

If X < X1 + 10 And X > X1 And Y < Y1 + 10 And Y > Y1 Then
Picture3.Line (X1 + 1, Y1 + 1)-(X1 + 9, Y1 + 9), a, BF
Picture1.PSet (X1 \ 10, Y1 \ 10), a
End If

X1 = X1 + 10
If X1 = 320 Then
X1 = 0
Y1 = Y1 + 10
End If


Next p
Next J
'31 31 31 31 31 31 !!!!!!!!!

End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lin = 0 Then Exit Sub
'0000000000000

X1 = 0
Y1 = 0
For J = 0 To 31
For p = 0 To 31

If X < X1 + 10 And X > X1 And Y < Y1 + 10 And Y > Y1 Then
Picture3.Line (X1 + 1, Y1 + 1)-(X1 + 9, Y1 + 9), a, BF
Picture1.PSet (X1 \ 10, Y1 \ 10), a
End If

X1 = X1 + 10
If X1 = 320 Then
X1 = 0
Y1 = Y1 + 10
End If


Next p
Next J
'31 31 31 31 31 31 !!!!!!!!!

End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lin = 0
If Button = vbLeftButton Then
Undo2
End If
End Sub

Private Sub Picture4_Click()
a = &HFFC0C0
For e = 0 To 15
Line (Picture2(e).Left - 1, Picture2(e).Top - 1)-(Picture2(e).Left + Picture2(e).Width, Picture2(e).Top + Picture2(e).Height), QBColor(0), B
Next e
Line (Picture4.Left - 1, Picture4.Top - 1)-(Picture4.Left + Picture4.Width, Picture4.Top + Picture4.Height), QBColor(15), B
Picture5.BackColor = a
End Sub

Private Sub Rotat_Click()
Undo1
For J = 0 To Picture1.ScaleWidth - 1
For p = 0 To Picture1.ScaleHeight - 1
Picture1.PSet (J, p), Picture10.Point(p, J)
Next p
Next J
Picture10 = Picture1.Image
Picture1.PaintPicture Picture10, Picture10.ScaleWidth - 1, 0, -Picture10.ScaleWidth, Picture10.ScaleHeight
PaintDown
Undo2

End Sub

Private Sub sva_Click()
CommonDialog1.CancelError = True
On Error GoTo ex
CommonDialog1.Filename = ""
CommonDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
CommonDialog1.Filter = "Icons (*.ico)|*.ico|Bitmaps (*.bmp)|*.bmp"
CommonDialog1.ShowSave
If CommonDialog1.FilterIndex = 1 Then
MousePointer = 11
Dim imgX As ListImage
Set imgX = ImageList1.ListImages. _
Add(1, , Picture1.Image)
Dim picX As Picture
Set picX = ImageList1.ListImages(1).ExtractIcon
SavePicture picX, CommonDialog1.Filename
UNL = 0
MousePointer = 0
End If
If CommonDialog1.FilterIndex = 2 Then
SavePicture Picture1.Image, CommonDialog1.Filename
UNL = 0
End If
Exit Sub
ex:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case Is = "rotate"
    Rotat_Click
    Case Is = "open"
    op_Click
    Case Is = "save"
    sva_Click
    Case Is = "fill"
    Fill
    Case Is = "paint"
    Toolbar1.Buttons(5).Value = tbrUnpressed
    Toolbar1.Buttons(4).Value = tbrPressed
    Picture3.MousePointer = 99
    Case Is = "change"
    Toolbar1.Buttons(4).Value = tbrUnpressed
    Toolbar1.Buttons(5).Value = tbrPressed
    Picture3.MousePointer = 10
    Case Is = "undo"
    und_Click
    Case Is = "grate"
If Picture3.Point(0, 0) = &H4040& Then
For F = 0 To Picture3.ScaleHeight Step 10
Picture3.Line (0, F)-(Picture3.ScaleWidth, F), &HFFC0FF
Next F
For F = 0 To Picture3.ScaleWidth Step 10
Picture3.Line (F, 0)-(F, Picture3.ScaleHeight), &HFFC0FF
Next F
Line (Picture1.Left - 1, Picture1.Top - 1)-(Picture1.Left + Picture1.Width, Picture1.Top + Picture1.Height), QBColor(15), B
Else
For F = 0 To Picture3.ScaleHeight Step 10
Picture3.Line (0, F)-(Picture3.ScaleWidth, F), &H4040&
Next F
For F = 0 To Picture3.ScaleWidth Step 10
Picture3.Line (F, 0)-(F, Picture3.ScaleHeight), &H4040&
Next F
Line (Picture1.Left - 1, Picture1.Top - 1)-(Picture1.Left + Picture1.Width, Picture1.Top + Picture1.Height), QBColor(0), B
End If
    Case Is = "up"
    upp
    Case Is = "down"
    downn
    Case Is = "left"
    leftt
    Case Is = "right"
    rightt
    'Case Else

End Select
End Sub

Private Sub und_Click()
    If w = 0 Then
    Picture1.PaintPicture Picture6, 0, 0
    Picture3.PaintPicture Picture7, 0, 0
    w = 1
    Else
    Picture1.PaintPicture Picture8, 0, 0
    Picture3.PaintPicture Picture9, 0, 0
    w = 0
    End If
    Picture10 = Picture1.Image

End Sub
