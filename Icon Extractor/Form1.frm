VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Extractor"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox myP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   12240
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   14
      Top             =   3960
      Width           =   492
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Options"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   2055
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   12840
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   3960
      Width           =   492
   End
   Begin MSComctlLib.ProgressBar P 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4665
      Visible         =   0   'False
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5880
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList IL 
      Left            =   5760
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&About"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Extract Icons"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Selected &Icons Only"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Change Folder"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3855
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6800
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select Icon Library"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "LOADING...PLEASE WAIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1680
      Left            =   2400
      TabIndex        =   12
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "There is no Library opened"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "&Folder Directory:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Type tDeviceImage
   iSizeX As Long
   iSizeY As Long
   cDepth As Long
   cPal As cPalette
End Type
Dim h&
Private m_tDeviceImage As tDeviceImage
Dim IconCont&
Dim IconCount
Dim hModule
Dim oIcon As cFileIcon
Dim IConh
Const BIF_RETURNONLYFSDIRS = 250
Const MAX_PATH = 260
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Sub Command1_Click()
CD1.DialogTitle = "Open Icon Library"
CD1.Flags = 5
CD1.Filter = "Icon Libraries|**.exe;**.icl;**.ni;**.dll;**.ico|All Files|**.**"
CD1.ShowOpen
If CD1.Filename = "" Then Exit Sub

ReadIcoLibrary CD1.Filename
End Sub

Private Sub Command2_Click()
'Taken Fram Api Guide
    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        'Set the owner window
        .hWndOwner = Me.hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("Choose The Folder You Want to Put Icons in it.", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            Text1 = Left$(sPath, iNull - 1)
        End If
    End If
End Sub

Private Sub Command3_Click()
Dim l As New FileSystemObject
If l.FolderExists(Text1) = False Then
MsgBox "Folder Doesnt Exist,Please Choose Another Folder", vbCritical, "Error"
Exit Sub
End If
Dim bSelectedOnly As Boolean, J&
bSelectedOnly = Check1.Value
   On Error GoTo 100
    Pic.Visible = True
    P.Visible = True
    For J = 1 To IconCount
        If bSelectedOnly Then
            If LV.ListItems(J).Selected = False Then GoTo SkipIcon
        End If
        SaveIconToFile J, Topath(Text1) + IcoName(J) + ".ico"
        P.Value = J
        DoEvents
SkipIcon:
    Next J
    Pic.Visible = False
    P.Visible = False
    Exit Sub
100
    MsgBox Err.Description, vbCritical, "Error"
    Exit Sub
End Sub

Private Sub SaveIconToFile(ByVal Index As Long, ByVal SaveName As String)
On Error Resume Next
Dim p15 As Picture
If Dime1 <> 32 Then
     Pic.Width = Dime1 * 15: Pic.Height = Dime1 * 15
    myP.Width = Dime1 * 15: myP.Height = Dime1 * 15
    myP.AutoRedraw = True
 Set p15 = IL.ListImages(Index).Picture
   myP.PaintPicture p15, 0, 0, Pic.ScaleWidth, Pic.ScaleHeight
      Set Pic.Picture = myP.Image
   Else
      Set Pic.Picture = IL.ListImages(Index).Picture
   End If
     oIcon.AddImage Dime1, Dime1, Colors
    m_tDeviceImage.cPal.SetPaletteToIcon oIcon, 1
    oIcon.SetIconFromBitmap Pic.hDC, 1, 0, 0, True, vbWhite
    oIcon.SaveIcon SaveName
    oIcon.RemoveImage 1
End Sub
Private Sub Command4_Click()
frmAbout.Show
End Sub
Private Sub Command5_Click()
On Error Resume Next
Set Form1 = Nothing
Set frmAbout = Nothing
Set Form2 = Nothing
End
End Sub

Private Sub Command6_Click()
Form2.Show
End Sub
Private Sub Form_Initialize()
On Error Resume Next
InitCommonControls
End Sub
Private Sub Form_Load()
Colors = 24
Dime1 = 32
Set oIcon = New cFileIcon
  hModule = Me.hWnd
    Pic.AutoRedraw = True
    myP.AutoRedraw = True
    With m_tDeviceImage
      .iSizeX = 32
      .iSizeY = 32
      .cDepth = 256
      Set .cPal = New cPalette
        .cPal.CreateWebSafe
    End With
  End Sub
Private Sub Form_Unload(Cancel As Integer)
Command5_Click
End Sub
Private Sub ReadIcoLibrary(FileN As String)
On Error GoTo 100
Dim Workon$
    LV.Icons = Nothing
    IL.ListImages.Clear
    LV.ListItems.Clear
Workon = FileN & Chr$(0)
IconCount = ExtractIcon(hModule, Workon, -1)
If IconCount > 0 Then
Command6.Enabled = True: Command2.Enabled = True: Label1.Enabled = True: Text1.Enabled = True: Check1.Enabled = True
Label2.Caption = " Icons Found in This File is " & IconCount & " Icon(s)"
P.Max = IconCount
P.Visible = True
LV.Visible = False
Getnames FileN
For i = 1 To IconCount
 Set Pic.Picture = LoadPicture("")
 IConh = ExtractIcon(hModule, FileN, i - 1)
 h = DrawIcon(Pic.hDC, 0, 0, IConh)
  IL.ListImages.Add i, , Pic.Image
LV.Icons = IL
 LV.ListItems.Add , , IcoName(i), i
 P.Value = i: DoEvents
 Next
       LV.Visible = True
        P.Visible = False
Else
Command2.Enabled = False: Label1.Enabled = False: Text1.Enabled = False: Check1.Enabled = False
Command6.Enabled = False
Label2.Caption = "Thers is no Icons in This File"
End If
Exit Sub
100:
MsgBox "Error Number " & Err.Number & " Occured", vbCritical, "Error"
    LV.Visible = True
        P.Visible = False
Exit Sub
End Sub
Private Sub Getnames(FileN As String)
Dim S As String, FN As Long
Dim x1 As Long
Dim Cnt As Long
Dim Z As Long
ReDim IcoName(1 To IconCount)
If UCase(Right(FileN, 3)) = UCase("icl") Then
    Cnt = 0
    FN = FreeFile
    Open FileN For Binary As #FN
    S = Space(LOF(FN))
    Get FN, , S
    Close #FN
    x1 = InStr(1, S, "ICL", vbBinaryCompare)
    If x1 = 0 Then GoTo PutNo
    x1 = x1 + 3
    Do
        Z = Asc(Mid(S, x1, 1))
        If Z = 0 Then Exit Do
        Cnt = Cnt + 1
        IcoName(Cnt) = Mid(S, x1 + 1, Z)
        x1 = x1 + Z + 1
    Loop
Else
PutNo:
    For i = 1 To IconCount
        IcoName(i) = "Icon" + Format(i, "0000")
    Next i
End If
End Sub

Private Sub Text1_Change()
If Len(Text1) > 0 Then Command3.Enabled = True Else Command3.Enabled = False
End Sub
