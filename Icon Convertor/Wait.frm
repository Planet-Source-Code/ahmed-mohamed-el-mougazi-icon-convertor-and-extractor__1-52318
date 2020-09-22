VERSION 5.00
Begin VB.Form Wait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   525
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "One Minute Please.............................................."
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Initialize()
InitCommonControls
End Sub
Private Sub Form_Load()
Top = (Screen.Height / 2) - (0.5 * Height)
Left = (Screen.Width / 2) - (0.5 * Width)
End Sub
