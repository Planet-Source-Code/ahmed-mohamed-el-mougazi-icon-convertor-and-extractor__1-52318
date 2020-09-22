Attribute VB_Name = "BuildandEnd"
Option Explicit 'Declare
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public ChooseT As Byte
Public Many As Boolean
Public dire() As String
Public lHeight1&, lWidth1&
Public pic As Picture, Dime1$, Pic1 As Picture
Public myX&, myY&, i&, J&
Public Colors As Integer, Dimension%
Sub Main() 'Starting
On Error Resume Next
Form1.Show
Set Form3 = Nothing
Set Form5 = Nothing
Set Form7 = Nothing
Set frmAbout = Nothing
Set Crop = Nothing
Set MskColor = Nothing
Set Wait = Nothing
Unload Wait
Unload MskColor
Unload Form3
Unload Form5
Unload Form7
Unload frmAbout
Unload Crop
End Sub
Sub Cancel1() 'Ending
On Error Resume Next
Kill App.Path & "\Clip.bmp"
Set Form1 = Nothing
Set Form3 = Nothing
Set Form5 = Nothing
Set Form7 = Nothing
Set frmAbout = Nothing
Set Crop = Nothing
Set MskColor = Nothing
Set Wait = Nothing
End
End Sub
