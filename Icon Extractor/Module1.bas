Attribute VB_Name = "Module1"
Option Explicit
Public IcoName() As String
Public Declare Sub InitCommonControls Lib "comctl32" ()
Public Dime1 As Byte, Colors As Integer, i&
Public Function Topath(Filename As String) As String
If Right(Filename, 1) = "\" Then Topath = Filename Else Topath = Filename & "\"
End Function

