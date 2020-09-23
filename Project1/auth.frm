VERSION 5.00
Begin VB.Form auth 
   Caption         =   "Generate authentication Number"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5415
   ControlBox      =   0   'False
   Icon            =   "auth.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&END"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtAuth1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaxLength       =   5
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtAuth2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtAuth3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtAuth4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtRec1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtRec2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtRec3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtRec4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   $"auth.frx":030A
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "This is the number to give the customer for authentication."
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   5055
   End
End
Attribute VB_Name = "auth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Dim TempStr As String
    Dim Temp As String
    Dim AuthString As String
    Dim I As Integer
    
    Temp = txtAuth1.Text
    Temp = Temp + txtAuth2.Text
    Temp = Temp + txtAuth3.Text
    Temp = Temp + txtAuth4.Text
    For I = 1 To Len(Temp)
        TempStr = Mid$(Temp, I, 1)
        Select Case TempStr
            Case "0"
                AuthString = AuthString + "G"
            Case "1"
                AuthString = AuthString + "I"
            Case "2"
                AuthString = AuthString + "K"
            Case "3"
                AuthString = AuthString + "M"
            Case "4"
                AuthString = AuthString + "O"
            Case "5"
                AuthString = AuthString + "Q"
            Case "6"
                AuthString = AuthString + "S"
            Case "7"
                AuthString = AuthString + "U"
            Case "8"
                AuthString = AuthString + "V"
            Case "9"
                AuthString = AuthString + "T"
            Case "A"
                AuthString = AuthString + "R"
            Case "B"
                AuthString = AuthString + "P"
            Case "C"
                AuthString = AuthString + "N"
            Case "D"
                AuthString = AuthString + "L"
            Case "E"
                AuthString = AuthString + "J"
            Case "F"
                AuthString = AuthString + "H"
        End Select
    Next I
    txtRec1.Text = Left$(AuthString, 5)
    txtRec2.Text = Mid$(AuthString, 6, 5)
    txtRec3.Text = Mid$(AuthString, 11, 5)
    txtRec4.Text = Right$(AuthString, 5)

End Sub

Private Sub Command1_Click()
    Unload Me
    End
End Sub

Private Sub txtAuth1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 8 Then
        Exit Sub
    End If
    If KeyAscii > 47 And KeyAscii < 58 Then
    
    ElseIf KeyAscii < 65 Then
        KeyAscii = 0
    ElseIf KeyAscii > 64 And KeyAscii < 91 Then
        'ok
    ElseIf KeyAscii > 96 And KeyAscii < 103 Then
        KeyAscii = KeyAscii - 32
    Else
        KeyAscii = 0
    End If
    
    If Len(txtAuth1.Text) > 3 Then
        txtAuth2.SetFocus
    End If
    
End Sub

Private Sub txtAuth2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 8 Then
        Exit Sub
    End If
    If KeyAscii > 47 And KeyAscii < 58 Then
    
    ElseIf KeyAscii < 65 Then
        KeyAscii = 0
    ElseIf KeyAscii > 64 And KeyAscii < 91 Then
        'ok
    ElseIf KeyAscii > 96 And KeyAscii < 103 Then
        KeyAscii = KeyAscii - 32
    Else
        KeyAscii = 0
    End If
    
    If Len(txtAuth2.Text) > 3 Then
        txtAuth3.SetFocus
    End If
    
End Sub

Private Sub txtAuth3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 8 Then
        Exit Sub
    End If
    If KeyAscii > 47 And KeyAscii < 58 Then
    
    ElseIf KeyAscii < 65 Then
        KeyAscii = 0
    ElseIf KeyAscii > 64 And KeyAscii < 91 Then
        'ok
    ElseIf KeyAscii > 96 And KeyAscii < 103 Then
        KeyAscii = KeyAscii - 32
    Else
        KeyAscii = 0
    End If
    
    If Len(txtAuth3.Text) > 3 Then
        txtAuth4.SetFocus
    End If
    
End Sub

Private Sub txtAuth4_KeyPress(KeyAscii As Integer)

    If KeyAscii = 8 Then
        Exit Sub
    End If
    If KeyAscii > 47 And KeyAscii < 58 Then
    
    ElseIf KeyAscii < 65 Then
        KeyAscii = 0
    ElseIf KeyAscii > 64 And KeyAscii < 91 Then
        'ok
    ElseIf KeyAscii > 96 And KeyAscii < 103 Then
        KeyAscii = KeyAscii - 32
    Else
        KeyAscii = 0
    End If
    
    If Len(txtAuth4.Text) > 3 Then
        cmdOk.SetFocus
    End If
    
End Sub
