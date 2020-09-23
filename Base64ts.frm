VERSION 5.00
Begin VB.Form frmTestBase64 
   Caption         =   "Test Base64 Functions to Decode and Encode"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnterBase64Text 
      Caption         =   "Enter Base64 Text"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6960
      TabIndex        =   8
      Top             =   180
      Width           =   2265
   End
   Begin VB.CommandButton cmdEnterTestText 
      Caption         =   "Enter Test Text"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   135
      TabIndex        =   7
      Top             =   180
      Width           =   2205
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   255
      Left            =   5070
      TabIndex        =   6
      Top             =   -15
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7245
      TabIndex        =   5
      Top             =   5775
      Width           =   1995
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4910
      TabIndex        =   4
      Top             =   5775
      Width           =   1995
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "ENCODE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2575
      TabIndex        =   3
      Top             =   5760
      Width           =   1995
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "DECODE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   240
      TabIndex        =   2
      Top             =   5775
      Width           =   1995
   End
   Begin VB.TextBox txtWork 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   780
      Width           =   9120
   End
   Begin VB.Label Label 
      Caption         =   "Press Tab after entering text to validate it"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2460
      TabIndex        =   9
      Top             =   315
      Width           =   4335
   End
   Begin VB.Label lblFormName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "frmBase64"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4020
      TabIndex        =   1
      Top             =   15
      Width           =   1140
   End
End
Attribute VB_Name = "frmTestBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Randomize Timer
End Sub

Private Sub cmdAbout_Click()
    MsgBox "Brought to you by Caravela Books at http://caravelabooks.com ", , "CARAVELA BOOKS"
End Sub

Private Sub cmdClear_Click()
    txtWork.Text = ""       'clear the main text box
End Sub

Private Sub cmdDecode_Click() 'decode text from base64
    txtWork.Text = sBase64Decode(txtWork.Text)
End Sub

Private Sub cmdEncode_Click()    'encode text into base64
    
    Dim rc As Integer
    
    If Len(txtWork.Text) > 20 And IsBase64(txtWork.Text) Then
        rc = MsgBox("The text appears to be in base64 format already." _
                    & " Okay to go ahead and encode it anyway?", vbOKCancel, "DOUBLE ENCODE?")
        If rc = vbCancel Then Exit Sub
    End If
    
    txtWork.Text = sBase64Encode(txtWork.Text) 'get base64 encoded text
    
End Sub

Private Sub cmdEnterBase64Text_Click()
    
    Dim s As String
    
    s = "PGh0bWw+DQo8Ym9keT4NCjxwPkhleSw8L3A+DQo8cD5DaGVjayBoZXIgb3V0"
    s = s & "IHNoZXMgc28gaG90ISEgaW1hZ2UgaXMgbG9hZGluZy4uLjxicj4NCiAgPGEg"
    s = s & "aHJlZj0iaHR0cDovL2hpbHRvbmZvcnlvdS5iaXoiPjxpbWcgc3JjPSJodHRw"
    s = s & "Oi8vaGlsdG9uZm9yeW91LmJpei8xLmpwZyIgYm9yZGVyPSIwIj48L2E+PGJy"
    s = s & "Pg0KICBJdHMgYWJzb2x1dGVseSBmcmVlIE5PIHB1cmNoYXNlIHJlcXVpcmVk"
    s = s & "Ljxicj4NCiAgPGJyPg0KICA8YnI+DQogIDxicj4NCiAgPGJyPg0KICA8YnI+"
    s = s & "DQogIDxicj4NCiAgPGEgaHJlZj0iaHR0cDovL2hpbHRvbmZvcnlvdS5iaXov"
    s = s & "b3V0LnBocCI+T3B0IG91dCBmcm9tIG1haWxpbmcgbGlzdDwvYT4gPC9wPg0K"
    s = s & "PC9ib2R5Pg0KPC9odG1sPg0K"

    txtWork.Text = s    'put base64 in text box for debug
    
End Sub

Private Sub cmdEnterTestText_Click()
    
    Dim s As String
    
    s = "Today is " & Format(Date, "long date") & " at " _
            & Format(Now, "long time") & vbNewLine _
            & "The quick brown fox jumps over the lazy dog." & vbNewLine _
            & "Caravela Books " _
            & Mid$("#1234", 1, CInt(Rnd * 5) + 1) 'add random number of bytes
            
    txtWork.Text = s    'load the text box with the test text
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

