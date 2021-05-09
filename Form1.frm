VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   1320
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ping"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label IP 
      Alignment       =   2  'Center
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.Clear
Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Integer
   
     Call Ping(Text1.Text, ECHO)
   

    List1.AddItem GetStatusCode(ECHO.status)
   List1.AddItem ECHO.Address
   List1.AddItem ECHO.RoundTripTime & " ms"
   List1.AddItem ECHO.DataSize & " bytes"
   
   If Left$(ECHO.Data, 1) <> Chr$(0) Then
      pos = InStr(ECHO.Data, Chr$(0))
   List1.AddItem Left$(ECHO.Data, pos - 1)
   End If

   List1.AddItem ECHO.DataPointer

End Sub

Private Sub Command2_Click()
List1.Clear
End Sub
