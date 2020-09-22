VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "BurP"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3720
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   120
   End
   Begin MSWinsockLib.Winsock sckTCP2 
      Index           =   0
      Left            =   960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make Active"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sckTCP 
      Index           =   0
      Left            =   480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label7 
      Caption         =   "(not to exceed 1 minute)"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "ms."
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Check Socket States Every:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Host Or IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Redirect to:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Listen On Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.c6d.net
Private Sub Command1_Click()
For a = 1 To 10
sckTCP(a).Close
sckTCP2(a).Close
DoEvents
Next a
sckTCP(0).Close
sckTCP(0).LocalPort = Text1
sckTCP(0).Listen
Timer1.Interval = Text4
End Sub

Private Sub Form_Load()
For a = 1 To 10
Load sckTCP(a)
Load sckTCP2(a)
DoEvents
Next a
Text4 = "10000"
'Form2.Show
End Sub

Private Sub sckTCP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
sckTCP(0).Close
For a = 1 To 10
If sckTCP(a).State = 0 Then
sckTCP2(a).Close
sckTCP2(a).RemoteHost = Text2
sckTCP2(a).RemotePort = Text3
sckTCP2(a).Connect
sckTCP(a).Accept requestID
DoEvents
a = 11
End If
DoEvents
Next a
sckTCP(0).Close
sckTCP(0).Listen
End Sub

Private Sub sckTCP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
If temp = "" Then sckTCP(Index).GetData temp
DoEvents
Do
If sckTCP2(Index).State = 7 Then
sckTCP2(Index).SendData temp
DoEvents
temp = ""
Exit Do
End If
If sckTCP(Index).State = 0 Then Exit Do
DoEvents
Loop
End Sub

Private Sub sckTCP2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
If temp = "" Then sckTCP2(Index).GetData temp
DoEvents
Do
If sckTCP(Index).State = 7 Then
sckTCP(Index).SendData temp
DoEvents
temp = ""
Exit Do
End If
If sckTCP(Index).State = 0 Then Exit Do
DoEvents
Loop
End Sub

Private Sub Timer1_Timer()
For a = 1 To 10
If sckTCP(a).State = 8 Then sckTCP(a).Close: sckTCP2(a).Close
If sckTCP(a).State = 9 Then sckTCP(a).Close: sckTCP2(a).Close
If sckTCP2(a).State = 8 Then sckTCP2(a).Close: sckTCP(a).Close
If sckTCP2(a).State = 9 Then sckTCP2(a).Close: sckTCP(a).Close
DoEvents
Next a
End Sub
