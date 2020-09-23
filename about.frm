VERSION 5.00
Begin VB.Form about 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7845
   LinkTopic       =   "Form3"
   Picture         =   "about.frx":0000
   ScaleHeight     =   2685
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   6120
      Top             =   840
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click form for close about window."
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   1560
      Width           =   2445
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A simple program to connect and animation PLC program on ModBus+ TCP/IP Connection......."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   7575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Version : 1.50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   800
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Modbus TCP/IP Client with Animation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Click()
If Label5 = 1 Then
Me.Hide
End If
End Sub

Private Sub Form_Load()
Dim dildosya As String
Dim b
If Label5 = 0 Then
Timer1.Enabled = True
Label4.Visible = False
dildosya = INIYukle("PROGRAM", "Dil", App.Path & "\modsim.ini")
dilimiz = dildosya
loadlanguage App.Path & "\" & dildosya
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Label5 = 0
End Sub

Private Sub Timer1_Timer()
Me.Hide
Form1.Show
Timer1.Enabled = False
End Sub
