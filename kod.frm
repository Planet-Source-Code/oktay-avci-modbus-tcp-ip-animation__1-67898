VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form kod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Animation Code Editor"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   Icon            =   "kod.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8493
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      TextRTF         =   $"kod.frx":3502
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "kod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
oascript = Text1.Text
On Error GoTo 10
Form1.Script.AddCode oascript
Form1.RichBox1.TextRTF = Text1.TextRTF
Me.Hide
10
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
SendMessageLong Text1.hWnd, EM_SETTARGETDEVICE, 0, 1 ' Text Kaydýrma yok.. Lastdaki 1 , 0 yapýlýrsa Wrap yani kaydýrma var
Text1.TextRTF = Form1.RichBox1.TextRTF
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Colorize Text1, RGB(78, 166, 13), RGB(128, 64, 0), &H800000
Text1.SelColor = 0
End If
MenageRTF KeyAscii, Text1
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
Colorize Text1, RGB(78, 166, 13), RGB(128, 64, 0), &H800000
Text1.SelColor = 0
End If
End Sub
