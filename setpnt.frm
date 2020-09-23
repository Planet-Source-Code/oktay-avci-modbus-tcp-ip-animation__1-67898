VERSION 5.00
Begin VB.Form setpnt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Value"
   ClientHeight    =   1305
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2895
   Icon            =   "setpnt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "0"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Digital Output Value"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label index 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label types 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "setpnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub Check1_Click()
Check1.Caption = Dil(37) & " " & Check1.Value
End Sub

Private Sub OKButton_Click()
Select Case types.Caption
    Case "dig"
        Form1.DOO(index.Caption).Caption = Check1.Value
        Me.Hide
    Case "ana"
        Form1.analog(index.Caption).Caption = Text1.Text
        Me.Hide
End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim a
a = Val(Text1.Text & Chr(KeyAscii))
If a > 32767 Then KeyAscii = 0
End Sub
