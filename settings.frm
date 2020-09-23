VERSION 5.00
Begin VB.Form settings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Settings"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   2775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   120
      Pattern         =   "*.lng"
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Languages :"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Hide
End Sub



Private Sub File1_Click()
OKButton.Enabled = True
End Sub

Private Sub Form_Load()
Dim i
File1.Path = App.Path
For i = 0 To File1.ListCount
File1.Selected(i) = True
    If File1.filename = dilimiz Then
        File1.Selected(i) = True
        Exit Sub
    End If
Next i
End Sub

Private Sub OKButton_Click()
loadlanguage App.Path & "\" & File1.filename
dilimiz = File1.filename
Me.Hide
End Sub
