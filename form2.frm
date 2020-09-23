VERSION 5.00
Begin VB.Form form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Object Properties"
   ClientHeight    =   5055
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3885
   Icon            =   "form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar scrol 
      Height          =   255
      Left            =   120
      Max             =   100
      TabIndex        =   24
      Top             =   4080
      Width           =   3615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Apply"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Border Style"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3480
         Width           =   3255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Load"
         Height          =   255
         Left            =   2880
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   2760
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Text            =   "Text4"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text6"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text7"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text8"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "Text8"
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Object Type : "
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Index : "
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Left :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Top :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Width :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Height :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Border Color :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Color :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Text : "
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
With OBJECT(Val(Text2.Text))
    .OBJECT = Text1
    .Left = Val(Text3)
    .Top = Val(Text4)
    .Width = Val(Text5)
    .Height = Val(Text6)
    .BorderColor = Val(Text7)
    .MainColor = Val(Text8)
    .Text = Text9
    .Clicked = Check1.Value
End With
ReBuildScreen Form1.Picture1
Form1.tool.Buttons(11).Enabled = False
Form1.tool.Buttons(12).Enabled = False
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Command3_Click()
Form1.cm.Color = Val(Text7.Text)
Form1.cm.ShowColor
Text7.Text = Form1.cm.Color
End Sub

Private Sub Command4_Click()
Form1.cm.Color = Val(Text8.Text)
Form1.cm.ShowColor
Text8.Text = Form1.cm.Color
End Sub

Private Sub Command5_Click()
If Val(Text2.Text) > UBound(OBJECT) Then
    MsgBox "Girilen Kayýt Bulunamadý..."
Else
    OBJECTozellikyukle
End If
End Sub

Sub OBJECTozellikyukle()
        With OBJECT(Val(Text2.Text))
            Text1 = .OBJECT
            Text3 = .Left
            Text4 = .Top
            Text5 = .Width
            Text6 = .Height
            Text7 = .BorderColor
            Text8 = .MainColor
            Text9 = .Text
            If .Clicked = True Then
            Check1.Value = 1
            Else
            Check1.Value = 0
            End If
        End With
End Sub

Private Sub Command6_Click()
With OBJECT(Val(Text2.Text))
    .OBJECT = Text1
    .index = Val(Text2.Text)
    .Left = Val(Text3)
    .Top = Val(Text4)
    .Width = Val(Text5)
    .Height = Val(Text6)
    .BorderColor = Val(Text7)
    .MainColor = Val(Text8)
    .Text = Text9
    .Clicked = Check1.Value
End With
ReBuildScreen Form1.Picture1
Form1.tool.Buttons(11).Enabled = False
Form1.tool.Buttons(12).Enabled = False
End Sub

Private Sub scrol_Change()
Text2.Text = scrol.Value
OBJECTozellikyukle
End Sub
