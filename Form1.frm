VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modbus TCP/IP Client with Animation v1.50"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13650
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   13650
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichBox1 
      Height          =   375
      Left            =   9120
      TabIndex        =   138
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form1.frx":08CA
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
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12840
      TabIndex        =   136
      Text            =   "50"
      Top             =   75
      Width           =   615
   End
   Begin VB.Frame Frame5 
      Caption         =   "Connection Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   10320
      TabIndex        =   127
      Top             =   6360
      Width           =   3135
      Begin VB.TextBox IP 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   131
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Portm 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   130
         Text            =   "502"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Connect"
         Height          =   375
         Left            =   120
         TabIndex        =   129
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox timout 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2280
         TabIndex        =   128
         Text            =   "2"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "IP Number :"
         Height          =   255
         Left            =   120
         TabIndex        =   135
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Port :"
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label17 
         Caption         =   "Time - Out (sn) :"
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label BagDurum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   1920
         Width           =   2895
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "To PLC : (40017)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10320
      TabIndex        =   105
      Top             =   4920
      Width           =   3135
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "29:32"
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "25:28"
         Height          =   255
         Left            =   120
         TabIndex        =   124
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "21:24"
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "17:20"
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   360
         Width           =   615
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   2520
         TabIndex        =   121
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   1920
         TabIndex        =   120
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   1320
         TabIndex        =   119
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   720
         TabIndex        =   118
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   2520
         TabIndex        =   117
         Top             =   840
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   1920
         TabIndex        =   116
         Top             =   840
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   1320
         TabIndex        =   115
         Top             =   840
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   720
         TabIndex        =   114
         Top             =   840
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   2520
         TabIndex        =   113
         Top             =   600
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   1920
         TabIndex        =   112
         Top             =   600
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   1320
         TabIndex        =   111
         Top             =   600
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   720
         TabIndex        =   110
         Top             =   600
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   2520
         TabIndex        =   109
         Top             =   360
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   1920
         TabIndex        =   108
         Top             =   360
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   1320
         TabIndex        =   107
         Top             =   360
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   720
         TabIndex        =   106
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "From PLC : (40001)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10320
      TabIndex        =   84
      Top             =   3480
      Width           =   3135
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   104
         Top             =   360
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   103
         Top             =   360
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   102
         Top             =   360
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   101
         Top             =   360
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   100
         Top             =   600
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   99
         Top             =   600
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   98
         Top             =   600
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2520
         TabIndex        =   97
         Top             =   600
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   720
         TabIndex        =   96
         Top             =   840
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   1320
         TabIndex        =   95
         Top             =   840
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1920
         TabIndex        =   94
         Top             =   840
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   2520
         TabIndex        =   93
         Top             =   840
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   720
         TabIndex        =   92
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   1320
         TabIndex        =   91
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   1920
         TabIndex        =   90
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label analog 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   2520
         TabIndex        =   89
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "1 : 4"
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "5 : 8"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "8 :12"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "13:16"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "To PLC : (00033)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10320
      TabIndex        =   45
      Top             =   2040
      Width           =   3135
      Begin VB.Label Label8 
         Caption         =   "57:64"
         Height          =   255
         Left            =   120
         TabIndex        =   83
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label7 
         Caption         =   "49:56"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   840
         Width           =   450
      End
      Begin VB.Label DOINT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   81
         Top             =   360
         Width           =   495
      End
      Begin VB.Label DOINT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   2520
         TabIndex        =   80
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "41:48"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label5 
         Caption         =   "33:40"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   450
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   600
         TabIndex        =   77
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   840
         TabIndex        =   76
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   34
         Left            =   1080
         TabIndex        =   75
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   35
         Left            =   1320
         TabIndex        =   74
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   1560
         TabIndex        =   73
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   1800
         TabIndex        =   72
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   2040
         TabIndex        =   71
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   2280
         TabIndex        =   70
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   600
         TabIndex        =   69
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   41
         Left            =   840
         TabIndex        =   68
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   42
         Left            =   1080
         TabIndex        =   67
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   43
         Left            =   1320
         TabIndex        =   66
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   44
         Left            =   1560
         TabIndex        =   65
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   45
         Left            =   1800
         TabIndex        =   64
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   46
         Left            =   2040
         TabIndex        =   63
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   47
         Left            =   2280
         TabIndex        =   62
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   48
         Left            =   600
         TabIndex        =   61
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   49
         Left            =   840
         TabIndex        =   60
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   50
         Left            =   1080
         TabIndex        =   59
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   51
         Left            =   1320
         TabIndex        =   58
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   52
         Left            =   1560
         TabIndex        =   57
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   53
         Left            =   1800
         TabIndex        =   56
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   54
         Left            =   2040
         TabIndex        =   55
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   55
         Left            =   2280
         TabIndex        =   54
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   56
         Left            =   600
         TabIndex        =   53
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   57
         Left            =   840
         TabIndex        =   52
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   58
         Left            =   1080
         TabIndex        =   51
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   59
         Left            =   1320
         TabIndex        =   50
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   60
         Left            =   1560
         TabIndex        =   49
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   61
         Left            =   1800
         TabIndex        =   48
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   62
         Left            =   2040
         TabIndex        =   47
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   63
         Left            =   2280
         TabIndex        =   46
         Top             =   1080
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "From PLC : (00001)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10320
      TabIndex        =   6
      Top             =   600
      Width           =   3135
      Begin VB.Label Label4 
         Caption         =   "17:24"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label3 
         Caption         =   "8:16"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   600
         Width           =   375
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   42
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   41
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   40
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   39
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   38
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   37
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   36
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   35
         Top             =   360
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   34
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   33
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   1080
         TabIndex        =   32
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   31
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   30
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   1800
         TabIndex        =   29
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   2040
         TabIndex        =   28
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   2280
         TabIndex        =   27
         Top             =   600
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   600
         TabIndex        =   26
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   840
         TabIndex        =   25
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   1080
         TabIndex        =   24
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   1320
         TabIndex        =   23
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   1560
         TabIndex        =   22
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   1800
         TabIndex        =   21
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   2040
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   2280
         TabIndex        =   19
         Top             =   840
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   600
         TabIndex        =   18
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   840
         TabIndex        =   17
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   1080
         TabIndex        =   16
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   1320
         TabIndex        =   15
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   1560
         TabIndex        =   14
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   1800
         TabIndex        =   13
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   2040
         TabIndex        =   12
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label DOO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   2280
         TabIndex        =   11
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "1:8"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "24:32"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label DOINT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.Label DOINT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9120
      Top             =   4320
   End
   Begin MSComctlLib.StatusBar dur 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   8760
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1323
            MinWidth        =   1323
            Text            =   "X : 0"
            TextSave        =   "X : 0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1323
            MinWidth        =   1323
            Text            =   "Y : 0"
            TextSave        =   "Y : 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Active Process : there are nothing yet..."
            TextSave        =   "Active Process : there are nothing yet..."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   2
            TextSave        =   "18.02.2007"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tool 
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ResimListe"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Line"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Rectangle"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Circle"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Label"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Text"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Button"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Active Color"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Selection"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Move"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Delete Object"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Object Properties"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Run"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cm 
      Left            =   9240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Animasyon Projesi (*.anp)|*.anp"
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      DrawWidth       =   2
      Height          =   8175
      Left            =   120
      ScaleHeight     =   8115
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   600
      Width           =   10095
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Left            =   9000
         Top             =   2880
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   9000
         Top             =   2400
      End
      Begin MSScriptControlCtl.ScriptControl Script 
         Left            =   9000
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin VB.TextBox met 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComctlLib.ImageList ResimListe 
         Left            =   9000
         Top             =   1800
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0966
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0A78
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0B8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0C9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0DAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0EC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0FD2
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":10E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":11F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1308
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":141A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":152C
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":187E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   9120
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemoteHost      =   "85.16.23.23"
         RemotePort      =   502
      End
      Begin VB.Shape tarama 
         BorderStyle     =   3  'Dot
         Height          =   495
         Left            =   4680
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin Project1.OBJE OBJECT1 
      Height          =   405
      Left            =   13680
      TabIndex        =   5
      Top             =   120
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   714
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Con. Speed (ms) :"
      Height          =   255
      Left            =   11400
      TabIndex        =   137
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Color :"
      Height          =   255
      Left            =   10080
      TabIndex        =   126
      Top             =   120
      Width           =   855
   End
   Begin VB.Label renk 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11040
      TabIndex        =   2
      Top             =   75
      Width           =   255
   End
   Begin VB.Menu dosya 
      Caption         =   "&File"
      Begin VB.Menu yeni 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu dosyac 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu kaydet 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu farklikaydet 
         Caption         =   "Save as.."
      End
      Begin VB.Menu bo 
         Caption         =   "-"
      End
      Begin VB.Menu resimyap 
         Caption         =   "Save as Bitmap"
         Shortcut        =   ^E
      End
      Begin VB.Menu boo 
         Caption         =   "-"
      End
      Begin VB.Menu cikis 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu duzen 
      Caption         =   "&Edit"
      Begin VB.Menu kes 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu kopyala 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu yapistir 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu booo 
         Caption         =   "-"
      End
      Begin VB.Menu kodedit 
         Caption         =   "Code Editor"
      End
   End
   Begin VB.Menu sagtus 
      Caption         =   "Right"
      Visible         =   0   'False
      Begin VB.Menu kopyam 
         Caption         =   "Cut"
         Index           =   1
      End
      Begin VB.Menu kopyam 
         Caption         =   "Copy"
         Index           =   2
      End
      Begin VB.Menu kopyam 
         Caption         =   "Paste"
         Index           =   3
      End
      Begin VB.Menu bos1 
         Caption         =   "-"
      End
      Begin VB.Menu sagtusum 
         Caption         =   "Delete Object"
         Index           =   0
      End
      Begin VB.Menu sagtusum 
         Caption         =   "Move Object"
         Index           =   1
      End
      Begin VB.Menu sagtusum 
         Caption         =   "Scale Object"
         Index           =   2
      End
      Begin VB.Menu sagtusum 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu sagtusum 
         Caption         =   "Object Properties"
         Index           =   4
      End
   End
   Begin VB.Menu ayarlar 
      Caption         =   "Settings"
      Begin VB.Menu progsettings 
         Caption         =   "Program Settings"
      End
   End
   Begin VB.Menu yardim 
      Caption         =   "&Help"
      Begin VB.Menu hakkinda 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Status As String
Dim myMouse As Boolean
Dim blank As Long
Dim SelectedOBJECT As Long
Dim kont As Long
Dim ActiveKont As Boolean
Dim kopyaaktif As Boolean
Dim RightClickIndex
Dim SlideBlank
'Taþýma ve secim
Dim mousx
Dim mousy
Dim SelectAccess As Boolean
'Scalingdýrma
Dim Scaling As Boolean
Dim ScalingObj As Long
Dim ScalingAccess As Boolean
Dim ourObject
Dim Where

'Baðlantý Tanýmlamalarý
Dim MbusQuery(11) As Byte
Public MbusResponse As String
Dim MbusByteArray(255) As Long
Public MbusRead As Boolean
Public analogout As Boolean
Public DigitalOut As Boolean

Public MbusWrite As Boolean
Public analogin As Boolean
Public digitalin As Boolean
Dim ModbusTimeOut As Integer
Dim ModbusWait As Boolean

Private Sub analog_Change(index As Integer)
On Error Resume Next
Script.ExecuteStatement "ACO" & index + 1 & "_change"
End Sub

Private Sub analog_Click(index As Integer)
setpointform "ana", analog(index).Caption, index
End Sub

Function BlankData() As Long
Dim i
For i = 0 To 100
    If OBJECT(i).OBJECT = "" Then BlankData = i: Exit Function
Next i
End Function



Private Sub hakkinda_Click()
about.Label5 = 1
about.Show 1, Form1
If about.Label4.Visible = False Then about.Label4.Visible = True
End Sub

Private Sub kes_Click()
If tool.Buttons(9).Value = tbrPressed Then
    If SelectedOBJECT = -1 Then Exit Sub
    kopyamiz = OBJECT(SelectedOBJECT)
    OBJECT(SelectedOBJECT).OBJECT = ""
    tool.Buttons(9).Value = tbrUnpressed
    SelectedOBJECT = -1
    Status = ""
    ReBuildScreen Picture1
    kopyaaktif = True
Else
    If RightClickIndex = -1 Then Exit Sub
    kopyamiz = OBJECT(RightClickIndex)
    OBJECT(RightClickIndex).OBJECT = ""
    RightClickIndex = -1
    Status = ""
    ReBuildScreen Picture1
    kopyaaktif = True
End If
End Sub

Private Sub kopyala_Click()
If tool.Buttons(9).Value = tbrPressed Then
    If SelectedOBJECT = -1 Then Exit Sub
    kopyamiz = OBJECT(SelectedOBJECT)
    tool.Buttons(9).Value = tbrUnpressed
    SelectedOBJECT = -1
    Status = ""
    ReBuildScreen Picture1
    kopyaaktif = True
Else
    If RightClickIndex = -1 Then Exit Sub
    kopyamiz = OBJECT(RightClickIndex)
    RightClickIndex = -1
    Status = ""
    ReBuildScreen Picture1
    kopyaaktif = True
End If
End Sub

Private Sub kopyam_Click(index As Integer)
Select Case index
    Case 1 ' kes
        kes_Click
    Case 2 ' kopyala
        kopyala_Click
    Case 3 ' yapýþtýr
        yapistir_Click
End Select
End Sub

Private Sub progsettings_Click()
settings.Show 1, Form1
End Sub

Private Sub tool_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.index
    Case 1
        Status = "Line"
        dur.Panels(3).Text = Dil(0) & " " & Dil(1)
        ClearSelected
    Case 2
        Status = "Rectangle"
        dur.Panels(3).Text = Dil(0) & " " & Dil(2)
        ClearSelected
    Case 3
        Status = "Circle"
        dur.Panels(3).Text = Dil(0) & " " & Dil(3)
        ClearSelected
    Case 4
        Status = "Label"
        dur.Panels(3).Text = Dil(0) & " " & Dil(4)
        ClearSelected
    Case 5
        Status = "Text"
        dur.Panels(3).Text = Dil(0) & " " & Dil(5)
        ClearSelected
    Case 6
        Status = "Button"
        dur.Panels(3).Text = Dil(0) & " " & Dil(6)
        ClearSelected
    'Case 7 Seperatör
    Case 8
        cm.Flags = &H200
        cm.Color = renk.BackColor
        cm.ShowColor
        renk.BackColor = cm.Color
        ClearSelected
    Case 9
        If tool.Buttons(9).Value = tbrUnpressed Then
            Status = ""
            dur.Panels(3).Text = Dil(0) & " " & Dil(16)
            ClearSelected
            ReBuildScreen Picture1
        Else
            Status = "Selection"
            dur.Panels(3).Text = Dil(0) & " " & Dil(17)
            ReBuildScreen Picture1
        End If
    Case 10
            Status = "Move"
            dur.Panels(3).Text = Dil(0) & " " & Dil(18)
            tool.Buttons.Item(9).Value = tbrUnpressed
            ClearSelected
            ReBuildScreen Picture1
    Case 11
        DeleteObjects
        ClearSelected
    Case 12
        ozellikpenceresi SelectedOBJECT
    Case 14
        If Winsock1.State <> 7 Then
            MsgBox Dil(19), vbInformation
            Exit Sub
        End If
        Timer3.Interval = Val(Text1.Text)
        Timer3.Enabled = True
        tool.Buttons(14).Enabled = False
        tool.Buttons(15).Enabled = True
    Case 15
        Timer3.Enabled = False
        tool.Buttons(15).Enabled = False
        tool.Buttons(14).Enabled = True
End Select
End Sub



Private Sub cikis_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Command2.Caption = Dil(20) Then
    baglan
Exit Sub
Else
If Timer3.Enabled = True Then
    Timer3.Enabled = False
    tool.Buttons(14).Enabled = True
    tool.Buttons(15).Enabled = False
End If
    Winsock1.Close
    Command2.Caption = Dil(20)
    BagDurum = IP.Text + Dil(21)
    BagDurum.BackColor = &HFF
End If
End Sub



Private Sub DOINT_Change(index As Integer)
Dim ver, i
Select Case index
Case 0
ver = 0
    For i = 0 To 15
        ver = Val(DOINT(index)) And (2 ^ i)
        If ver <> 0 Then
            DOO(i).Caption = 1
        Else
            DOO(i).Caption = 0
        End If
    Next i
Case 1
ver = 0
    For i = 16 To 31
        ver = Val(DOINT(index)) And (2 ^ (i - 16))
        If ver <> 0 Then
            DOO(i).Caption = 1
        Else
            DOO(i).Caption = 0
        End If
    Next i
End Select
End Sub

Private Sub DOO_Change(index As Integer)
Dim i, dat1
On Error Resume Next
Script.ExecuteStatement "DDO" & index + 1 & "_change"
dat1 = 0
Select Case index
    Case 0 To 15
        For i = 0 To 15 Step 1
            If DOO(i).Caption = 1 Then
               dat1 = Val(dat1) + (2 ^ i)
            End If
        Next i
        DOINT(0).Caption = dat1
    Case 16 To 31
        dat1 = 0
        For i = 16 To 31 Step 1
            If DOO(i).Caption = 1 Then
               dat1 = Val(dat1) + (2 ^ (i - 16))
            End If
        Next i
        DOINT(1).Caption = dat1
    Case 32 To 47
        dat1 = 0
        For i = 32 To 47 Step 1
            If DOO(i).Caption = 1 Then
               dat1 = Val(dat1) + (2 ^ (i - 32))
            End If
        Next i
        DOINT(2).Caption = dat1
    Case 48 To 63
        dat1 = 0
        For i = 48 To 63 Step 1
            If DOO(i).Caption = 1 Then
               dat1 = Val(dat1) + (2 ^ (i - 48))
            End If
        Next i
        DOINT(3).Caption = dat1
End Select
End Sub

Private Sub DOO_Click(index As Integer)
setpointform "dig", DOO(index).Caption, index
End Sub

Private Sub dosyac_Click()
cm.filename = ""
cm.Filter = Dil(84)
cm.ShowOpen
If cm.filename = "" Then Exit Sub
dosyaac cm.filename
LoadedFile = cm.filename
End Sub

Private Sub farklikaydet_Click()
cm.filename = ""
cm.Filter = Dil(84)
cm.ShowSave
If cm.filename = "" Then Exit Sub
dosyakaydet cm.filename
LoadedFile = cm.filename
End Sub

Private Sub Form_Load()
SelectedOBJECT = -1
Script.AddObject "OBJECT", OBJECT1
Script.AddObject "timer1", Timer1
Script.AddObject "DDO", DOO
End Sub

Private Sub Form_Unload(Cancel As Integer)
INIYaz "PROGRAM", "Dil", dilimiz, App.Path & "\modsim.ini"
End
End Sub

Private Sub kaydet_Click()
cm.filename = ""
cm.Filter = Dil(84)
If LoadedFile <> "" Then
dosyakaydet LoadedFile
Else
cm.ShowSave
If cm.filename = "" Then Exit Sub
dosyakaydet cm.filename
LoadedFile = cm.filename
End If
End Sub

Private Sub kodedit_Click()
kod.Show 1, Form1
End Sub

Private Sub met_GotFocus()
met.SelStart = 0
met.SelLength = Len(met.Text)
End Sub

Private Sub met_KeyPress(KeyAscii As Integer)
Dim a
a = KeyAscii
If a = 13 Then
    OBJECT(Val(met.Tag)).Text = met.Text
    ReBuildScreen Picture1
    met.Visible = False
ElseIf a = 27 Then
    met.Visible = False
End If
End Sub
Function scripdosyaadi(projedosyasi) As String
Dim a
scripdosyaadi = Mid$(projedosyasi, 1, Len(projedosyasi) - 1) & "s"
End Function

Sub dosyakaydet(dosyaadi)
Dim i
Open dosyaadi For Random As #1
    For i = 1 To 101 Step 1
        Put #1, i, OBJECT(i - 1)
    Next i
Close #1
Form1.RichBox1.SaveFile scripdosyaadi(dosyaadi)
End Sub

Sub dosyaac(dosyaadi)
Dim i, a
oascript = ""
Open dosyaadi For Random As #1
    For i = 1 To 101 Step 1
        Get #1, i, OBJECT(i - 1)
    Next i
Close #1
kod.Text1.LoadFile scripdosyaadi(dosyaadi)
RichBox1.LoadFile scripdosyaadi(dosyaadi)
oascript = RichBox1.Text
Script.AddCode oascript
ReBuildScreen Picture1
End Sub


Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If SelectedOBJECT <> -1 Then
Select Case KeyCode
    Case 37 ' Sol Ok tuþu
        OBJECT(SelectedOBJECT).Left = OBJECT(SelectedOBJECT).Left - 10
    Case 38 ' Yukarý Ok tuþu
        OBJECT(SelectedOBJECT).Top = OBJECT(SelectedOBJECT).Top - 10
    Case 39 ' Sað Ok tuþu
        OBJECT(SelectedOBJECT).Left = OBJECT(SelectedOBJECT).Left + 10
    Case 40 ' Aþaðý Ok tuþu
        OBJECT(SelectedOBJECT).Top = OBJECT(SelectedOBJECT).Top + 10
    Case 46 ' Delete Tuþu
        If prog.Selected <> "" Then
            DeleteObjects
            ClearSelected
        End If
End Select
ReBuildScreen Picture1
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If FindObject(x, y) = -1 Then SelectedOBJECT = -1
If Button = 2 Then
    RightClickIndex = FindObject(x, y)
    If RightClickIndex = -1 Then
        If kopyaaktif = True Then
            Dim i
                For i = 0 To 4
                    sagtusum(i).Visible = False
                Next i
            PopupMenu sagtus
        End If
        Exit Sub
    Else
        dur.Panels(3).Text = Dil(22) & OBJECT(RightClickIndex).index
        
                For i = 0 To 4
                    sagtusum(i).Visible = True
                Next i
        makeSelect RightClickIndex
        PopupMenu sagtus
    End If
Else
If Scaling = True Then
ScalingAccess = True
End If
blank = BlankData()
myMouse = True
Select Case Status
    Case ""
        tool.Buttons(11).Enabled = False
        tool.Buttons(12).Enabled = False
        kont = FindObject(x, y)
        If met.Visible = True Then
            OBJECT(Val(met.Tag)).Text = met.Text
            met.Visible = False
            ReBuildScreen Picture1
            Exit Sub
        End If
        If kont <> -1 And Button = 1 Then
            If OBJECT(kont).OBJECT = "Button" And SelectedOBJECT = -1 Then
                OBJECT(kont).Clicked = True
                ActiveKont = True
            End If
            If OBJECT(kont).OBJECT = "Text" And SelectedOBJECT = -1 Then
                met.Top = OBJECT(kont).Top
                met.Left = OBJECT(kont).Left
                met.Text = OBJECT(kont).Text
                met.Width = OBJECT(kont).Width
                met.Height = OBJECT(kont).Height
                met.Visible = True
                met.Tag = kont
                met.SetFocus
            End If
        End If
        ReBuildScreen Picture1
    Case "Line"
        OBJECT(blank).OBJECT = Status
        OBJECT(blank).index = blank
        OBJECT(blank).Left = x
        OBJECT(blank).Top = y
        OBJECT(blank).MainColor = renk.BackColor
    Case "Rectangle"
        OBJECT(blank).OBJECT = Status
        OBJECT(blank).index = blank
        OBJECT(blank).Left = x
        OBJECT(blank).Top = y
        OBJECT(blank).MainColor = renk.BackColor
    Case "Circle"
        OBJECT(blank).OBJECT = Status
        OBJECT(blank).index = blank
        OBJECT(blank).Left = x
        OBJECT(blank).Top = y
        OBJECT(blank).MainColor = renk.BackColor
    Case "Button"
        OBJECT(blank).OBJECT = Status
        OBJECT(blank).index = blank
        OBJECT(blank).Left = x
        OBJECT(blank).Top = y
        OBJECT(blank).Text = "Button"
        OBJECT(blank).Clicked = False
        OBJECT(blank).MainColor = 13421772
    Case "Label"
        OBJECT(blank).OBJECT = Status
        OBJECT(blank).index = blank
        OBJECT(blank).Left = x
        OBJECT(blank).Top = y
        OBJECT(blank).Text = "Label"
    Case "Text"
        OBJECT(blank).OBJECT = Status
        OBJECT(blank).index = blank
        OBJECT(blank).Left = x
        OBJECT(blank).Top = y
        OBJECT(blank).Text = "Text"
    Case "Move"
        SelectedOBJECT = FindObject(x, y)
        If SelectedOBJECT = -1 Then
            ReBuildScreen Picture1
            tool.Buttons(11).Enabled = False
            tool.Buttons(12).Enabled = False
            prog.Selected = ""
        Exit Sub
        End If
        mousx = x
        mousy = y
        makeSelect SelectedOBJECT
    Case "Selection"
        SelectedOBJECT = FindObject(x, y)
        If SelectedOBJECT = -1 Then
           tarama.Top = y
           tarama.Left = x
           tarama.Width = 10
           tarama.Height = 10
           tarama.Visible = True
           SelectAccess = True
        Else
            tool.Buttons(11).Enabled = True
            tool.Buttons(12).Enabled = True
            Selectedeekle SelectedOBJECT
            makeSelect SelectedOBJECT
        End If
End Select
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim a, b, c, d
If SelectAccess = True Then
    If x < tarama.Left + 50 Or y < tarama.Top + 50 Then Exit Sub
    tarama.Width = x - tarama.Left
    tarama.Height = y - tarama.Top
    Exit Sub
End If
SlideBlank = FindObject(x, y)
If SlideBlank <> -1 Then
        dur.Panels(3).Text = Dil(22) & SlideBlank
End If
If ScalingAccess = True Then
If ScalingObj = -1 Or ourObject = "" Or Where = "" Then ScalingAccess = False: Exit Sub
ReScale ScalingObj, x, y, ourObject, Where
ReBuildScreen Picture1
Exit Sub
End If
If tool.Buttons.Item(9).Value <> tbrPressed Then
ReBuildScreen Picture1
End If

If SelectedOBJECT <> -1 Then
    makeSelect SelectedOBJECT
    a = ObjectPoint(SelectedOBJECT, x, y, b, c)
    If a = True Then
        Scaling = True
        ourObject = b
        Where = c
        ScalingObj = SelectedOBJECT
        Select Case b
        Case "DA"
            Picture1.MousePointer = 7
        Case "DI"
            Picture1.MousePointer = 8
        End Select
    Else
        Picture1.MousePointer = 0
        Scaling = False
        ScalingObj = -1
        ourObject = ""
        Where = ""
    End If
End If
If tool.Buttons.Item(9).Value = tbrPressed Then Exit Sub
dur.Panels(1).Text = "X : " & x
dur.Panels(2).Text = "Y : " & y
If myMouse = True Then
If OBJECT(blank).Width < 0 Or OBJECT(blank).Height < 0 Then
OBJECT(blank).Width = 10
OBJECT(blank).Height = 10
Exit Sub
End If
Select Case Status
    Case "Circle"
        OBJECT(blank).Height = Abs(y - OBJECT(blank).Top)
        OBJECT(blank).Width = Abs(y - OBJECT(blank).Top)
    Case "Move"
    If SelectedOBJECT = -1 Then Exit Sub
        If OBJECT(SelectedOBJECT).OBJECT = "Circle" Then
            OBJECT(SelectedOBJECT).Left = x
            OBJECT(SelectedOBJECT).Top = y
            Exit Sub
        End If
            a = OBJECT(SelectedOBJECT).Left
            b = OBJECT(SelectedOBJECT).Top
            c = x - mousx
            d = y - mousy
            mousx = x
            mousy = y
            OBJECT(SelectedOBJECT).Left = a + c
            OBJECT(SelectedOBJECT).Top = b + d
    Case "Line"
        OBJECT(blank).Width = x - OBJECT(blank).Left
        OBJECT(blank).Height = y - OBJECT(blank).Top
    Case "Rectangle"
        OBJECT(blank).Width = x - OBJECT(blank).Left
        OBJECT(blank).Height = y - OBJECT(blank).Top
    Case "Button"
        OBJECT(blank).Width = x - OBJECT(blank).Left
        OBJECT(blank).Height = y - OBJECT(blank).Top
    Case "Label"
        OBJECT(blank).Width = x - OBJECT(blank).Left
        OBJECT(blank).Height = y - OBJECT(blank).Top
    Case "Text"
        OBJECT(blank).Width = x - OBJECT(blank).Left
        OBJECT(blank).Height = y - OBJECT(blank).Top
End Select
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
myMouse = False
If SelectAccess = True Then
    taralialansec tarama.Left, tarama.Top, x, y
    SelectAccess = False
    tarama.Visible = False
End If
ScalingAccess = False
If tool.Buttons(9).Value = tbrPressed Then
Status = "Selection"
Else
Status = ""
End If
If ActiveKont = True Then
    OBJECT(kont).Clicked = False
    On Error Resume Next
    Script.ExecuteStatement "Button" & kont & "_Click"
    ReBuildScreen Picture1
    ActiveKont = False
End If
dur.Panels(3).Text = Dil(0) & Dil(16) & Status
GoTo 12
11
MsgBox Dil(23)
ActiveKont = False
12
End Sub


Private Sub resimyap_Click()
cm.filename = ""
cm.Filter = Dil(81)
cm.ShowSave
If cm.filename = "" Then GoTo 10
Picture1.Picture = Picture1.Image
SavePicture Picture1.Picture, cm.filename
10
Picture1.Cls
Picture1.Picture = LoadPicture("")
ReBuildScreen Picture1
cm.Filter = Dil(84)
End Sub

Private Sub sagtusum_Click(index As Integer)
Select Case index
    Case 0
        prog.Selected = "*" & RightClickIndex & "*"
        DeleteObjects
        ClearSelected
    Case 1
        Status = "Move"
    Case 2
        SelectedOBJECT = RightClickIndex
        makeSelect SelectedOBJECT
    'case 3 Seperatör
    
    Case 4
        ozellikpenceresi RightClickIndex
End Select
End Sub

Private Sub Script_Error()
If Script.Error.Text = "" Then Exit Sub
MsgBox Dil(26) & Script.Error.Text & Dil(27)
Script.Error.Clear
End Sub

Private Sub Timer1_Timer()
Script.ExecuteStatement "timer1_timer"
End Sub

Private Sub Timer2_Timer()
ModbusTimeOut = ModbusTimeOut + 1
If ModbusTimeOut > Val(timout.Text) Then
ModbusWait = False
ModbusTimeOut = 0
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
prog.Counter = prog.Counter + 1
If prog.Counter = 5 Then prog.Counter = 1
Counteruygula prog.Counter
End Sub

Private Sub Winsock1_DataArrival(ByVal datalength As Long)
Dim b As Byte
Dim c, i
Dim j As Byte
For i = 1 To datalength
    Winsock1.GetData b
    MbusByteArray(i) = b
Next
j = 0
If MbusRead Then
        'Analog Outputlarý Alýyor...
         c = MbusByteArray(9) + 9
Select Case MbusByteArray(8)
    Case 3    ' Analog Output verileri geliyor
         For i = 10 To c Step 2
        'For i = 1 To datalength
        'Text1.Text = Str(j) + ": " + " [ " + Str((MbusByteArray(i) * 255) + MbusByteArray(i + 1)) + " ]"
        'Text1.Text = Str(j) + ": " + " [ " + Str(MbusByteArray(i)) + " ]"
        'List1.AddItem (Text1.Text)
        analog(j).Caption = Str(MbusByteArray(i) * 256) + MbusByteArray(i + 1)
        j = j + 1
        Next i
        ModbusWait = False
        ModbusTimeOut = 0
        Timer2.Enabled = False
    Case 1  ' Dijital Output verileri geliyor
        For i = 10 To c Step 2
        DOINT(j).Caption = Str(MbusByteArray(i + 1) * 256) + MbusByteArray(i)
        j = j + 1
        Next i
        ModbusWait = False
        ModbusTimeOut = 0
        Timer2.Enabled = False
End Select
End If

If MbusWrite Then
ModbusWait = False
ModbusTimeOut = 0
Timer2.Enabled = False
End If

End Sub

Private Sub yapistir_Click()
If kopyaaktif = True Then
    blank = BlankData()
    OBJECT(blank) = kopyamiz
    Status = "Move"
End If
End Sub

Private Sub yeni_Click()
Dim i
For i = 0 To 100
    With OBJECT(i)
        .Clicked = False
        .MainColor = 0
        .Height = 0
        .index = 0
        .Left = 0
        .Text = ""
        .OBJECT = ""
        .BorderColor = 0
        .Top = 0
        .Width = 0
    End With
Next i
ReBuildScreen Picture1
oascript = ""
LoadedFile = ""
RichBox1.TextRTF = "{\rtf1\ansi\ansicpg1254\deff0\deflang1055{\fonttbl{\f0\fnil\fcharset162{\*\fname Courier New;}Courier New TUR;}}" & vbCrLf & "\viewkind4\uc1\pard\f0\fs20" & vbCrLf & "\par }"
End Sub

' MODBUS FONKSÝYONLARI.....

Function baglan() As Boolean
Dim StartTime
If (Winsock1.State <> sckClosed) Then
    Winsock1.Close
End If
Winsock1.RemoteHost = IP.Text
Winsock1.Connect

StartTime = Timer

Do While ((Timer < StartTime + Val(timout.Text)) And (Winsock1.State <> 7))
DoEvents
Loop
If (Winsock1.State = 7) Then
   BagDurum = IP.Text & Dil(28)
   baglan = True
   Command2.Caption = Dil(31)
   BagDurum.BackColor = &HFF00&
Else
   BagDurum = IP.Text + Dil(29)
   baglan = False
   BagDurum.BackColor = &HFF
   Command2.Caption = Dil(30)
End If
End Function

Sub digitaloutverigonder()
Dim MbusWriteCommand As String
Dim StartLow As Byte
Dim StartHigh As Byte
Dim ByteLow As Byte
Dim ByteHigh As Byte
Dim LengthLow As Byte
Dim LengthHigh As Byte
Dim MbusWriteQuery
Dim i As Integer
If (Winsock1.State = 7) Then
'Start Low hangi outputtan baþlayacaðýný belirler.. 0 ise 0 dan
'en yüksek belirtilene kadar (Lenght tanýmlamasýyla)
'deðerleri belirler.
StartLow = 32 Mod 256  '32 den baþla
StartHigh = 32 \ 256
LengthLow = 64 Mod 256 ' Burada Kullanýan  64 ler 64 adet DO ' yu simgeler
LengthHigh = 64 \ 256  ' 64'e kadar devam et...

MbusWriteQuery = Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(7 + (2 * 32)) + Chr(1) + Chr(15) + Chr(StartHigh) + Chr(StartLow) + Chr(0) + Chr(32) + Chr(2 * 32)
For i = 2 To 3
ByteLow = Val(DOINT(i)) \ 256
ByteHigh = Val(DOINT(i)) Mod 256
MbusWriteQuery = MbusWriteQuery + Chr(ByteHigh) + Chr(ByteLow)
Next i
MbusRead = False
MbusWrite = True
Winsock1.SendData MbusWriteQuery
ModbusWait = True
ModbusTimeOut = 0
Timer2.Enabled = True
Else
MsgBox Dil(32)
Timer3.Enabled = False
tool.Buttons(14).Enabled = False
tool.Buttons(13).Enabled = True
Command2.Caption = Dil(30)
BagDurum = IP.Text + Dil(33)
BagDurum.BackColor = &HFF

End If
End Sub

Sub outputistek(veri As Long)
Dim StartLow As Byte
Dim StartHigh As Byte
Dim LengthLow As Byte
Dim LengthHigh As Byte
Dim loww
Dim highh
If veri = 1 Then
    loww = 32 ' 1 DO oluyor ve 32 adet DO verisi alýnacak
    highh = 32
Else
    loww = 16
    highh = 16
End If
If (Winsock1.State = 7) Then
' Analog Outputlarý Alýyor....
StartLow = 0 Mod 256
StartHigh = 0 \ 256
LengthLow = loww Mod 256
LengthHigh = highh \ 256
' Burada istek gönderiliyor...
MbusQuery(0) = 0
MbusQuery(1) = 0
MbusQuery(2) = 0
MbusQuery(3) = 0
MbusQuery(4) = 0
MbusQuery(5) = 6
MbusQuery(6) = 1
MbusQuery(7) = veri  ' 0= bilmem, 1= D.O. Lar 2= D.I.lar 3=A.Outputlar, 4=Analog Ýnputlar
MbusQuery(8) = StartHigh
MbusQuery(9) = StartLow
MbusQuery(10) = LengthHigh
MbusQuery(11) = LengthLow
MbusRead = True
analogout = True
DigitalOut = False
MbusWrite = False
'MbusQuery = Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(6) + Chr(1) + Chr(3) + Chr(StartHigh) + Chr(StartLow) + Chr(LengtHigh) + Chr(LengthLow)
Winsock1.SendData MbusQuery
ModbusWait = True
ModbusTimeOut = 0
Timer2.Enabled = True
Else
MsgBox Dil(32)
Timer3.Enabled = False
tool.Buttons(14).Enabled = False
tool.Buttons(13).Enabled = True
Command2.Caption = Dil(30)
BagDurum = IP.Text + Dil(33)
BagDurum.BackColor = &HFF
End If
End Sub

Sub Counteruygula(veri As Integer)
Select Case veri
    Case 1 ' Alýnacan DO larý iþleyecek
       outputistek 1
    Case 2 ' Yazýlacak DO larý iþleyecek
       digitaloutverigonder
    Case 3 'Alýnacak AO larý iþleyecek
       outputistek 3
    Case 4  'Yazýlacak AO larý iþleyecek
       analogoutverigonder
End Select
End Sub

Sub analogoutverigonder()
Dim MbusWriteCommand As String
Dim StartLow As Byte
Dim StartHigh As Byte
Dim LengthLow As Byte
Dim LengthHigh As Byte
Dim ByteLow As Byte
Dim ByteHigh As Byte
Dim MbusWriteQuery
Dim i As Integer
If (Winsock1.State = 7) Then
StartLow = 16 Mod 256
StartHigh = 16 \ 256
LengthLow = 32 Mod 256
LengthHigh = 32 \ 256

MbusWriteQuery = Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(7 + 2 * 32) + Chr(1) + Chr(16) + Chr(StartHigh) + Chr(StartLow) + Chr(0) + Chr(32) + Chr(2 * 32)
For i = 16 To 31
ByteLow = Val(analog(i).Caption) Mod 256
ByteHigh = Val(analog(i).Caption) \ 256
MbusWriteQuery = MbusWriteQuery + Chr(ByteHigh) + Chr(ByteLow)
Next i
MbusRead = False
MbusWrite = True
Winsock1.SendData MbusWriteQuery
ModbusWait = True
ModbusTimeOut = 0
Timer2.Enabled = True
Else
dur.Panels(3).Text = Dil(32)
Timer3.Enabled = False
tool.Buttons(14).Enabled = False
tool.Buttons(13).Enabled = True
Command2.Caption = Dil(31)
BagDurum = IP.Text + Dil(33)
BagDurum.BackColor = &HFF
End If
End Sub

Sub Selectedeekle(indexim)
If prog.Selected = "" Then prog.Selected = "*" & indexim & "*": Exit Sub
If InStr(prog.Selected, "*" & indexim & "*") <> 0 Then
    Exit Sub
Else
    prog.Selected = prog.Selected & indexim & "*"
End If
End Sub

Sub DeleteObjects()
Dim a, i, b
Dim seci
a = Mid$(prog.Selected, 1, Len(prog.Selected) - 1)
a = Mid$(a, 2, Len(prog.Selected))
seci = Split(a, "*")

        b = MsgBox(Dil(34), vbYesNo, Dil(69))
        If b = 6 Then
            For i = 0 To UBound(seci)
            With OBJECT(seci(i))
                .MainColor = 0
                .Height = 0
                .index = -1
                .Left = 0
                .OBJECT = ""
                .BorderColor = 0
                .Top = 0
                .Width = 0
            End With
            Next i
        End If
End Sub
Sub ClearSelected()
tool.Buttons(9).Value = tbrUnpressed
prog.Selected = ""
tool.Buttons(11).Enabled = False
tool.Buttons(12).Enabled = False
SelectedOBJECT = -1
ReBuildScreen Picture1
End Sub

Sub makeSelect(indexno)
        If OBJECT(indexno).OBJECT = "Circle" Then
            Picture1.Circle (OBJECT(indexno).Left, OBJECT(indexno).Top), 40, RGB(255, 0, 255)
            Picture1.Circle (OBJECT(indexno).Left, OBJECT(indexno).Top + OBJECT(indexno).Height), 40, RGB(255, 0, 255)
        Else
            Picture1.Circle (OBJECT(indexno).Left, OBJECT(indexno).Top), 40, RGB(255, 0, 255)
            Picture1.Circle (OBJECT(indexno).Left + OBJECT(indexno).Width, OBJECT(indexno).Top + OBJECT(indexno).Height), 40, RGB(255, 0, 255)
        End If
End Sub

Sub ozellikpenceresi(index)
        With OBJECT(index)
            form2.Text1 = .OBJECT
            form2.Text2 = .index
            form2.Text3 = .Left
            form2.Text4 = .Top
            form2.Text5 = .Width
            form2.Text6 = .Height
            form2.Text7 = .BorderColor
            form2.Text8 = .MainColor
            form2.Text9 = .Text
            form2.scrol.Value = .index
        End With
        form2.Show 1, Form1
End Sub

Sub taralialansec(X1, Y1, X2, Y2)
Dim a, i, indeximiz
For a = Y1 To Y2 Step 150
    For i = X1 To X2 Step 150
        indeximiz = FindObject(i, a)
        If indeximiz <> -1 Then
            makeSelect indeximiz
            Selectedeekle indeximiz
            SelectedOBJECT = indeximiz
            tool.Buttons(11).Enabled = True
            tool.Buttons(12).Enabled = True
        End If
    Next i
Next a
End Sub
