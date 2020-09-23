VERSION 5.00
Begin VB.UserControl OBJE 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   ScaleHeight     =   405
   ScaleWidth      =   405
End
Attribute VB_Name = "OBJE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
UserControl.Height = 400
UserControl.Width = 400
End Sub

Public Function GetOBJECTInfo(index, types) As String
Select Case types
    Case "Left"
        GetOBJECTInfo = OBJECT(Val(index)).Left
    Case "Top"
        GetOBJECTInfo = OBJECT(Val(index)).Top
    Case "Height"
        GetOBJECTInfo = OBJECT(Val(index)).Height
    Case "Width"
        GetOBJECTInfo = OBJECT(Val(index)).Width
    Case "BorderColor"
        GetOBJECTInfo = OBJECT(Val(index)).BorderColor
    Case "MainColor"
        GetOBJECTInfo = OBJECT(Val(index)).MainColor
    Case "Text"
        GetOBJECTInfo = OBJECT(Val(index)).Text
End Select
End Function

Public Sub SetOBJECTInfo(index, types, info)
Select Case types
    Case "Left"
        OBJECT(Val(index)).Left = info
        ReBuildScreen Form1.Picture1
    Case "Top"
        OBJECT(Val(index)).Top = info
        ReBuildScreen Form1.Picture1
    Case "Height"
        OBJECT(Val(index)).Height = info
        ReBuildScreen Form1.Picture1
    Case "Width"
        OBJECT(Val(index)).Width = info
        ReBuildScreen Form1.Picture1
    Case "BorderColor"
        OBJECT(Val(index)).BorderColor = info
        ReBuildScreen Form1.Picture1
    Case "MainColor"
        OBJECT(Val(index)).MainColor = info
        ReBuildScreen Form1.Picture1
    Case "Text"
        OBJECT(Val(index)).Text = info
        ReBuildScreen Form1.Picture1
End Select

End Sub

Sub toPLC(DataType, index, Value)
Select Case DataType
    Case "DDO"
        Form1.DOO(Val(index) - 1).Caption = Val(Value)
    Case "ACO"
        Form1.analog(Val(index) - 1).Caption = Val(Value)
End Select
End Sub
Public Function fromPLC(DataType, index) As String
Select Case DataType
    Case "DDO"
        fromPLC = Form1.DOO(Val(index) - 1).Caption
    Case "ACO"
        fromPLC = Form1.analog(Val(index) - 1).Caption
End Select
End Function

Sub myTimer(Value)
Form1.Timer1.Enabled = Value
End Sub
