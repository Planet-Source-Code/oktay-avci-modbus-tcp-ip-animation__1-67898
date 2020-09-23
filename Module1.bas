Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Type types
    OBJECT As String
    Index As Long
    Left As Long
    Top As Long
    Width As Long
    Height As Long
    BorderColor As Long
    MainColor As Long
    Text As String
    Clicked As Boolean
End Type

Public Type COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type

Type secsec
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Type program
    Counter As Integer
    DigitalOut As Boolean
    analogout As Boolean
    Saved As String
    Selected As String
    dilimiz As String
End Type

Public Const WM_USER = &H400
Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72

Global pint As POINTAPI
Global prog As program
Global LoadedFile As String
Global oascript As String
Global OBJECT(100) As types
Global kopyamiz As types
Global SelectOBJECTct As secsec
Global Dil(150) As String
Global dilimiz As String

Public Function INIYukle(Baslik As String, Anahtar As String, dosya As String) As String
Dim strBuffer As String
 strBuffer = String(750, Chr(0))
 Anahtar$ = LCase$(Anahtar$)
 INIYukle$ = Left(strBuffer, GetPrivateProfileString(Baslik$, ByVal Anahtar$, "", strBuffer, Len(strBuffer), dosya$))
End Function
Public Sub INIYaz(Baslik As String, Anahtar As String, Deger As String, dosya As String)
 Call WritePrivateProfileString(Baslik$, UCase$(Anahtar$), Deger$, dosya$)
 DoEvents
End Sub
Sub ReBuildScreen(pic As PictureBox)
Dim i
pic.Cls
For i = 0 To 100
    Select Case OBJECT(i).OBJECT
        Case "Line"
            pic.Line (OBJECT(i).Left, OBJECT(i).Top)-(OBJECT(i).Left + OBJECT(i).Width, OBJECT(i).Top + OBJECT(i).Height), OBJECT(i).MainColor
        Case "Rectangle"
            pic.FillColor = OBJECT(i).MainColor
            pic.FillStyle = 0
            pic.Line (OBJECT(i).Left, OBJECT(i).Top)-(OBJECT(i).Left + OBJECT(i).Width, OBJECT(i).Top + OBJECT(i).Height), OBJECT(i).BorderColor, B
            pic.FillStyle = 1
        Case "Circle"
            pic.FillColor = OBJECT(i).MainColor
            pic.FillStyle = 0
            On Error GoTo 10
            pic.Circle (OBJECT(i).Left, OBJECT(i).Top), OBJECT(i).Height, OBJECT(i).BorderColor
            pic.FillStyle = 1
        Case "Button"
            CreateButton pic, OBJECT(i).Left, OBJECT(i).Top, OBJECT(i).Width, OBJECT(i).Height, OBJECT(i).Text, OBJECT(i).Clicked, OBJECT(i).MainColor
        Case "Label"
            CreateLabel pic, OBJECT(i).Left, OBJECT(i).Top, OBJECT(i).Width, OBJECT(i).Height, OBJECT(i).Text, OBJECT(i).MainColor, OBJECT(i).BorderColor, OBJECT(i).Clicked
        Case "Text"
            CreateText pic, OBJECT(i).Left, OBJECT(i).Top, OBJECT(i).Width, OBJECT(i).Height, OBJECT(i).Text

    End Select
Next i
10
End Sub

Public Function FindObject(x, y) As Long
Dim i, a, b, c, d, e, f, diff
For i = 100 To 0 Step -1
If OBJECT(i).OBJECT = "" Then GoTo 13
If OBJECT(i).OBJECT = "Circle" Then
    a = OBJECT(i).Left - OBJECT(i).Width
    b = OBJECT(i).Top - OBJECT(i).Height
    c = OBJECT(i).Left + OBJECT(i).Width
    d = OBJECT(i).Top + OBJECT(i).Height
Else
    a = OBJECT(i).Left
    b = OBJECT(i).Top
    c = OBJECT(i).Left + OBJECT(i).Width
    d = OBJECT(i).Top + OBJECT(i).Height
End If
If x > a - 25 And x < c + 25 And y > b - 25 And y < d + 25 Then
    FindObject = i
    Exit Function
End If
13 Next i
FindObject = -1 ' bulunamadý hatasý nosu
End Function

Sub CreateButton(pic As PictureBox, Left, Top, Width, Height, Text, Clicked As Boolean, MainColor)
Dim a, b, c
b = 0
c = 0
a = pic.DrawWidth
pic.DrawWidth = 1
pic.Line (Left, Top)-(Left + Width, Top + Height), 0, B
If Clicked = True Then
    pic.Line (Left + 20, Top + 20)-((Left + Width) - 20, Top + 20), RGB(100, 100, 100)
    pic.Line (Left + 30, Top + 30)-((Left + Width) - 30, Top + 30), RGB(100, 100, 100)
    pic.Line (Left + 20, Top + 20)-(Left + 20, (Top + Height) - 20), RGB(100, 100, 100)
    pic.Line (Left + 30, Top + 30)-(Left + 30, (Top + Height) - 30), RGB(100, 100, 100)

    pic.Line (Left + 20, (Top + Height) - 20)-((Left + Width) - 20, (Top + Height) - 20), RGB(220, 220, 220)
    pic.Line (Left + 30, (Top + Height) - 30)-((Left + Width) - 30, (Top + Height) - 30), RGB(220, 220, 220)
    pic.Line ((Left + Width) - 20, Top + 20)-((Left + Width) - 20, (Top + Height) - 20), RGB(220, 220, 220)
    pic.Line ((Left + Width) - 30, Top + 30)-((Left + Width) - 30, (Top + Height) - 30), RGB(220, 220, 220)
Else
    pic.Line (Left + 20, Top + 20)-((Left + Width) - 20, Top + 20), RGB(220, 220, 220)
    pic.Line (Left + 30, Top + 30)-((Left + Width) - 30, Top + 30), RGB(220, 220, 220)
    pic.Line (Left + 20, Top + 20)-(Left + 20, (Top + Height) - 20), RGB(220, 220, 220)
    pic.Line (Left + 30, Top + 30)-(Left + 30, (Top + Height) - 30), RGB(220, 220, 220)

    pic.Line (Left + 20, (Top + Height) - 20)-((Left + Width) - 20, (Top + Height) - 20), RGB(100, 100, 100)
    pic.Line (Left + 30, (Top + Height) - 30)-((Left + Width) - 30, (Top + Height) - 30), RGB(100, 100, 100)
    pic.Line ((Left + Width) - 20, Top + 20)-((Left + Width) - 20, (Top + Height) - 20), RGB(100, 100, 100)
    pic.Line ((Left + Width) - 30, Top + 30)-((Left + Width) - 30, (Top + Height) - 30), RGB(100, 100, 100)
End If
pic.Line (Left + 40, Top + 40)-((Left + Width) - 40, (Top + Height) - 40), MainColor, BF
pic.FontBold = True
If Clicked = True Then b = 20
pic.CurrentX = b + Left + (Width - pic.TextWidth(Text)) / 2
pic.CurrentY = b + Top + (Height - pic.TextHeight(Text)) / 2
If Width > pic.TextWidth(Text) And Height > pic.TextHeight(Text) Then pic.Print Text
pic.DrawWidth = a
pic.FontBold = False
End Sub

Sub CreateLabel(pic As PictureBox, Left, Top, Width, Height, Text, arkarenk, kenarrenk, borderim)
Dim a
a = pic.DrawWidth
pic.DrawWidth = 1
If borderim = True Then
pic.Line (Left, Top)-(Left + Width, Top + Height), arkarenk, BF
pic.Line (Left, Top)-(Left + Width, Top + Height), kenarrenk, B
End If
pic.CurrentX = Left + (Width - pic.TextWidth(Text)) / 2
pic.CurrentY = Top + (Height - pic.TextHeight(Text)) / 2
If Width > pic.TextWidth(Text) And Height > pic.TextHeight(Text) Then pic.Print Text
pic.DrawWidth = a
End Sub

Sub CreateText(pic As PictureBox, Left, Top, Width, Height, Text)
Dim a
a = pic.DrawWidth
pic.DrawWidth = 1
pic.Line (Left, Top)-(Left + Width, Top + Height), &HFFFFFF, BF
pic.Line (Left, Top)-(Left + Width, Top + Height), 0, B
pic.CurrentX = Left + (Width - pic.TextWidth(Text)) / 2
pic.CurrentY = Top + (Height - pic.TextHeight(Text)) / 2
If Width > pic.TextWidth(Text) And Height > pic.TextHeight(Text) Then pic.Print Text
pic.DrawWidth = a
End Sub

'Functions TCP/IP....
Public Sub WriteValue(Index, Value, types)
'Dialogtan deðer girmek için
Select Case types
    Case "digital"
        Form1.DOO(Index) = Value
    Case "analog"
        Form1.analog(Index) = Value
End Select
End Sub

Public Sub setpointform(types, Value, Index)
Select Case types
    Case "ana"
        setpnt.Text1.Left = 120
        setpnt.Text1 = Value
        setpnt.Text1.Visible = True
        setpnt.Check1.Visible = False
        setpnt.types = types
        setpnt.Index = Index
        setpnt.Show 1, Form1
    Case "dig"
        setpnt.Check1.Left = 120
        setpnt.Check1.Visible = True
        setpnt.Check1.Value = Value
        setpnt.Text1.Visible = False
        setpnt.types = types
        setpnt.Index = Index
        setpnt.Show 1, Form1
End Select
End Sub

Public Sub MenageRTF(KeyAscii, rtf As RichTextBox)
If KeyAscii = 20 Or KeyAscii = 8 Then
    rtf.SelColor = 0
End If
End Sub

Public Function ObjectPoint(Selected, x, y, ByRef myObject, ByRef Where) As Boolean
'Where : First = Top - left yapýsý , Last = height - width yapýsý
'myObject  : DA = Dairedir, DI = digerleri

Dim a, b, c, d, e

ObjectPoint = False
If OBJECT(Selected).OBJECT = "Circle" Then
    a = OBJECT(Selected).Height + OBJECT(Selected).Top
    If y > a - 25 And y < a + 25 Then
        ObjectPoint = True
        myObject = "DA"
        Where = "Last"
    End If
Else
    a = OBJECT(Selected).Top
    b = OBJECT(Selected).Left
    If y > a - 25 And y < a + 25 And x > b - 25 And x < b + 25 Then
        ObjectPoint = True
        myObject = "DI"
        Where = "First"
        Exit Function
    End If
    a = OBJECT(Selected).Top + OBJECT(Selected).Height
    b = OBJECT(Selected).Left + OBJECT(Selected).Width
    If y > a - 25 And y < a + 25 And x > b - 25 And x < b + 25 Then
        ObjectPoint = True
        myObject = "DI"
        Where = "Last"
    End If
End If
End Function

Public Sub ReScale(Index, x, y, myObject, Where)
Dim a, b
Select Case myObject
    Case "DA"
        OBJECT(Index).Width = y - OBJECT(Index).Top
        OBJECT(Index).Height = y - OBJECT(Index).Top
        If OBJECT(Index).Width < 50 Then
            OBJECT(Index).Width = 50
            OBJECT(Index).Height = 50
        End If
    Case "DI"
        Select Case Where
            Case "First"
                a = OBJECT(Index).Top + OBJECT(Index).Height
                b = OBJECT(Index).Left + OBJECT(Index).Width
                OBJECT(Index).Top = y
                OBJECT(Index).Left = x
                OBJECT(Index).Width = b - x
                OBJECT(Index).Height = a - y

            Case "Last"
                OBJECT(Index).Width = x - OBJECT(Index).Left
                OBJECT(Index).Height = y - OBJECT(Index).Top
        End Select
End Select
End Sub
Private Sub GradientClr(ByRef PicBox As PictureBox, ByVal c1 As Long, ByVal c2 As Long)
'paints a gradient to the referred picture box
'mainly used for displaying the color selections
Dim r(2) As Single, g(2) As Single, b(2) As Single
Dim i As Integer, ix As Integer

 r(0) = GetRGB(c1).Red
 g(0) = GetRGB(c1).Green
 b(0) = GetRGB(c1).Blue

 r(1) = GetRGB(c2).Red
 g(1) = GetRGB(c2).Green
 b(1) = GetRGB(c2).Blue

i = PicBox.ScaleWidth
If i > 255 Then i = 255

 r(2) = (r(1) - r(0)) / i
 g(2) = (g(1) - g(0)) / i
 b(2) = (b(1) - b(0)) / i

For ix = 0 To PicBox.ScaleWidth

 If r(0) < 0 Then r(0) = 0
 If r(0) > 255 Then r(0) = 255
 If g(0) < 0 Then g(0) = 0
 If g(0) > 255 Then g(0) = 255
 If b(0) < 0 Then b(0) = 0
 If b(0) > 255 Then b(0) = 255

 PicBox.Line (ix, 0)-(ix, PicBox.ScaleHeight), RGB(r(0), g(0), b(0)), BF
 r(0) = r(0) + r(2)
 g(0) = g(0) + g(2)
 b(0) = b(0) + b(2)
 
Next ix
End Sub

Public Function GetRGB(ByVal CVal As Long) As COLORRGB
'returns rgb values
  GetRGB.Blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
  GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function

Public Sub loadlanguage(filename)
Dim i
Open filename For Input As #1
For i = 0 To 150
    If EOF(1) = True Then GoTo 10
    Line Input #1, Dil(i)
Next i
10
Close #1

'FORM1 Yükleniyor.....
With Form1
    .dosya.Caption = Dil(53)
    .yeni.Caption = Dil(54)
    .dosyac.Caption = Dil(55)
    .kaydet.Caption = Dil(56)
    .farklikaydet.Caption = Dil(57)
    .resimyap.Caption = Dil(58)
    .cikis.Caption = Dil(59)
    .duzen.Caption = Dil(60)
    .kes.Caption = Dil(61)
    .kopyala.Caption = Dil(62)
    .yapistir.Caption = Dil(63)
    .kopyam(1).Caption = Dil(61)
    .kopyam(2).Caption = Dil(62)
    .kopyam(3).Caption = Dil(63)
    .kodedit.Caption = Dil(64)
    .ayarlar.Caption = Dil(65)
    .progsettings.Caption = Dil(66)
    .yardim.Caption = Dil(67)
    .hakkinda.Caption = Dil(68)
    .sagtusum(0).Caption = Dil(10)
    .sagtusum(1).Caption = Dil(11)
    .sagtusum(2).Caption = Dil(12)
    .sagtusum(4).Caption = Dil(13)
    .Caption = Dil(69) & " [ " & Dil(70) & " ]"
    .tool.Buttons(1).ToolTipText = Dil(1)
    .tool.Buttons(2).ToolTipText = Dil(2)
    .tool.Buttons(3).ToolTipText = Dil(3)
    .tool.Buttons(4).ToolTipText = Dil(4)
    .tool.Buttons(5).ToolTipText = Dil(5)
    .tool.Buttons(6).ToolTipText = Dil(6)
    .tool.Buttons(8).ToolTipText = Dil(7)
    .tool.Buttons(9).ToolTipText = Dil(8)
    .tool.Buttons(10).ToolTipText = Dil(9)
    .tool.Buttons(11).ToolTipText = Dil(10)
    .tool.Buttons(12).ToolTipText = Dil(13)
    .tool.Buttons(14).ToolTipText = Dil(14)
    .tool.Buttons(15).ToolTipText = Dil(15)
    .Label16.Caption = Dil(50)
    .Label40.Caption = Dil(72)
    .Frame1.Caption = Dil(73)
    .Frame2.Caption = Dil(74)
    .Frame3.Caption = Dil(75)
    .Frame4.Caption = Dil(76)
    .Frame5.Caption = Dil(77)
    .Label19.Caption = Dil(78)
    .Label18.Caption = Dil(79)
    .Label17.Caption = Dil(80)
    .Command2.Caption = Dil(20)
    .dur.Panels(3).Text = Dil(0) & " " & Dil(16)
End With

With about
    .Label1.Caption = Dil(69)
    .Label2.Caption = Dil(70)
    .Label4.Caption = Dil(83)
    .Label3.Caption = Dil(82)
End With

With form2
.Caption = Dil(13)
.Label1.Caption = Dil(43)
.Label2.Caption = Dil(44)
.Label3.Caption = Dil(45)
.Label4.Caption = Dil(46)
.Label5.Caption = Dil(47)
.Label6.Caption = Dil(48)
.Label7.Caption = Dil(49)
.Label10.Caption = Dil(50)
.Label8.Caption = Dil(51)
.Check1.Caption = Dil(52)
.Command5.Caption = Dil(41)
.Command1.Caption = Dil(38)
.Command2.Caption = Dil(39)
.Command6.Caption = Dil(40)
End With

With kod
    .Caption = Dil(42)
    .Command1.Caption = Dil(38)
    .Command2.Caption = Dil(39)
End With

With setpnt
    .Caption = Dil(36)
    .Check1.Caption = Dil(37)
    .OKButton.Caption = Dil(38)
    .CancelButton.Caption = Dil(39)
End With

With settings
    .Caption = Dil(66)
    .Label1.Caption = Dil(85)
    .OKButton.Caption = Dil(38)
    .CancelButton.Caption = Dil(39)
End With
App.Title = Dil(69)
End Sub
