VERSION 5.00
Begin VB.Form titlebar 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   360
   ClientLeft      =   4470
   ClientTop       =   3840
   ClientWidth     =   8835
   Icon            =   "titlebar.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "titlebar.frx":1272
   ScaleHeight     =   360
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picmainskin 
      Height          =   255
      Left            =   2880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   8280
      Picture         =   "titlebar.frx":B934
      Top             =   60
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   7950
      Picture         =   "titlebar.frx":BCEA
      Top             =   60
      Width           =   255
   End
   Begin VB.Image titleimage 
      Height          =   330
      Left            =   360
      Picture         =   "titlebar.frx":C0A0
      Top             =   30
      Width           =   3045
   End
End
Attribute VB_Name = "titlebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'>>>>>>>>> PROGRAMIN TRAY-ICON MODU ÝÇÝN GEREKLÝ API VE CONSTANTLAR
Private Type NOTIFYICONDATA
cbSize As Long
hWnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim nid As NOTIFYICONDATA
'<<<<<<<<<




'>>>>> FORM SÜRÜKLEME OLAYI
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
'<<<<<<<


'>>>>> FORM HEP ÜSTTE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'<<<<


Sub sekillendir()
'>>>> BU KOD ANA SAYFAYI TITTLE BAR A OTURTUYOR.
         mainpage.Top = titlebar.Top + titlebar.Height
         mainpage.Left = titlebar.Left
         mainpage.Width = titlebar.Width - 80

         
'<<<<<<<<
End Sub




Private Sub Form_Activate()
Call sekillendir
End Sub


Private Sub Form_Click()

Call mainpage.menulerikapat

End Sub

Private Sub Form_DblClick()
Call mainpage.menulerikapat
End Sub

Private Sub Form_Load()
Call sekillendir
         
'>>>>>>>>> BU KOD FORM`UN HEP USTTE DURMASINI SAÐLIYOR
        SetWindowPos mainpage.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'<<<<<<<<<


'>>>>>>>>> BU KOD TITTLE BAR IN HEP USTTE DURMASINI SAÐLIYOR
         SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'<<<<<<<<<




'>>>>>>>>>ÞEFFAF FORM KODU
    picmainskin.ScaleMode = vbPixels
    picmainskin.AutoRedraw = True
    picmainskin.AutoSize = True
    picmainskin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
    Set picmainskin.Picture = LoadPicture(App.Path & "\bar.bmp")
    Me.Width = picmainskin.Width
    Me.Height = picmainskin.Height
    WindowRegion = MakeRegion(picmainskin)
    SetWindowRgn Me.hWnd, WindowRegion, True
'<<<<<<<<<<<<<<<<<


End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim msg As Long
msg = X / Screen.TwipsPerPixelX
Select Case msg

Case WM_LBUTTONUP
titlebar.Visible = True
mainpage.Visible = True
Shell_NotifyIcon NIM_DELETE, nid
End Select



Dim lngReturnValue As Long
    If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

Call sekillendir
  Call mainpage.menulerikapat
    End If
End Sub



Private Sub Image1_Click()
'>>>>>>> PROGRAMIN BÝR SÜRE ÝÇÝN GÖRÜNMEZ OLMASINI SAÐLAYAN KODLAR
mainpage.Visible = False
titlebar.Visible = False
'<<<<<<<

Call mainpage.menulerikapat

mainpage.Visible = False
titlebar.Visible = False

'>>>>>>> PROGRAMI TRAY-ICON A ÝNDÝREN KODLAR

nid.cbSize = Len(nid)
nid.hWnd = Me.hWnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Me.Icon
nid.szTip = "Yazýlýmý görüntülemek için týklayýnýz!!!" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid

''<<<<<<<
End Sub




Private Sub Image2_Click()
Call mainpage.menulerikapat

'>>>>>>> PROGRAMIN BÝR SÜRE ÝÇÝN GÖRÜNMEZ OLMASINI SAÐLAYAN KODLAR
mainpage.Visible = False
titlebar.Visible = False
'<<<<<<<

X = MsgBox("Bekçi Kontrol Sistemi Yazýlýmýný kapatmak istediðinize emin misiniz?", vbYesNo + vbInformation, "Programý Kapat")
If X = vbYes Then
End
Else
mainpage.Visible = True
titlebar.Visible = True

End If
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Top = Image2.Top - 15
End Sub

Private Sub Image1_DblClick()
Call mainpage.menulerikapat
Image1.Top = Image1.Top + 15

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image1.Top = Image1.Top + 15

End Sub


Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Top = Image1.Top - 15

End Sub


Private Sub Image2_DblClick()
Call mainpage.menulerikapat
Image2.Top = Image2.Top + 15

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image2.Top = Image2.Top + 15

End Sub



Private Sub Timer1_Timer()
Call sekillendir
End Sub

Private Sub titleimage_Click()

Call mainpage.menulerikapat
End Sub

Private Sub titleimage_DblClick()
Call mainpage.menulerikapat
End Sub

Private Sub titleimage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngReturnValue As Long
    If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    Call mainpage.menulerikapat
        Call sekillendir
    End If


End Sub
