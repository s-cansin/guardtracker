VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form mainpage 
   BorderStyle     =   0  'None
   ClientHeight    =   7815
   ClientLeft      =   4470
   ClientTop       =   3000
   ClientWidth     =   8805
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H8000000A&
   Icon            =   "mainpage.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ayarlar 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2090
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Ýletiþim Portu"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Gecikme Yönetcisi"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Tablo Rengi"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H80000016&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         FillColor       =   &H00E0E0E0&
         Height          =   1215
         Left            =   0
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox cihaz 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   1490
      ScaleHeight     =   1575
      ScaleWidth      =   1695
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Cihaz Bilgileri"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Saat Ýþlemleri"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cihazý Temizle"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Veri Transferi"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000016&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         FillColor       =   &H00E0E0E0&
         Height          =   1575
         Left            =   0
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox kayitlar 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   760
      ScaleHeight     =   855
      ScaleWidth      =   1935
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tüm Kayýtlarý Sil"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Tüm Kayýtlarý Göster"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000016&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         FillColor       =   &H00E0E0E0&
         Height          =   855
         Left            =   0
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox gorevler 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   40
      ScaleHeight     =   1215
      ScaleWidth      =   1695
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Görevleri Yazdýr"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Görev Sil"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Yeni Görev Ekle"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000016&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000003&
         FillColor       =   &H00E0E0E0&
         Height          =   1215
         Index           =   1
         Left            =   0
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.PictureBox picmainskin 
      Height          =   255
      Left            =   4440
      ScaleHeight     =   255
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   10000
      Width           =   15
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   13120
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7575
      Left            =   0
      TabIndex        =   23
      Top             =   240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   13361
      _Version        =   393216
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      X1              =   8790
      X2              =   8790
      Y1              =   0
      Y2              =   7800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   8800
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   7800
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kayýtlar"
      Height          =   255
      Left            =   885
      TabIndex        =   14
      Top             =   20
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cihaz"
      Height          =   255
      Left            =   1585
      TabIndex        =   13
      Top             =   20
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ayarlar"
      Height          =   255
      Left            =   2140
      TabIndex        =   12
      Top             =   20
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Hakkýnda"
      Height          =   255
      Left            =   2805
      TabIndex        =   11
      Top             =   20
      Width           =   735
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000003&
      Height          =   255
      Left            =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Görevler"
      Height          =   255
      Left            =   110
      TabIndex        =   10
      Top             =   20
      Width           =   615
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000003&
      Height          =   255
      Left            =   760
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000003&
      Height          =   255
      Left            =   1480
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H80000003&
      Height          =   255
      Left            =   2080
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H80000003&
      Height          =   255
      Left            =   2680
      Top             =   0
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "mainpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'>>>>> FORM HEP ÜSTTE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'<<<<<<<<<<<<<<<<<<<<


Sub menulerikapat()
gorevler.Visible = False
kayitlar.Visible = False
cihaz.Visible = False
ayarlar.Visible = False
End Sub




Private Sub Form_Terminate()
MSComm1.PortOpen = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
MSComm1.PortOpen = False
End Sub

Private Sub Label10_Click()

Call menulerikapat
titlebar.Visible = False
mainpage.Visible = False
deleteplace.Show
End Sub

Private Sub Label11_Click()
Call menulerikapat

MSFlexGrid1.Visible = True
MSFlexGrid1.Rows = 1


Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from logs order by id desc", Conn, adOpenKeyset, adLockOptimistic

If Not Rs.EOF Or Not BOF Then




X = ""

For i = 1 To Rs.RecordCount
  
MSFlexGrid1.AddItem (vbNullString)

If Not Len(Rs!id) = 0 Then
MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 0) = MSFlexGrid1.Rows - 1
End If


If Not Len(Rs!yer) = 0 Then

Dim Rd As New ADODB.Recordset
Rd.Open "Select * from dp where IButton='" & Rs!yer & "'", Conn, adOpenKeyset, adLockOptimistic



If Rd.EOF Or Rd.BOF Then
MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = "Bilinmeyen yer"
Else
MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 1) = Rd!yer

End If
Rd.Close
End If




If Not Len(Rs!tarih) = 0 Then
MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 2) = Rs!tarih
End If

If Not Len(Rs!saat) = 0 Then
MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 3) = Rs!saat

dakikax = Mid(Rs!saat, 4, 2) ' iþlemin yapýldýðý dakika


Dim Rw As New ADODB.Recordset
Rw.Open "Select * from sets", Conn, adOpenKeyset, adLockOptimistic

Dim Rq As New ADODB.Recordset
Rq.Open "Select * from dp where IButton='" & Rs("yer") & "' ", Conn, adOpenKeyset, adLockOptimistic



If Rw("gecikme") < dakikax Then
MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 5) = "X"
End If


Rq.Close
Rw.Close







End If

MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = dakikax

If Not Len(Rs!bekci) = 0 Then
Dim Rx As New ADODB.Recordset
Rx.Open "Select * from guard where kod='" & Rs("bekci") & "' ", Conn, adOpenKeyset, adLockOptimistic
If Rx.EOF Or Rx.BOF Then
MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = "Bilinmeyen Bekçi"
Else
MSFlexGrid1.TextMatrix(MSFlexGrid1.Rows - 1, 4) = Rx!ad
End If
Rx.Close
End If



Rs.MoveNext
Next

End If











End Sub

Private Sub Label12_Click()
Call menulerikapat
titlebar.Visible = False
mainpage.Visible = False
devicetime.Show
End Sub

Private Sub Label13_Click()
Call menulerikapat

mainpage.Visible = False
titlebar.Visible = False

lateadmin.Show
End Sub

Private Sub Label14_Click()
Call menulerikapat
titlebar.Visible = False
mainpage.Visible = False
deviceinfo.Show
End Sub

Private Sub Label15_Click()
Call menulerikapat

mainpage.Visible = False
titlebar.Visible = False

tablecolors.Show
End Sub

Private Sub Label16_Click()
Call menulerikapat

titlebar.Visible = False
mainpage.Visible = False

printerpage.Show
End Sub

Private Sub Label17_Click()
Call menulerikapat

mainpage.Visible = False
titlebar.Visible = False
comport.Show
End Sub

Private Sub Label7_Click()
Call menulerikapat

mainpage.Visible = False
titlebar.Visible = False
X = MsgBox("Dikkat! Cihazdaki tüm bilgileri silmek üzeresiniz!!! " & vbCrLf & "Geri dönüþü olmayan bu iþlemi yaparak cihazdaki bilgileri silmek istediðinize emin misiniz?", vbExclamation + vbYesNo, "Cihaz Temizleme")

If X = vbYes Then
wait.Show

MSComm1.Output = "*T"







DoEvents
'>>>>>PIC TEN GELEN BÝLGÝYÝ BEKLETME MODU
For i = 1 To 1000000
For z = 1 To 10
Next z
Next i
'<<<<<<<

Dim buffer
buffert = mainpage.MSComm1.Input


Unload wait
If buffert = "SECTOR CLEAR" Then
Y = MsgBox("Cihazdaki tüm bilgiler silinmiþtir!!!")
Else
Y = MsgBox("Dikkat!!! Cihazdaki bilgiler silinemedi..." & vbCrLf & "Lütfen cihazýn baðlantýsýný kontrol edip tekrar deneyiniz.", vbCritical + vbOKOnly, "Hata!")
End If




End If


titlebar.Visible = True
mainpage.Visible = True
End Sub

Private Sub Label8_Click()
Call menulerikapat

titlebar.Visible = False
mainpage.Visible = False




mesaj = MsgBox("Dikkat!" & vbCrLf & "Cihazdaki tüm veriler veri tabanýna kopyalanacaktýr!!! " & vbCrLf & "Lütfen cihaz ile olan iletiþimi kesmeyin!", vbOKOnly + vbInformation, "Veri Transferi")
wait.Show


Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from logs", Conn, adOpenKeyset, adLockOptimistic


MSComm1.Output = "*O"





DoEvents
'>>>>>PIC TEN GELEN BÝLGÝYÝ BEKLETME MODU

For i = 1 To 1000000
For z = 1 To 10

Next z
Next i
'<<<<<<<


buffer = MSComm1.Input



stat2 = 1
stat2 = 1

If Len(buffer) > 0 Then

'Buraya msgbox ilave edilerek aygýttan gelen toplu bilgi test edilebilir!!!
For i = 2 To Len(buffer)

If Mid(buffer, i, 1) = "#" Then

stat = stat2
stat2 = i
Text = Mid(buffer, stat, stat2 - stat)


Select Case Left(Text, 2)

Case "#K"

Rs.AddNew

Rs!yer = Right(Text, Len(Text) - 2)

Case "#S"

Rs!saat = Right(Text, Len(Text) - 2)

Case "#N"

Rs!bekci = Right(Text, Len(Text) - 2)

Case "#T"

Rs!tarih = Left(Right(Text, Len(Text) - 2), Len(Right(Text, Len(Text) - 2)) - 2)

Case "#E"

Rs.Update

End Select

End If

Next i

message = MsgBox("Cihazdaki veriler baþarýyla veritabanýna eklenmiþtir!!!")

titlebar.Visible = True
mainpage.Visible = True
Else

mesaj = MsgBox("Cihazdan bilgi alýnamadý!!!", vbOKOnly + vbExclamation, "Hata")


End If


Unload wait
titlebar.Visible = True
mainpage.Visible = True

Call Label11_Click
End Sub

Private Sub Label9_Click()
Call menulerikapat

Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from logs", Conn, adOpenKeyset, adLockOptimistic

If Not Rs.EOF Or Not Rs.BOF Then

mainpage.Visible = False
titlebar.Visible = False

mesaj = MsgBox("Dikkat!!!" & vbCrLf & "Geçmiþteki tüm bekçi faaliyetleri ve kayýtlarý silinecektir..." & vbCrLf & "Gerçekten tüm kayýtlarý silmek istiyor musunuz?", vbOKCancel + vbExclamation, "Dikkat!")

If mesaj = vbOK Then
For i = 1 To Rs.RecordCount
Rs.Delete
Rs.MoveNext
Next i
End If


mainpage.Visible = True
titlebar.Visible = True
Call Label11_Click
End If
End Sub



Private Sub Form_Activate()
'>>>>>>>>> BU KOD FORM`UN HEP USTTE DURMASINI SAÐLIYOR
SetWindowPos titlebar.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'>>>>>>>>>

If Not Check1.Value = Checked Then
titlebar.Show
Check1.Value = Checked
Else

End If
Call Label11_Click
End Sub


Private Sub Form_Click()
Call menulerikapat
End Sub

Private Sub Form_DblClick()
Call menulerikapat
End Sub


Private Sub Form_Load()

Dim Connw As New ADODB.Connection
Connw.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Connw.Open

Dim Rsw As New ADODB.Recordset
Rsw.Open "Select * from sets", Connw, adOpenKeyset, adLockOptimistic





MSComm1.CommPort = Rsw("port")
MSComm1.PortOpen = True



Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from sets", Conn, adOpenKeyset, adLockOptimistic


With MSFlexGrid1
.ForeColor = Rs!baslik
.BackColor = Rs!satir
.BackColorFixed = Rs!yazi

.Cols = 6
ColWidth = (MSFlexGrid1.Width / (.Cols + 1)) + 155



.TextMatrix(0, 0) = "No"
.ColAlignment(0) = 4
.ColWidth(0) = ColWidth

.TextMatrix(0, 1) = "Yer"
.ColAlignment(1) = 4
.ColWidth(1) = ColWidth

.TextMatrix(0, 2) = "Tarih"
.ColAlignment(2) = 4
.ColWidth(2) = ColWidth

.TextMatrix(0, 3) = "Saat"
.ColAlignment(3) = 4
.ColWidth(3) = ColWidth

.TextMatrix(0, 4) = "Bekçi"
.ColAlignment(4) = 4
.ColWidth(4) = ColWidth

.TextMatrix(0, 5) = "Gecikme"
.ColAlignment(5) = 5
.ColWidth(5) = ColWidth


End With

Call Label11_Click



End Sub

Private Sub Label1_Click()
If gorevler.Visible = True Then
gorevler.Visible = False
Else
gorevler.Visible = True
kayitlar.Visible = False
cihaz.Visible = False
ayarlar.Visible = False
End If
End Sub


Private Sub Label2_Click()
If kayitlar.Visible = True Then
kayitlar.Visible = False
Else
gorevler.Visible = False
kayitlar.Visible = True
cihaz.Visible = False
ayarlar.Visible = False
End If
End Sub

Private Sub Label3_Click()
If cihaz.Visible = True Then
cihaz.Visible = False
Else
gorevler.Visible = False
kayitlar.Visible = False
cihaz.Visible = True
ayarlar.Visible = False
End If
End Sub





Private Sub Label4_Click()
If ayarlar.Visible = True Then
ayarlar.Visible = False
Else
gorevler.Visible = False
kayitlar.Visible = False
cihaz.Visible = False
ayarlar.Visible = True
End If
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape9.Visible = True
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape9.Visible = False

End Sub
Private Sub Label5_Click()
Call menulerikapat

mainpage.Visible = False
titlebar.Visible = False

about.Show
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.Visible = True
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.Visible = False

End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape6.Visible = True
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape6.Visible = False

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape7.Visible = True
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape7.Visible = False

End Sub


Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape8.Visible = True
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape8.Visible = False

End Sub

Private Sub TabStrip1_Change()

End Sub

Private Sub Label6_Click()
Call menulerikapat
titlebar.Visible = False
mainpage.Visible = False
guardplace.Show
End Sub

Private Sub MSFlexGrid1_Click()
Call menulerikapat
End Sub

Private Sub MSFlexGrid1_DblClick()
Call menulerikapat
End Sub

Private Sub MSFlexGrid1_GotFocus()
Call menulerikapat
End Sub
