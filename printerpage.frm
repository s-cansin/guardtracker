VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1DF66C92-C83D-11DA-ADAC-0002449D97D0}#2.0#0"; "XPBUTTON.OCX"
Begin VB.Form printerpage 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yazdýrma Yöneticisi"
   ClientHeight    =   1545
   ClientLeft      =   6000
   ClientTop       =   5115
   ClientWidth     =   5400
   Icon            =   "printerpage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5400
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.UserControl1 UserControl12 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ýptal"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tüm Kayýtlarý Yazdýr"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Bu yüzden kayýt yazdýrma iþlemini yapamazsýnýz..."
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Toplam  XX   bekçi faaliyet kaydý bulunmaktadýr..."
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "printerpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub sayfaolustur()


Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from logs order by id desc", Conn, adOpenKeyset, adLockOptimistic



If Rs.RecordCount > 0 Then






'Bir sayfada kaç kayýt olacak????
kayitsayisi = 65



formprint.Label10.Caption = Date

If Rs.RecordCount < kayitsayisi + 1 Then
Call temizle
formprint.Label9.Caption = "Sayfa No: 1"
For r = 1 To Rs.RecordCount
formprint.Label1.Caption = formprint.Label1.Caption & mainpage.MSFlexGrid1.TextMatrix(r, 0) & vbCrLf
formprint.Label2.Caption = formprint.Label2.Caption & mainpage.MSFlexGrid1.TextMatrix(r, 1) & vbCrLf
formprint.Label3.Caption = formprint.Label3.Caption & mainpage.MSFlexGrid1.TextMatrix(r, 2) & vbCrLf
formprint.Label4.Caption = formprint.Label4.Caption & mainpage.MSFlexGrid1.TextMatrix(r, 3) & vbCrLf
formprint.Label5.Caption = formprint.Label5.Caption & mainpage.MSFlexGrid1.TextMatrix(r, 4) & vbCrLf
formprint.Label6.Caption = formprint.Label6.Caption & mainpage.MSFlexGrid1.TextMatrix(r, 5) & vbCrLf

Next r
On Error GoTo Hata
totalsayfa = 1
formprint.PrintForm
End If



If Rs.RecordCount > kayitsayisi Then


kacsayfa = Fix(Rs.RecordCount / kayitsayisi)
fark = Rs.RecordCount - xtx * kayitsayisi



For t = 1 To kacsayfa
Call temizle

formprint.Label9.Caption = "Sayfa No: " & t
For g = t * kayitsayisi - kayitsayisi + 1 To t * kayitsayisi

formprint.Label1.Caption = formprint.Label1.Caption & mainpage.MSFlexGrid1.TextMatrix(g, 0) & vbCrLf
formprint.Label2.Caption = formprint.Label2.Caption & mainpage.MSFlexGrid1.TextMatrix(g, 1) & vbCrLf
formprint.Label3.Caption = formprint.Label3.Caption & mainpage.MSFlexGrid1.TextMatrix(g, 2) & vbCrLf
formprint.Label4.Caption = formprint.Label4.Caption & mainpage.MSFlexGrid1.TextMatrix(g, 3) & vbCrLf
formprint.Label5.Caption = formprint.Label5.Caption & mainpage.MSFlexGrid1.TextMatrix(g, 4) & vbCrLf
formprint.Label6.Caption = formprint.Label6.Caption & mainpage.MSFlexGrid1.TextMatrix(g, 5) & vbCrLf

Next g
On Error GoTo Hata

formprint.PrintForm

totalsayfa = t
Next t




If fark > 0 Then
Call temizle
formprint.Label9.Caption = "Sayfa No: " & totalsayfa + 1
For X = kacsayfa * kayitsayisi + 1 To fark
formprint.Label1.Caption = formprint.Label1.Caption & mainpage.MSFlexGrid1.TextMatrix(X, 0) & vbCrLf
formprint.Label2.Caption = formprint.Label2.Caption & mainpage.MSFlexGrid1.TextMatrix(X, 1) & vbCrLf
formprint.Label3.Caption = formprint.Label3.Caption & mainpage.MSFlexGrid1.TextMatrix(X, 2) & vbCrLf
formprint.Label4.Caption = formprint.Label4.Caption & mainpage.MSFlexGrid1.TextMatrix(X, 3) & vbCrLf
formprint.Label5.Caption = formprint.Label5.Caption & mainpage.MSFlexGrid1.TextMatrix(X, 4) & vbCrLf
formprint.Label6.Caption = formprint.Label6.Caption & mainpage.MSFlexGrid1.TextMatrix(X, 5) & vbCrLf


Next X
On Error GoTo Hata

formprint.PrintForm
End If





End If















End If

Hata:
ff = MsgBox("Dikkat!!! Yazdýrma iþlemi ile ilgili bir problem oluþtu..." & vbCrLf & "Lütfen yazýcýnýzýn ayarlarýný denetleyin ve yazdýrma iþlemini tekrar deneyin...", vbExclamation, "Hata!")

End Sub
Sub temizle()
formprint.Label1.Caption = ""
formprint.Label2.Caption = ""
formprint.Label3.Caption = ""
formprint.Label4.Caption = ""
formprint.Label5.Caption = ""
formprint.Label6.Caption = ""
End Sub



Private Sub Form_Load()
Label2.Visible = False


Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from logs order by id desc", Conn, adOpenKeyset, adLockOptimistic



Label1.Caption = "Toplam " & Rs.RecordCount & " bekçi faaliyet kaydý bulunmaktadýr..."

If Rs.RecordCount = 0 Then
UserControl11.Visible = False
Label2.Visible = True
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
mainpage.Visible = True
titlebar.Visible = True
End Sub



Private Sub UserControl11_Click()
Me.Visible = False

wait.Show

DoEvents
On Error Resume Next
CommonDialog1.Action = 5
Call sayfaolustur
Unload wait
Me.Visible = True


End Sub

Private Sub UserControl12_Click()
Unload Me
End Sub
