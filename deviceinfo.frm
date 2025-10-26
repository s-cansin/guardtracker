VERSION 5.00
Object = "{1DF66C92-C83D-11DA-ADAC-0002449D97D0}#2.0#0"; "XPBUTTON.OCX"
Begin VB.Form deviceinfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cihaz Bilgileri"
   ClientHeight    =   2010
   ClientLeft      =   6330
   ClientTop       =   5115
   ClientWidth     =   4185
   Icon            =   "deviceinfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4185
   Begin Project1.UserControl1 UserControl12 
      Height          =   405
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cihaz bilgisini bekçi adýna kaydet"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cihaz No Al"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Bekçi:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Cihaz No:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "deviceinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
End Sub



Private Sub Form_Unload(Cancel As Integer)
titlebar.Visible = True
mainpage.Visible = True
End Sub

Private Sub UserControl11_Click()
mainpage.MSComm1.Output = "*N"

'>>>>>PIC TEN GELEN BÝLGÝYÝ BEKLETME MODU
For i = 1 To 100000
For z = 1 To 10
Next z
Next i
'<<<<<<<

Dim buffer
buffer = mainpage.MSComm1.Input

If Left(buffer, 3) = "#DN" Then
buffer = Right(buffer, Len(buffer) - 3)

Text1.Text = buffer

End If

End Sub

Private Sub UserControl12_Click()

Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from guard", Conn, adOpenKeyset, adLockOptimistic


If Not Text2.Text = "" And Not Text1.Text = "" Then

Rs.AddNew
Rs("ad") = Text2.Text
Rs("kod") = Text1.Text
Rs.Update
mesaj = MsgBox("Ýlgili görev yeri veritabanýna baþarýyla kaydedilmiþtir!")
Unload Me
End If
End Sub
