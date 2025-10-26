VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1DF66C92-C83D-11DA-ADAC-0002449D97D0}#2.0#0"; "XPBUTTON.OCX"
Begin VB.Form tablecolors 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tablo Renk Ayarlarý"
   ClientHeight    =   1560
   ClientLeft      =   6495
   ClientTop       =   5115
   ClientWidth     =   4560
   Icon            =   "tablecolors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4560
   Begin Project1.UserControl1 UserControl13 
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Deðiþiklikleri Ýptal Et"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin Project1.UserControl1 UserControl12 
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Varsayýlan Renk Ayarlarý"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ayarlarý Güncelle"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   1560
      ScaleHeight     =   38.235
      ScaleMode       =   0  'User
      ScaleWidth      =   152.941
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   1560
      ScaleHeight     =   104
      ScaleMode       =   0  'User
      ScaleWidth      =   104
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Left            =   1560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   1080
      Width           =   255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Baþlýk Rengi:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Satýr Rengi:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Yazý Rengi:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "tablecolors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
mainpage.Visible = False
titlebar.Visible = False

Picture1.BackColor = mainpage.MSFlexGrid1.ForeColor
Picture2.BackColor = mainpage.MSFlexGrid1.BackColor
Picture3.BackColor = mainpage.MSFlexGrid1.BackColorFixed
End Sub

Private Sub Form_Unload(Cancel As Integer)
titlebar.Visible = True
mainpage.Visible = True
End Sub

Private Sub Label1_Click()
CommonDialog1.ShowColor
If CommonDialog1.Color <> 0 Then
Picture1.BackColor = CommonDialog1.Color
End If
End Sub

Private Sub Label2_Click()
CommonDialog1.ShowColor
If CommonDialog1.Color <> 0 Then
Picture2.BackColor = CommonDialog1.Color

End If
End Sub

Private Sub Label3_Click()
CommonDialog1.ShowColor
If CommonDialog1.Color <> 0 Then
Picture3.BackColor = CommonDialog1.Color
End If
End Sub

Private Sub Picture1_Click()
CommonDialog1.ShowColor
If CommonDialog1.Color <> 0 Then
Picture1.BackColor = CommonDialog1.Color

End If

End Sub

Private Sub Picture2_Click()

CommonDialog1.ShowColor
If CommonDialog1.Color <> 0 Then
Picture2.BackColor = CommonDialog1.Color

End If

End Sub

Private Sub Picture3_Click()
CommonDialog1.ShowColor
If CommonDialog1.Color <> 0 Then
Picture3.BackColor = CommonDialog1.Color
End If

End Sub

Private Sub UserControl11_Click()

mainpage.MSFlexGrid1.ForeColor = Picture1.BackColor
mainpage.MSFlexGrid1.BackColor = Picture2.BackColor
mainpage.MSFlexGrid1.BackColorFixed = Picture3.BackColor

Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from sets", Conn, adOpenKeyset, adLockOptimistic

Rs!baslik = Picture1.BackColor
Rs!satir = Picture2.BackColor
Rs!yazi = Picture3.BackColor

Rs.Update

Unload Me
End Sub

Private Sub UserControl12_Click()
Picture1.BackColor = 0
Picture2.BackColor = 12648447
Picture3.BackColor = 14671839

End Sub

Private Sub UserControl13_Click()
Unload Me
End Sub
