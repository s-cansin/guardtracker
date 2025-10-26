VERSION 5.00
Object = "{1DF66C92-C83D-11DA-ADAC-0002449D97D0}#2.0#0"; "XPBUTTON.OCX"
Begin VB.Form comport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ýletiþim Portu"
   ClientHeight    =   1410
   ClientLeft      =   6660
   ClientTop       =   5280
   ClientWidth     =   4680
   Icon            =   "comport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   Begin Project1.UserControl1 UserControl12 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
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
      Left            =   840
      TabIndex        =   3
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "Ayarla"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin VB.OptionButton Option2 
      Caption         =   "COM2"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "COM1"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Kullanýlacak iletiþim portu:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "comport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If mainpage.MSComm1.CommPort = 1 Then
Option1.Value = True
Else
Option2.Value = True
End If

mainpage.MSComm1.PortOpen = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
mainpage.Visible = True
titlebar.Visible = True
End Sub

Private Sub UserControl11_Click()
If Option1.Value = True Then
CommPort1 = 1
Else
CommPort1 = 2
End If


Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from sets", Conn, adOpenKeyset, adLockOptimistic


Rs("port") = CommPort1
Rs.Update


mainpage.MSComm1.CommPort = CommPort1
mainpage.MSComm1.PortOpen = True

Unload Me
End Sub

Private Sub UserControl12_Click()
Unload Me
End Sub
