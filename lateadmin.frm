VERSION 5.00
Object = "{1DF66C92-C83D-11DA-ADAC-0002449D97D0}#2.0#0"; "XPBUTTON.OCX"
Begin VB.Form lateadmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gecikme Yöneticisi"
   ClientHeight    =   1425
   ClientLeft      =   6495
   ClientTop       =   5115
   ClientWidth     =   4755
   Icon            =   "lateadmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4755
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Tamam"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin VB.Label Label1 
      Caption         =   "dakika gecikme yaptýðý faaliyetleri iþaretle"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   285
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Görevlinin"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   285
      Width           =   735
   End
End
Attribute VB_Name = "lateadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

Dim Connw As New ADODB.Connection
Connw.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Connw.Open

Dim Rx As New ADODB.Recordset
Rx.Open "Select * from sets", Connw, adOpenKeyset, adLockOptimistic

Text1.Text = Rx("gecikme")
Rx.Close
Set Rx = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
titlebar.Visible = True
mainpage.Visible = True

End Sub

Private Sub UserControl11_Click()
If Not Text1.Text = "" Then

Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open


Dim Rs As New ADODB.Recordset
Rs.Open "Select * from sets", Conn, adOpenKeyset, adLockOptimistic


If Len(Text1.Text) = 1 Then Text1.Text = "0" & Text1.Text
Rs!gecikme = Text1.Text
Rs.Update

Unload Me
End If
End Sub
