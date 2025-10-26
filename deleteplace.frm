VERSION 5.00
Object = "{1DF66C92-C83D-11DA-ADAC-0002449D97D0}#2.0#0"; "XPBUTTON.OCX"
Begin VB.Form deleteplace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Görev Silme"
   ClientHeight    =   2850
   ClientLeft      =   6330
   ClientTop       =   5115
   ClientWidth     =   4680
   Icon            =   "deleteplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   Begin Project1.UserControl1 UserControl12 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   2400
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
      Caption         =   "Ýptal"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Ýlgili Görevi Sil"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin VB.ListBox List2 
      Height          =   255
      ItemData        =   "deleteplace.frx":1272
      Left            =   4200
      List            =   "deleteplace.frx":1274
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "deleteplace.frx":1276
      Left            =   720
      List            =   "deleteplace.frx":1278
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "deleteplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
List1.Clear
List2.Clear

Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from dp", Conn, adOpenKeyset, adLockOptimistic

If Not Rs.EOF Or Not BOF Then

For i = 1 To Rs.RecordCount

List1.AddItem Rs("IButton") & "     " & Rs("yer")
List2.AddItem Rs("IButton")
Rs.MoveNext
Next

End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
mainpage.Visible = True
titlebar.Visible = True
End Sub

Private Sub UserControl11_Click()

If List1.ListIndex <> "-1" Then


cc = MsgBox("Onayýnýz ile birlikte ilgili görev yerinin geçmiþ tüm kayýtlarý ve eylemleri silinecektir!!!" & vbCrLf & "Gerçekten silmek istediðinize emin misiniz?", vbExclamation + vbOKCancel, "Dikkat")
If cc = vbOK Then

Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Valw = List2.List(List1.ListIndex)

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from dp where IButton='" & Valw & "'", Conn, adOpenKeyset, adLockOptimistic

If Not Rs.EOF Or Not BOF Then
Rs.Delete
End If


Dim Rx As New ADODB.Recordset
Rx.Open "Select * from logs where yer='" & Valw & "'", Conn, adOpenKeyset, adLockOptimistic

If Not Rx.EOF Or Not BOF Then
For X = 1 To Rx.RecordCount

Rx.Delete
Rx.MoveNext
Next
End If

Unload Me
End If

End If
End Sub

Private Sub UserControl12_Click()
Unload Me
End Sub
