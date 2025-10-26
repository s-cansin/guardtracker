VERSION 5.00
Object = "{1DF66C92-C83D-11DA-ADAC-0002449D97D0}#2.0#0"; "XPBUTTON.OCX"
Begin VB.Form guardplace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yeni Görev Ekleme"
   ClientHeight    =   2745
   ClientLeft      =   6330
   ClientTop       =   3435
   ClientWidth     =   4785
   Icon            =   "guardplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4785
   Begin Project1.UserControl1 UserControl13 
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Yeni Görev Ekle"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin Project1.UserControl1 UserControl12 
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
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
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Birim Kodu Al"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Birim Kodu [ iButton seri numarasý ]:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Görevlinin bulunmasý gereken yer:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "guardplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
titlebar.Visible = True
mainpage.Visible = True
End Sub



Private Sub UserControl11_Click()
buffer = ""
mainpage.MSComm1.Output = "*R"

For i = 1 To 1000000
For z = 1 To 10
Next
Next
buffer = mainpage.MSComm1.Input

If Left(buffer, 2) = "*W" Then
buffer = Right(buffer, Len(buffer) - 2)
buffer = Left(buffer, Len(buffer) - 1)

Text2.Text = buffer
End If

End Sub

Private Sub UserControl12_Click()
Unload Me
End Sub

Private Sub UserControl13_Click()
If Not Text1.Text = "" And Not Text2.Text = "" Then
Dim Conn As New ADODB.Connection
Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=db.mdb"
Conn.Open

Dim Rs As New ADODB.Recordset
Rs.Open "Select * from dp", Conn, adOpenKeyset, adLockOptimistic

Rs.AddNew

Rs!yer = Text1.Text
Rs!IButton = Text2.Text


Rs.Update
Rs.Close
MsgBox ("Belirttiðiniz görev detaylarý kayýtlara baþarýyla eklenmiþtir!!!")


Dim Rww As New ADODB.Recordset
Rww.Open "Select * from dp ORDER BY id desc", Conn, adOpenKeyset, adLockOptimistic


MsgBox (Rww!IButton & " numaralý iButton baþarýyla kayýt edilmiþtir...")

Unload Me
End If
End Sub
