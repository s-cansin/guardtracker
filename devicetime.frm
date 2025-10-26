VERSION 5.00
Object = "{1DF66C92-C83D-11DA-ADAC-0002449D97D0}#2.0#0"; "XPBUTTON.OCX"
Begin VB.Form devicetime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saat Ýþlemleri"
   ClientHeight    =   2130
   ClientLeft      =   5670
   ClientTop       =   5115
   ClientWidth     =   5940
   Icon            =   "devicetime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5940
   Begin Project1.UserControl1 UserControl13 
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Aygýt Saatýný Al"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin Project1.UserControl1 UserControl12 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Aygýt Saatýný Güncelle"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
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
      Caption         =   "Sistem Saatýný Güncelle"
      ForeColor       =   -2147483630
      ForeHover       =   0
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   360
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Sistem Tarihi:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Sistem Saatý:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Aygýt Saatý:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "devicetime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Interval = 1000
End Sub



Private Sub Form_Unload(Cancel As Integer)
mainpage.Visible = True
titlebar.Visible = True
End Sub

Private Sub Text1_Click()
Timer1.Interval = 0
End Sub

Private Sub Timer1_Timer()
Text1.Text = Time
Text3.Text = Date
End Sub



Private Sub UserControl11_Click()
Time$ = Text1.Text
Date = Text3.Text
Timer1.Interval = 1000
End Sub

Private Sub UserControl12_Click()

If Text1.Text <> "" And Text3.Text <> "" Then

mainpage.MSComm1.Output = "*A"
mainpage.MSComm1.Output = Text1.Text & " " & Text3.Text
Label4.Caption = ""

End If

End Sub

Private Sub UserControl13_Click()

mainpage.MSComm1.Output = "*C"

'>>>>>PIC TEN GELEN BÝLGÝYÝ BEKLETME MODU
For i = 1 To 1000000
For z = 1 To 10
Next z
Next i
'<<<<<<<

Dim buffer
buffer = mainpage.MSComm1.Input
Dim data As String
If Left(buffer, 2) = "?S" And Right(buffer, 1) = "&" Then
buffer = Right(buffer, Len(buffer) - 2)
buffer = Left(buffer, Len(buffer) - 1)

Label4.Caption = buffer
End If
End Sub
