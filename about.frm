VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Hakkýnda"
   ClientHeight    =   3735
   ClientLeft      =   5505
   ClientTop       =   5115
   ClientWidth     =   6525
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6525
   Begin VB.Label Label7 
      Caption         =   "®"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2610
      TabIndex        =   6
      Top             =   3350
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "ÇINARLI ANADOLU MESLEK LÝSESÝ ELEKTRONÝK BÖLÜMÜ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "http://www.samedcansin.com"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      FillColor       =   &H00808080&
      Height          =   3000
      Left            =   3840
      Top             =   120
      Width           =   2600
   End
   Begin VB.Label Label5 
      Caption         =   "Ýrtibat"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "COPYRIGHT 2006     SAMED CANSIN ESKÝOÐLU"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   3360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "samedcansin@hotmail.com"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2985
      Left            =   3840
      Picture         =   "about.frx":1272
      Top             =   120
      Width           =   2580
   End
   Begin VB.Label Label2 
      Caption         =   $"about.frx":1A3D0
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
mainpage.Visible = True
titlebar.Visible = True
End Sub

