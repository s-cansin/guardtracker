VERSION 5.00
Begin VB.Form wait 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   855
   ClientLeft      =   6945
   ClientTop       =   6180
   ClientWidth     =   4455
   Icon            =   "wait.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Lütfen Bekleyiniz... Ýþleminiz Yapýlýyor"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1440
      Picture         =   "wait.frx":1272
      Top             =   240
      Width           =   1350
   End
End
Attribute VB_Name = "wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
