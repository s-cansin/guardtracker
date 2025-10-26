VERSION 5.00
Begin VB.Form formprint 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   17520
   ClientLeft      =   4635
   ClientTop       =   1005
   ClientWidth     =   10965
   Icon            =   "formprint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   17520
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8880
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BEKÇÝ KONTROL RAPORU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ÇINARLI ANADOLU TEKNÝK VE ENDÜSTRÝ MESLEK LÝSESÝ "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   13725
      Left            =   8880
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   13710
      Left            =   7440
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   13710
      Left            =   6000
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   13710
      Left            =   4560
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   13710
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   13710
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Line Line13 
      X1              =   1560
      X2              =   10200
      Y1              =   15360
      Y2              =   15360
   End
   Begin VB.Line Line12 
      X1              =   1560
      X2              =   1560
      Y1              =   8280
      Y2              =   15360
   End
   Begin VB.Line Line11 
      X1              =   10200
      X2              =   10200
      Y1              =   1560
      Y2              =   15360
   End
   Begin VB.Line Line10 
      X1              =   8760
      X2              =   8760
      Y1              =   1320
      Y2              =   15360
   End
   Begin VB.Line Line9 
      X1              =   1560
      X2              =   1560
      Y1              =   1560
      Y2              =   9720
   End
   Begin VB.Line Line8 
      X1              =   7320
      X2              =   7320
      Y1              =   1320
      Y2              =   15360
   End
   Begin VB.Line Line7 
      X1              =   5880
      X2              =   5880
      Y1              =   1320
      Y2              =   15360
   End
   Begin VB.Line Line6 
      X1              =   4440
      X2              =   4440
      Y1              =   1320
      Y2              =   15360
   End
   Begin VB.Line Line5 
      X1              =   3000
      X2              =   3000
      Y1              =   1320
      Y2              =   15360
   End
   Begin VB.Label sutunlar 
      BackStyle       =   0  'Transparent
      Caption         =   $"formprint.frx":1272
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   1350
      Width           =   8295
   End
   Begin VB.Line Line4 
      X1              =   10200
      X2              =   10200
      Y1              =   1320
      Y2              =   1560
   End
   Begin VB.Line Line3 
      X1              =   1560
      X2              =   10200
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      X1              =   1560
      X2              =   1560
      Y1              =   1320
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   10200
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "formprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

