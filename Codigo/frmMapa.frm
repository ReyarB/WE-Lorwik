VERSION 5.00
Begin VB.Form frmMapa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmMapa"
   ClientHeight    =   12105
   ClientLeft      =   7245
   ClientTop       =   1815
   ClientWidth     =   15090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   807
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1006
   ShowInTaskbar   =   0   'False
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   16
      Left            =   0
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   15
      Left            =   12000
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   14
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   13
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   12
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   11
      Left            =   0
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   10
      Left            =   12000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   9
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   8
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   7
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   6
      Left            =   0
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   5
      Left            =   12000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   4
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   3
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   2
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   1
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Index           =   0
      Left            =   11640
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   1455
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
 Call ActualizarMapa
End Sub

Private Sub ActualizarMapa() ' modificar  no funciona ReyarB
  Dim i As Byte
  For i = 0 To 16 ' 99 original
    Mapa(i).Stretch = True
    'App.Path & "\Recursos\Graficos\MiniMapa
    If Len(Dir(App.Path & "\Recursos\MiniMapa\" & i + 1 & ".bmp", vbNormal)) > 0 Then
      Set Mapa(i).Picture = LoadPicture(App.Path & "\Recursos\MiniMapa\" & i + 1 & ".bmp")
    Else
     Set Mapa(i).Picture = LoadPicture(App.Path & "\Recursos\MiniMapa\no.bmp")
    End If
  Next i
End Sub

Private Sub Mapa_Click(index As Integer)
  If Len(Dir(App.Path & "\Conversor\Mapas CSM" & index + 1 & ".csm", vbNormal)) > 0 Then
    Call modMapIO.NuevoMapa
    Call modMapIO.Cargar_CSM("Mapa" & index + 1 & ".csm")
    Call ActualizarMapa
  End If
End Sub
