VERSION 5.00
Begin VB.Form frmMapa 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "frmMapa"
   ClientHeight    =   14160
   ClientLeft      =   -5235
   ClientTop       =   1005
   ClientWidth     =   49905
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
   ScaleHeight     =   944
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   3327
   ShowInTaskbar   =   0   'False
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   16
      Left            =   45000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   15
      Left            =   42000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   14
      Left            =   39000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   13
      Left            =   36000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   12
      Left            =   33000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   11
      Left            =   30000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   10
      Left            =   27000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   9
      Left            =   24000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   8
      Left            =   21000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   7
      Left            =   18000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image Mapa 
      Appearance      =   0  'Flat
      Height          =   3000
      Index           =   6
      Left            =   15000
      Stretch         =   -1  'True
      Top             =   0
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
      Left            =   48720
      Stretch         =   -1  'True
      Top             =   4440
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

Private Sub ActualizarMapa() ' modificar ReyarB
  Dim i As Byte
  For i = 0 To 99
    Mapa(i).Stretch = True
    'App.Path & "\Recursos\Graficos\MiniMapa
    If Len(Dir("C:\Users\Administrador\Desktop\AO1024\WE-Lorwik\Recursos\Graficos\MiniMapa\" & i + 1 & ".bmp", vbNormal)) > 0 Then
      Set Mapa(i).Picture = LoadPicture("C:\Users\Administrador\Desktop\AO1024\WE-Lorwik\Recursos\Graficos\MiniMapa\" & i + 1 & ".bmp")
    Else
     Set Mapa(i).Picture = LoadPicture("C:\Users\Administrador\Desktop\AO1024\WE-Lorwik\Recursos\Graficos\MiniMapa\no.bmp")
    End If
  Next i
End Sub

Private Sub Mapa_Click(index As Integer)
  If Len(Dir("C:\Users\Administrador\Desktop\AO1024\WE-Lorwik\Conversor\Mapas CSM" & index + 1 & ".csm", vbNormal)) > 0 Then
    Call modMapIO.NuevoMapa
    Call modMapIO.Cargar_CSM("Mapa" & index + 1 & ".csm")
    Call ActualizarMapa
  End If
End Sub
