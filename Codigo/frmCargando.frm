VERSION 5.00
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Argentum"
   ClientHeight    =   5940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Height          =   4560
      Left            =   0
      Picture         =   "frmCargando.frx":628A
      ScaleHeight     =   4500
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      Begin VB.Label verX 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "v?.?.?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   255
         TabIndex        =   2
         Top             =   0
         Width           =   555
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         DrawMode        =   3  'Not Merge Pen
         FillColor       =   &H00FF80FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   -120
         Shape           =   4  'Rounded Rectangle
         Top             =   -120
         Width           =   1335
      End
   End
   Begin VB.Image P6 
      Height          =   480
      Left            =   2355
      Picture         =   "frmCargando.frx":15180
      ToolTipText     =   "Funci�n de Trigger"
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trig."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBJ's"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1800
      TabIndex        =   7
      Top             =   5520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPC's"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Head"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Body"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label L 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BdD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image P5 
      Height          =   480
      Left            =   1320
      Picture         =   "frmCargando.frx":15DC2
      ToolTipText     =   "Objetos"
      Top             =   5400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image P1 
      Height          =   480
      Left            =   360
      Picture         =   "frmCargando.frx":16606
      ToolTipText     =   "Base de Datos"
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image P3 
      Height          =   480
      Left            =   2160
      Picture         =   "frmCargando.frx":16E4A
      ToolTipText     =   "Cabezas"
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image P4 
      Height          =   480
      Left            =   3120
      Picture         =   "frmCargando.frx":1768E
      ToolTipText     =   "NPC's"
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image P2 
      Height          =   480
      Left            =   1200
      Picture         =   "frmCargando.frx":182D0
      ToolTipText     =   "Cuerpos"
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label X 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4500
      Width           =   4335
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************
