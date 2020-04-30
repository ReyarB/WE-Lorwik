VERSION 5.00
Begin VB.Form frmImpCliente 
   Caption         =   "Importar archivos nesesarios al Editor"
   ClientHeight    =   6855
   ClientLeft      =   9210
   ClientTop       =   6405
   ClientWidth     =   5130
   Icon            =   "frmImpCliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   5130
   Begin VB.CommandButton Command2 
      Caption         =   "Importar del Server"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Importar del Cliente"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   240
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   3915
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.DirListBox Dir2 
      Appearance      =   0  'Flat
      Height          =   3465
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label LblCliente 
      Caption         =   "Seleccionar la carpeta Cliente    y luego la del Servidor "
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Disco:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Barra de estado:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   4935
   End
   Begin VB.Menu mnuMover 
      Caption         =   "Importar del Cliente"
   End
   Begin VB.Menu mnuServer 
      Caption         =   "Importar del Server"
   End
End
Attribute VB_Name = "frmImpCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Call mnuMover_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29 'Centra el formulario completamente
End Sub




Private Sub MnuSale_Click(index As Integer)
Dim Ruta, Ruta1, X, z
If MsgBox("Desea copiar los archivos del directorio:" + Chr$(10) + Dir1.Path + Chr$(10) + "A:" + Chr$(10) + Dir2.Path, 4 + 64 + 256, "Copiar archivos a otro directorio") = 6 Then
On Error Resume Next
If Right(Dir1.Path, 1) = "\" Then
  Ruta = Dir1.Path & ""
 Else
  Ruta = Dir1.Path & "\"
End If
If Right(Dir2.Path, 1) = "\" Then
  Ruta1 = Dir2.Path & ""
 Else
  Ruta1 = Dir2.Path & "\"
End If

Y = Ruta
z = Ruta1
result = xfilecopy("" & Y & "", "" & z & "", "*.*", Label1)
If err Then MsgBox "No existe el directorio de fuente ni del directorio destino", 16, "¡No copie nada!"


End If
End
End Sub

Private Function xfilecopy(origen$, destino$, archivo$, informa As Label)
' Copia varios archivos de una carpeta a otra
' Origen$= directorio de origen , terminado en "\"
' Destino$= directorio de destino , terminado en "\"
' archivo$= especificacion de archivos a copiar, con simb. comodin
' informa= un label en el que se muestra el progreso
'
' result = xfilecopy("c:\pat\", "h:\doc\", "*.exe", label1)
' copia todos los archivos exe de c:\pat en h:\doc
' muestra lo que esta haciendo en label1


Dim n, result, cuenta, pcent
' cuenta los archivos a copiar
cuenta = 0
n = Dir$(origen$ & archivo$)
While (n <> "")
 cuenta = cuenta + 1
 n = Dir$
Wend

' Copia
result = 0
n = Dir$(origen$ & archivo$)
On Error GoTo malxfilecopy
While (n <> "") And (result > -1)
 pcent = (result + 1) & "/" & cuenta & " "
 pcent = pcent & Format$(100 * result / cuenta, "#0.0") & "%"
 informa.Caption = pcent & " Copiando " & origen$ & n & " a " & destino$
 DoEvents

 FileCopy origen$ & n, destino$ & n
 result = result + 1
 n = Dir$
continuaxfilecopy:
Wend
informa.Caption = ""
xfilecopy = result
Exit Function

malxfilecopy:
 result = -1
 Resume continuaxfilecopy
End Function


Private Sub mnuMover_Click()
Dim Ruta, Ruta1, X, z
Dim wgraficos, wegraficos
Dim wminimapa, weminimapa

If MsgBox("Desea copiar los archivos del directorio:" + Chr$(10) + Dir1.Path + Chr$(10) + "A:" + Chr$(10) + Dir2.Path, 4 + 64 + 256, "Copiar archivos a otro directorio") = 6 Then
On Error Resume Next
If Right(Dir1.Path, 1) = "\" Then
  Ruta = Dir1.Path & ""
 Else
  Ruta = Dir1.Path & "\"
End If
If Right(Dir2.Path, 1) = "\" Then
  Ruta1 = Dir2.Path & ""
 Else
  Ruta1 = Dir2.Path & "\"
End If
'**************Rutas origen******************
Y = Ruta & "INIT\"
wgraficos = Ruta & "Graficos\"
wminimapa = Ruta & "Graficos\MiniMapa\"
'*************Rutas destinos******************
z = Ruta1 & "INIT\"
wegraficos = Ruta1 & "Recursos\graficos\"
weminimapa = Ruta1 & "Recursos\MiniMapa\"

'*************copiado*************************
result = xfilecopy("" & Y & "", "" & z & "", "*.ini", Label1)
result = xfilecopy("" & wgraficos & "", "" & wegraficos & "", "*.png", Label1)
result = xfilecopy("" & wminimapa & "", "" & weminimapa & "", "*.bmp", Label1)

If err Then MsgBox "No existe el directorio de fuente ni del directorio destino", 16, "¡No copie nada!"


End If
End Sub

Private Sub mnuServer_Click()
Dim Ruta, Ruta1, X, z
Dim wgraficos, wegraficos
Dim wminimapa, weminimapa
Dim wMapa, weMapa

If MsgBox("Desea copiar los archivos del directorio:" + Chr$(10) + Dir1.Path + Chr$(10) + "A:" + Chr$(10) + Dir2.Path, 4 + 64 + 256, "Copiar archivos a otro directorio") = 6 Then
On Error Resume Next
If Right(Dir1.Path, 1) = "\" Then
  Ruta = Dir1.Path & ""
 Else
  Ruta = Dir1.Path & "\"
End If
If Right(Dir2.Path, 1) = "\" Then
  Ruta1 = Dir2.Path & ""
 Else
  Ruta1 = Dir2.Path & "\"
End If
'**************Rutas origen******************
Y = Ruta & "Dat\"
wMapa = Ruta & "Mundos\Alkon\"

'*************Rutas destinos******************
z = Ruta1 & "Recursos\Dat\"
weMapa = Ruta1 & "Conversor\Mapas Long\"

'*************copiado*************************
result = xfilecopy("" & Y & "", "" & z & "", "NPCs.dat", Label1)
result = xfilecopy("" & Y & "", "" & z & "", "obj.dat", Label1)
result = xfilecopy("" & wMapa & "", "" & weMapa & "", "*.*", Label1)

If err Then MsgBox "No existe el directorio de fuente ni del directorio destino", 16, "¡No copie nada!"


End If
End Sub
