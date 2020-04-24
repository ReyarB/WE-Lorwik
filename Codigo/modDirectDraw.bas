Attribute VB_Name = "Mod_TileEngine"
Option Explicit

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1
Public bTecho As Boolean 'hay techo?

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Integer, ByRef tY As Integer)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    tX = UserPos.X + viewPortX \ 32 - frmMain.Renderer.ScaleWidth \ 64
    tY = UserPos.Y + viewPortY \ 32 - frmMain.Renderer.ScaleHeight \ 64
End Sub

Sub MakeChar(CharIndex As Integer, Body As Integer, Head As Integer, Heading As Byte, X As Integer, Y As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
On Error Resume Next

'Update LastChar
If CharIndex > LastChar Then LastChar = CharIndex
NumChars = NumChars + 1

'Update head, body, ect.
CharList(CharIndex).Body = BodyData(Body)
CharList(CharIndex).Head = Head
CharList(CharIndex).Heading = Heading

'Reset moving stats
CharList(CharIndex).Moving = 0
CharList(CharIndex).MoveOffset.X = 0
CharList(CharIndex).MoveOffset.Y = 0

'Update position
CharList(CharIndex).Pos.X = X
CharList(CharIndex).Pos.Y = Y

'Make active
CharList(CharIndex).Active = 1

'Plot on map
MapData(X, Y).CharIndex = CharIndex

bRefreshRadar = True ' GS

End Sub

Sub EraseChar(CharIndex As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
If CharIndex = 0 Then Exit Sub
'Make un-active
CharList(CharIndex).Active = 0

'Update lastchar
If CharIndex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If

MapData(CharList(CharIndex).Pos.X, CharList(CharIndex).Pos.Y).CharIndex = 0

'Update NumChars
NumChars = NumChars - 1

bRefreshRadar = True ' GS

End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Started As Byte = 2)
On Error Resume Next
'*****************************************************************
'Sets up a grh. MUST be done before rendering
'*****************************************************************
If GrhIndex <= 0 Or GrhIndex > GrhCount Then GrhIndex = 20299
    Grh.GrhIndex = GrhIndex
    
    If Started = 2 Then
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If
    Else
        'Make sure the graphic can be started
        If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
    
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
    
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.GrhIndex).Speed
End Sub

Sub MoveCharbyHead(CharIndex As Integer, nHeading As Byte)
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim addx As Integer
Dim addy As Integer
Dim X As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

X = CharList(CharIndex).Pos.X
Y = CharList(CharIndex).Pos.Y

'Figure out which way to move
Select Case nHeading

    Case NORTH
        addy = -1

    Case EAST
        addx = 1

    Case SOUTH
        addy = 1
    
    Case WEST
        addx = -1
        
End Select

nX = X + addx
nY = Y + addy

MapData(nX, nY).CharIndex = CharIndex
CharList(CharIndex).Pos.X = nX
CharList(CharIndex).Pos.Y = nY
MapData(X, Y).CharIndex = 0

CharList(CharIndex).MoveOffset.X = -1 * (TilePixelWidth * addx)
CharList(CharIndex).MoveOffset.Y = -1 * (TilePixelHeight * addy)

CharList(CharIndex).Moving = 1
CharList(CharIndex).Heading = nHeading

End Sub

Sub MoveCharbyPos(CharIndex As Integer, nX As Integer, nY As Integer)
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 by GS
'*************************************************
    Dim X As Integer
    Dim Y As Integer
    Dim addx As Integer
    Dim addy As Integer
    Dim nHeading As Byte
    
    With CharList(CharIndex)
        
        X = .Pos.X
        Y = .Pos.Y
        
        addx = nX - X
        addy = nY - Y
        
        If Sgn(addx) = 1 Then
            nHeading = EAST
        End If
        
        If Sgn(addx) = -1 Then
            nHeading = WEST
        End If
        
        If Sgn(addy) = -1 Then
            nHeading = NORTH
        End If
        
        If Sgn(addy) = 1 Then
            nHeading = SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffset.X = -1 * (TilePixelWidth * addx)
        .MoveOffset.Y = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        bRefreshRadar = True ' GS
    End With
End Sub


Function NextOpenChar() As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
Dim loopc As Integer

loopc = 1
Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function

Function LegalPos(X As Integer, Y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 28/05/06 - GS
'*************************************************

LegalPos = True

'Check to see if its out of bounds
If X - 8 < YMinMapSize Or X - 8 > XMaxMapSize Or Y - 6 < YMinMapSize Or Y - 6 > YMaxMapSize Then
    LegalPos = False
    Exit Function
End If

If X > XMaxMapSize Or X < XMinMapSize Then Exit Function
If Y > YMaxMapSize Or Y < YMinMapSize Then Exit Function

'Check to see if its blocked
If MapData(X, Y).blocked = 1 Then
    LegalPos = False
    Exit Function
End If

'Check for character
If MapData(X, Y).CharIndex > 0 Then
    LegalPos = False
    Exit Function
End If

End Function




Function InMapLegalBounds(X As Integer, Y As Integer) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

' [Loopzer]
Public Sub DePegar()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer

    For X = 0 To DeSeleccionAncho - 1
        For Y = 0 To DeSeleccionAlto - 1
             MapData(X + DeSeleccionOX, Y + DeSeleccionOY) = DeSeleccionMap(X, Y)
        Next
    Next
End Sub
Public Sub PegarSeleccion() '(mx As Integer, my As Integer)
On Error GoTo ErrHandler:

'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Static UltimoX As Integer
    Static UltimoY As Integer
    If UltimoX = SobreX And UltimoY = SobreY Then Exit Sub
    UltimoX = SobreX
    UltimoY = SobreY
    Dim X As Integer
    Dim Y As Integer
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SobreX
    DeSeleccionOY = SobreY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To DeSeleccionAncho - 1
        For Y = 0 To DeSeleccionAlto - 1
            If X <= MaxXBorder Then
              If X >= 0 Then
                If Y <= MaxYBorder Then
                  If Y >= 0 Then
                    DeSeleccionMap(X, Y) = MapData(X + SobreX, Y + SobreY)
                  End If
                End If
              End If
            End If
            
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            If X <= MaxXBorder Then
              If X >= 0 Then
                If Y <= MaxYBorder Then
                  If Y >= 0 Then
                 MapData(X + SobreX, Y + SobreY) = SeleccionMap(X, Y)
              End If
                End If
              End If
            End If
        Next
    Next
    Seleccionando = False
ErrHandler:
    
End Sub
Public Sub AccionSeleccion()
On Error GoTo ErrHandler:
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    Dim DeSeleccionMap() As MapBlock
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
           ClickEdit vbLeftButton, SeleccionIX + X, SeleccionIY + Y
        Next
    Next
    Seleccionando = False
ErrHandler:
    
End Sub

Public Sub BlockearSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    Dim X As Integer
    Dim Y As Integer
    Dim Vacio As MapBlock
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             If MapData(X + SeleccionIX, Y + SeleccionIY).blocked = 1 Then
                MapData(X + SeleccionIX, Y + SeleccionIY).blocked = 0
             Else
                MapData(X + SeleccionIX, Y + SeleccionIY).blocked = 1
            End If
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CortarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    CopiarSeleccion
    Dim X As Integer
    Dim Y As Integer
    Dim Vacio As MapBlock
    DeSeleccionAncho = SeleccionAncho
    DeSeleccionAlto = SeleccionAlto
    DeSeleccionOX = SeleccionIX
    DeSeleccionOY = SeleccionIY
    ReDim DeSeleccionMap(DeSeleccionAncho, DeSeleccionAlto) As MapBlock
    
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            DeSeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
             MapData(X + SeleccionIX, Y + SeleccionIY) = Vacio
        Next
    Next
    Seleccionando = False
End Sub
Public Sub CopiarSeleccion()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
    'podria usar copy mem , pero por las dudas no XD
    Dim X As Integer
    Dim Y As Integer
    Seleccionando = False
    SeleccionAncho = Abs(SeleccionIX - SeleccionFX) + 1
    SeleccionAlto = Abs(SeleccionIY - SeleccionFY) + 1
    ReDim SeleccionMap(SeleccionAncho, SeleccionAlto) As MapBlock
    For X = 0 To SeleccionAncho - 1
        For Y = 0 To SeleccionAlto - 1
            SeleccionMap(X, Y) = MapData(X + SeleccionIX, Y + SeleccionIY)
        Next
    Next
End Sub
Public Sub GenerarVista()
'*************************************************
'Author: Loopzer
'Last modified: 21/11/07
'*************************************************
   ' hacer una llamada a un seter o geter , es mas lento q una variable
   ' con esto hacemos q no este preguntando a el objeto cadavez
   ' q dibuja , Render mas rapido ;)
    VerBlockeados = frmMain.cVerBloqueos.value
    VerTriggers = frmMain.cVerTriggers.value
    VerCapa1 = frmMain.mnuVerCapa1.Checked
    VerCapa2 = frmMain.mnuVerCapa2.Checked
    VerCapa3 = frmMain.mnuVerCapa3.Checked
    VerCapa4 = frmMain.mnuVerCapa4.Checked
    VerTranslados = frmMain.mnuVerTranslados.Checked
    VerObjetos = frmMain.mnuVerObjetos.Checked
    VerNpcs = frmMain.mnuVerNPCs.Checked
    
End Sub

Function HayUserAbajo(X As Integer, Y As Integer, GrhIndex) As Boolean
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************
HayUserAbajo = _
    CharList(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
And CharList(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
And CharList(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
And CharList(UserCharIndex).Pos.Y <= Y
End Function

Function PixelPos(X As Integer) As Integer
'*************************************************
'Author: Unkwown
'Last modified: 20/05/06
'*************************************************

PixelPos = (TilePixelWidth * X) - TilePixelWidth

End Function

Public Function ARGB(ByVal R As Long, ByVal G As Long, ByVal B As Long, ByVal A As Long) As Long
        
    Dim c As Long
        
    If A > 127 Then
        A = A - 128
        c = A * 2 ^ 24 Or &H80000000
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or B
    Else
        c = A * 2 ^ 24
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or B
    End If
    
    ARGB = c

End Function

Sub DrawGrhtoHdc(picX As PictureBox, Grh As Long, ByVal X As Integer, ByVal Y As Integer)
    Dim destRect As RECT
    
    destRect.Bottom = picX.ScaleHeight / 16
    destRect.Right = picX.ScaleWidth / 16
    destRect.Left = 0
    destRect.Top = 0
    
    D3DDevice.BeginScene
    'D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    Draw_GrhIndex Grh, X, Y, LightIluminado()
    D3DDevice.EndScene
    D3DDevice.Present destRect, ByVal 0, picX.hwnd, ByVal 0
End Sub

'**************************************************************
Public Sub DibujarMiniMapa()
Dim map_x, map_y, Capas As Byte
Dim loopc As Long
    For map_y = YMinMapSize To YMaxMapSize
        For map_x = XMinMapSize To XMaxMapSize
            For Capas = 1 To 2
                If MapData(map_x, map_y).Graphic(Capas).GrhIndex > 0 Then
                    SetPixel frmMain.Minimap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(Capas).GrhIndex).MiniMap_color
                End If
                If MapData(map_x, map_y).Graphic(4).GrhIndex > 0 And frmMain.chkOptMinimap(2).value = 1 Then
                    SetPixel frmMain.Minimap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color
                End If
                If MapData(map_x, map_y).blocked = 1 And frmMain.chkOptMinimap(3).value = 1 Then
                    SetPixel frmMain.Minimap.hdc, map_x - 1, map_y - 1, RGB(255, 0, 0)
                End If
                If MapData(map_x, map_y).ObjGrh.GrhIndex <> 0 And frmMain.chkOptMinimap(1).value = 1 Then
                    SetPixel frmMain.Minimap.hdc, map_x - 1, map_y - 1, RGB(51, 196, 255)
                    SetPixel frmMain.Minimap.hdc, map_x - 2, map_y - 2, RGB(51, 196, 255)
                End If
            Next Capas
        Next map_x
    Next map_y
    
    If frmMain.chkOptMinimap(0).value = 1 Then
        For loopc = 1 To LastChar
            If CharList(loopc).Active = 1 Then
                MapData(CharList(loopc).Pos.X, CharList(loopc).Pos.Y).CharIndex = loopc
                If CharList(loopc).Heading <> 0 And frmMain.chkOptMinimap(0).value = 1 Then
                    SetPixel frmMain.Minimap.hdc, 0 + CharList(loopc).Pos.X, 0 + CharList(loopc).Pos.Y, RGB(0, 255, 0)
                    SetPixel frmMain.Minimap.hdc, 0 + CharList(loopc).Pos.X, 1 + CharList(loopc).Pos.Y, RGB(0, 255, 0)
                End If
            End If
        Next loopc
    End If
   
    frmMain.Minimap.Refresh
End Sub
Public Sub ActualizaMinimap()
    frmMain.UserArea.Left = UserPos.X - 9
    frmMain.UserArea.Top = UserPos.Y - 8
End Sub
'***********************************************************
