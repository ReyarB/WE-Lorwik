Attribute VB_Name = "modIndices"
Option Explicit

Public GrhCount As Long

''
' Carga los indices de Graficos
'

Public Function LoadGrhData() As Boolean
On Error GoTo ErrorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim Handle As Integer
    Dim fileVersion As Long
    Dim file As String
    
    file = Get_Extract(Scripts, "Graficos.ind")
    
    'Open files
    Handle = FreeFile()
    Open file For Binary Access Read As Handle
    Seek #1, 1
    
    'Get file version
    Get Handle, , fileVersion
    
    'Get number of grhs
    Get Handle, , GrhCount
    
    'Resize arrays
    ReDim GrhData(1 To GrhCount) As GrhData
    
    While Not EOF(Handle)
        Get Handle, , Grh
        If Grh <> 0 Then
            With GrhData(Grh)
                'Get number of frames
                Get Handle, , .NumFrames
                If .NumFrames <= 0 Then GoTo ErrorHandler
                
                .Active = True
                
                ReDim .Frames(1 To GrhData(Grh).NumFrames)
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Get Handle, , .Frames(Frame)
                        If .Frames(Frame) <= 0 Or .Frames(Frame) > GrhCount Then
                            GoTo ErrorHandler
                        End If
                    Next Frame
                    
                    Get Handle, , .Speed
                    
                    If .Speed <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .pixelHeight = GrhData(.Frames(1)).pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    .pixelWidth = GrhData(.Frames(1)).pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileWidth = GrhData(.Frames(1)).TileWidth
                    If .TileWidth <= 0 Then GoTo ErrorHandler
                    
                    .TileHeight = GrhData(.Frames(1)).TileHeight
                    If .TileHeight <= 0 Then GoTo ErrorHandler
                Else
                    'Read in normal GRH data
                    Get Handle, , .FileNum
                    If .FileNum <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , GrhData(Grh).sX
                    If .sX < 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .sY
                    If .sY < 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .pixelWidth
                    If .pixelWidth <= 0 Then GoTo ErrorHandler
                    
                    Get Handle, , .pixelHeight
                    If .pixelHeight <= 0 Then GoTo ErrorHandler
                    
                    'Compute width and height
                    .TileWidth = .pixelWidth / TilePixelHeight
                    .TileHeight = .pixelHeight / TilePixelWidth
                    
                    .Frames(1) = Grh
                End If
            End With
        End If
    Wend
    
    Close Handle
   Delete_File file

    LoadGrhData = True
Exit Function
 
ErrorHandler:
    LoadGrhData = False
    Debug.Print "Error en LoadGrhData... Grh: " & Grh
End Function


''
' Carga los indices de Superficie
'

Public Sub CargarIndicesSuperficie()
'*************************************************
'Author: ^[GS]^
'Last modified: 29/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(IniPath & "Datos\indices.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Datos\indices.ini'", vbCritical
        End
    End If
    Dim Leer As New clsIniReader
    Dim i As Integer
    Leer.Initialize IniPath & "Datos\indices.ini"
    MaxSup = Leer.GetValue("INIT", "Referencias")
    ReDim SupData(MaxSup) As SupData
    frmMain.lListado(0).Clear
    For i = 0 To MaxSup
        SupData(i).name = Leer.GetValue("REFERENCIA" & i, "Nombre")
        SupData(i).Grh = Val(Leer.GetValue("REFERENCIA" & i, "GrhIndice"))
        SupData(i).Width = Val(Leer.GetValue("REFERENCIA" & i, "Ancho"))
        SupData(i).Height = Val(Leer.GetValue("REFERENCIA" & i, "Alto"))
        SupData(i).Block = IIf(Val(Leer.GetValue("REFERENCIA" & i, "Bloquear")) = 1, True, False)
        SupData(i).Capa = Val(Leer.GetValue("REFERENCIA" & i, "Capa"))
        frmMain.lListado(0).AddItem SupData(i).name & " - #" & i
    Next
    
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el indice " & i & " de Datos\indices.ini" & vbCrLf & "Err: " & err.Number & " - " & err.Description, vbCritical + vbOKOnly
End Sub

Public Sub CargarBloqueables()
    If FileExist(IniPath & "Datos\Bloqueables.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Datos\Bloqueables.ini'", vbCritical
        End
    End If
    
    Dim Leer As New clsIniReader
    Dim i As Integer
    
    Leer.Initialize IniPath & "Datos\Bloqueables.ini"
    MaxBloqueables = Leer.GetValue("INIT", "Num")
    
    ReDim Bloqueables(MaxBloqueables) As Long
    
    
    For i = 1 To MaxBloqueables
        Bloqueables(i) = Leer.GetValue("BLOQUEABLES", i)
    Next i
End Sub

''
' Carga los indices de Objetos
'

Public Sub CargarIndicesOBJ()
'*************************************************
'Author: ^[GS]^
'Last modified: 20/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(DirDats & "\OBJ.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'OBJ.dat' en " & DirDats, vbCritical
        End
    End If
    Dim Obj As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(DirDats & "\OBJ.dat")
    frmMain.lListado(3).Clear
    NumOBJs = Val(Leer.GetValue("INIT", "NumOBJs"))
    ReDim ObjData(1 To NumOBJs) As ObjData
    For Obj = 1 To NumOBJs
        frmCargando.X.Caption = "Cargando Datos de Objetos..." & Obj & "/" & NumOBJs
        DoEvents
        ObjData(Obj).name = Leer.GetValue("OBJ" & Obj, "Name")
        ObjData(Obj).GrhIndex = Val(Leer.GetValue("OBJ" & Obj, "GrhIndex"))
        ObjData(Obj).ObjType = Val(Leer.GetValue("OBJ" & Obj, "ObjType"))
        ObjData(Obj).Ropaje = Val(Leer.GetValue("OBJ" & Obj, "NumRopaje"))
        ObjData(Obj).Info = Leer.GetValue("OBJ" & Obj, "Info")
        ObjData(Obj).WeaponAnim = Val(Leer.GetValue("OBJ" & Obj, "Anim"))
        ObjData(Obj).Texto = Leer.GetValue("OBJ" & Obj, "Texto")
        ObjData(Obj).GrhSecundario = Val(Leer.GetValue("OBJ" & Obj, "GrhSec"))
        frmMain.lListado(3).AddItem ObjData(Obj).name & " - #" & Obj
    Next Obj
    Exit Sub
Fallo:
MsgBox "Error al intentar cargar el Objteto " & Obj & " de OBJ.dat en " & DirDats & vbCrLf & "Err: " & err.Number & " - " & err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Triggers
'

Public Sub CargarIndicesTriggers()
'*************************************************
'Author: ^[GS]^
'Last modified: 28/05/06
'*************************************************

On Error GoTo Fallo
    If FileExist(App.Path & "\Datos\Triggers.ini", vbArchive) = False Then
        MsgBox "Falta el archivo 'Triggers.ini' en \Datos\Triggers.ini", vbCritical
        End
    End If
    Dim NumT As Integer
    Dim T As Integer
    Dim Leer As New clsIniReader
    Call Leer.Initialize(App.Path & "\Datos\Triggers.ini")
    frmMain.lListado(4).Clear
    NumT = Val(Leer.GetValue("INIT", "NumTriggers"))
    For T = 1 To NumT
         frmMain.lListado(4).AddItem Leer.GetValue("Trig" & T, "Name") & " - #" & (T - 1)
    Next T

Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el Trigger " & T & " de Triggers.ini en " & App.Path & "\Datos\Triggers.ini" & vbCrLf & "Err: " & err.Number & " - " & err.Description, vbCritical + vbOKOnly

End Sub

''
' Carga los indices de Cuerpos
'

Public Sub CargarIndicesDeCuerpos()
    Dim n As Integer
    Dim i As Long
    Dim j As Byte
    Dim file As String
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    n = FreeFile()
    file = Get_Extract(Scripts, "Personajes.ind")
    Open file For Binary Access Read As #n
    
    'cabecera
    Get #n, , MiCabecera
    
    'num de cabezas
    Get #n, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As tBodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #n, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            For j = 1 To 4
                InitGrh BodyData(i).Walk(j), MisCuerpos(i).Body(j), 0
            Next j

            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #n
    
    Delete_File file
End Sub

''
' Carga los indices de Cabezas
'

Public Sub CargarIndicesDeCabezas()
    Dim i As Integer
    Dim file As String
    
        file = Get_Extract(Scripts, "Cabezas.ind")
        Open file For Binary Access Read As #1
    
            Get #1, , Numheads   'cantidad de cabezas
             
            ReDim heads(1 To Numheads) As tHead
             
            Dim Texture As Integer
            Dim temp    As Integer
            Dim startX  As Integer
            Dim startY  As Integer
            Dim skip As Byte
            
            For i = 1 To Numheads
                Get #1, , Texture 'number of .bmp
                Get #1, , startX
                Get #1, , startY
             
                heads(i).Texture = Texture
                heads(i).startX = startX
                heads(i).startY = startY
                
            Next i
         
    Close #1
    Delete_File file
End Sub

''
' Carga los indices de NPCs
'

Public Sub CargarIndicesNPC()
'*************************************************
'Author: ^[GS]^
'Last modified: 26/05/06
'*************************************************
On Error Resume Next
'On Error GoTo Fallo
    If FileExist(DirDats & "\NPCs.dat", vbArchive) = False Then
        MsgBox "Falta el archivo 'NPCs.dat' en " & DirDats, vbCritical
        End
    End If
    'If FileExist(DirDats & "\NPCs-HOSTILES.dat", vbArchive) = False Then
    '    MsgBox "Falta el archivo 'NPCs-HOSTILES.dat' en " & DirDats, vbCritical
    '    End
    'End If
    Dim Trabajando As String
    Dim NPC As Integer
    Dim Leer As New clsIniReader
    frmMain.lListado(1).Clear
    frmMain.lListado(2).Clear
    Call Leer.Initialize(DirDats & "\NPCs.dat")
    NumNPCs = Val(Leer.GetValue("INIT", "NumNPCs"))
    'Call Leer.Initialize(DirDats & "\NPCs-HOSTILES.dat")
    'NumNPCsHOST = Val(Leer.GetValue("INIT", "NumNPCs"))
    ReDim NpcData(1000) As NpcData
    Trabajando = "Dats\NPCs.dat"
    'Call Leer.Initialize(DirDats & "\NPCs.dat")
    'MsgBox "  "
    For NPC = 1 To NumNPCs
        NpcData(NPC).name = Leer.GetValue("NPC" & NPC, "Name")
        
        NpcData(NPC).Body = Val(Leer.GetValue("NPC" & NPC, "Body"))
        NpcData(NPC).Head = Val(Leer.GetValue("NPC" & NPC, "Head"))
        NpcData(NPC).Heading = Val(Leer.GetValue("NPC" & NPC, "Heading"))
        If LenB(NpcData(NPC).name) <> 0 Then frmMain.lListado(1).AddItem NpcData(NPC).name & " - #" & NPC
    Next
    'MsgBox "  "
    'Trabajando = "Dats\NPCs-HOSTILES.dat"
    'Call Leer.Initialize(DirDats & "\NPCs-HOSTILES.dat")
    'For NPC = 1 To NumNPCsHOST
    '    NpcData(NPC + 499).name = Leer.GetValue("NPC" & (NPC + 499), "Name")
    '    NpcData(NPC + 499).Body = Val(Leer.GetValue("NPC" & (NPC + 499), "Body"))
    '    NpcData(NPC + 499).Head = Val(Leer.GetValue("NPC" & (NPC + 499), "Head"))
    '    NpcData(NPC + 499).Heading = Val(Leer.GetValue("NPC" & (NPC + 499), "Heading"))
    '    If LenB(NpcData(NPC + 499).name) <> 0 Then frmMain.lListado(2).AddItem NpcData(NPC + 499).name & " - #" & (NPC + 499)
    'Next NPC
    Exit Sub
Fallo:
    MsgBox "Error al intentar cargar el NPC " & NPC & " de " & Trabajando & " en " & DirDats & vbCrLf & "Err: " & err.Number & " - " & err.Description, vbCritical + vbOKOnly

End Sub

Public Sub LoadMiniMap()
On Error GoTo ErrorHandler
    Dim file As String
    Dim count As Long
    Dim Handle As Integer
    
    'Open files
    Handle = FreeFile()
    
    file = Get_Extract(Scripts, "minimap.dat")
    Open file For Binary As Handle
        Seek Handle, 1
        For count = 1 To GrhCount
            If GrhData(count).Active Then
                Get Handle, , GrhData(count).MiniMap_color
            End If
        Next count
    Close Handle
    Delete_File file
    
ErrorHandler:
    Debug.Print "Error en LoadMiniMap."
End Sub

Public Sub CargarParticulas()
'*************************************
'Coded by OneZero (onezero_ss@hotmail.com)
'Last Modified: 6/4/03
'Loads the Particles.ini file to the ComboBox
'Edited by Juan Martín Sotuyo Dodero to add speed and life
'*************************************
    
On Error GoTo ErrorHandler
    
    Dim loopc As Long
    Dim i As Long
    Dim GrhListing As String
    Dim TempSet As String
    Dim ColorSet As Long
    Dim StreamFile As String
    Dim Leer As New clsIniReader
    
    If Not Extract_File(Scripts, App.Path & "\Recursos", "particulas.ini", Windows_Temp_Dir, False) Then
        err.Description = "¡No se puede cargar el archivo de recurso!"
        GoTo ErrorHandler
    End If
    
    StreamFile = Windows_Temp_Dir & "Particulas.ini"
    Leer.Initialize StreamFile
    
    TotalStreams = Val(Leer.GetValue("INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
        For loopc = 1 To TotalStreams
            StreamData(loopc).name = Leer.GetValue(Val(loopc), "Name")
            frmMain.lstParticle.AddItem loopc & "-" & StreamData(loopc).name
            StreamData(loopc).NumOfParticles = Leer.GetValue(Val(loopc), "NumOfParticles")
            StreamData(loopc).X1 = Leer.GetValue(Val(loopc), "X1")
            StreamData(loopc).Y1 = Leer.GetValue(Val(loopc), "Y1")
            StreamData(loopc).X2 = Leer.GetValue(Val(loopc), "X2")
            StreamData(loopc).Y2 = Leer.GetValue(Val(loopc), "Y2")
            StreamData(loopc).angle = Leer.GetValue(Val(loopc), "Angle")
            StreamData(loopc).vecx1 = Leer.GetValue(Val(loopc), "VecX1")
            StreamData(loopc).vecx2 = Leer.GetValue(Val(loopc), "VecX2")
            StreamData(loopc).vecy1 = Leer.GetValue(Val(loopc), "VecY1")
            StreamData(loopc).vecy2 = Leer.GetValue(Val(loopc), "VecY2")
            StreamData(loopc).life1 = Leer.GetValue(Val(loopc), "Life1")
            StreamData(loopc).life2 = Leer.GetValue(Val(loopc), "Life2")
            StreamData(loopc).friction = Leer.GetValue(Val(loopc), "Friction")
            StreamData(loopc).spin = Leer.GetValue(Val(loopc), "Spin")
            StreamData(loopc).spin_speedL = Leer.GetValue(Val(loopc), "Spin_SpeedL")
            StreamData(loopc).spin_speedH = Leer.GetValue(Val(loopc), "Spin_SpeedH")
            StreamData(loopc).AlphaBlend = Leer.GetValue(Val(loopc), "AlphaBlend")
            StreamData(loopc).gravity = Leer.GetValue(Val(loopc), "Gravity")
            StreamData(loopc).grav_strength = Leer.GetValue(Val(loopc), "Grav_Strength")
            StreamData(loopc).bounce_strength = Leer.GetValue(Val(loopc), "Bounce_Strength")
            StreamData(loopc).XMove = Leer.GetValue(Val(loopc), "XMove")
            StreamData(loopc).YMove = Leer.GetValue(Val(loopc), "YMove")
            StreamData(loopc).move_x1 = Leer.GetValue(Val(loopc), "move_x1")
            StreamData(loopc).move_x2 = Leer.GetValue(Val(loopc), "move_x2")
            StreamData(loopc).move_y1 = Leer.GetValue(Val(loopc), "move_y1")
            StreamData(loopc).move_y2 = Leer.GetValue(Val(loopc), "move_y2")
            StreamData(loopc).Radio = Val(Leer.GetValue(Val(loopc), "Radio"))
            StreamData(loopc).life_counter = Leer.GetValue(Val(loopc), "life_counter")
            StreamData(loopc).Speed = Val(Leer.GetValue(Val(loopc), "Speed"))
            StreamData(loopc).NumGrhs = Leer.GetValue(Val(loopc), "NumGrhs")
           
            ReDim StreamData(loopc).grh_list(1 To StreamData(loopc).NumGrhs)
            GrhListing = Leer.GetValue(Val(loopc), "Grh_List")
           
            For i = 1 To StreamData(loopc).NumGrhs
                StreamData(loopc).grh_list(i) = ReadField(Str(i), GrhListing, 44)
            Next i
            StreamData(loopc).grh_list(i - 1) = StreamData(loopc).grh_list(i - 1)
            For ColorSet = 1 To 4
                TempSet = Leer.GetValue(Val(loopc), "ColorSet" & ColorSet)
                StreamData(loopc).colortint(ColorSet - 1).R = ReadField(1, TempSet, 44)
                StreamData(loopc).colortint(ColorSet - 1).G = ReadField(2, TempSet, 44)
                StreamData(loopc).colortint(ColorSet - 1).B = ReadField(3, TempSet, 44)
            Next ColorSet
        
    Next loopc
    
    Delete_File Windows_Temp_Dir & "particulas.ini"
    Set Leer = Nothing
    
Exit Sub
    
ErrorHandler:
    If FileExist(Windows_Temp_Dir & "particulas.ini", vbNormal) Then Delete_File Windows_Temp_Dir & "particulas.ini"
    
End Sub

