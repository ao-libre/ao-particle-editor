Attribute VB_Name = "Mod_TileEngine"
Option Explicit

'Map sizes in tiles
Public Const XMaxMapSize     As Byte = 100
Public Const XMinMapSize     As Byte = 1
Public Const YMaxMapSize     As Byte = 100
Public Const YMinMapSize     As Byte = 1

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1

'Posicion en un mapa
Public Type Position
    x As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    x As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData

    SX As Integer
    SY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    speed As Single
    
    active As Boolean

End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    grhindex As Integer
    FrameCounter As Single
    speed As Single
    Started As Byte
    Loops As Integer
    angle As Single

End Type

'Apariencia del personaje
Public Type Char

    active As Byte
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    
    Nombre As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    Muerto As Boolean
    invisible As Boolean
    priv As Byte

End Type

'Info de un objeto
Public Type Obj

    OBJIndex As Integer
    Amount As Integer

End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    particle_group As Integer

    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer

End Type

'Bordes del mapa
Public MinXBorder              As Byte
Public MaxXBorder              As Byte
Public MinYBorder              As Byte
Public MaxYBorder              As Byte

Public EngineRun               As Boolean

Public FPS                     As Long

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize          As Integer

' ARRAYS GLOBALES
Public GrhData()               As GrhData 'Guarda todos los grh
Public MapData()               As MapBlock ' Mapa
Public charlist(1 To 10000)    As Char

Private Declare Function SetPixel _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long) As Long

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Sub InitGrh(ByRef Grh As Grh, _
                   ByVal grhindex As Integer, _
                   Optional ByVal Started As Byte = 2)

    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    If grhindex = 0 Then Exit Sub
    
    With Grh
    
        .grhindex = grhindex
    
        If Started = 2 Then
            If GrhData(.grhindex).NumFrames > 1 Then
                .Started = 1
            Else
                .Started = 0
            End If

        Else

            'Make sure the graphic can be started
            If GrhData(Grh.grhindex).NumFrames = 1 Then Started = 0
            .Started = Started

        End If
    
        If .Started Then
            .Loops = INFINITE_LOOPS
        Else
            .Loops = 0
        End If
    
        .FrameCounter = 1
        .speed = GrhData(Grh.grhindex).speed
    
    End With

End Sub


Function InMapBounds(ByVal x As Integer, ByVal Y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************
    
    If x < XMinMapSize Or x > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True

End Function

Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, _
                            ByRef MaskDC As Long, _
                            ByVal src_x As Integer, _
                            ByVal src_y As Integer, _
                            ByVal src_width As Integer, _
                            ByVal src_height As Integer)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 8/30/2004
    'Creates a Mask hDC, and sets the source hDC to work for trans bliting.
    '**************************************************************
    Dim x          As Integer
    Dim Y          As Integer
    Dim TransColor As Long
    Dim ColorKey   As String

    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For Y = src_y To src_height + src_y
        For x = src_x To src_width + src_x

            If GetPixel(hdcsrc, x, Y) = TransColor Then
                Call SetPixel(MaskDC, x, Y, vbWhite)
                Call SetPixel(hdcsrc, x, Y, vbBlack)
            Else
                Call SetPixel(MaskDC, x, Y, vbBlack)
            End If

        Next x
    Next Y

End Sub
