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

'Lista de cuerpos
Public Type BodyData

    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position

End Type

'Lista de cabezas
Public Type HeadData

    Head(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de las armas
Type WeaponAnimData

    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData

    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh

End Type

'Apariencia del personaje
Public Type Char

    active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
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

    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    light_value(3) As Long
    
    luz As Integer
    Color(3) As Long
    
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

'Status del user
Public CurMap                  As Integer 'Mapa actual
Public UserIndex               As Integer
Public UserMoving              As Byte
Public UserBody                As Integer
Public UserHead                As Integer
Public UserPos                 As Position 'Posicion

Public AddtoUserPos            As Position 'Si se mueve
Public UserCharIndex           As Integer

Public EngineRun               As Boolean

Public FPS                     As Long
Public FramesPerSecCounter     As Long
Private fpsLastCheck           As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth        As Integer
Private WindowTileHeight       As Integer

Private HalfWindowTileWidth    As Integer
Private HalfWindowTileHeight   As Integer

'Offset del desde 0,0 del main view
Private MainViewTop            As Integer
Private MainViewLeft           As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize          As Integer
Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight         As Integer
Public TilePixelWidth          As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX   As Integer
Public ScrollPixelsPerFrameY   As Integer

Dim timerElapsedTime           As Single
Dim timerTicksPerFrame         As Single
Dim engineBaseSpeed            As Single

Public NumChars                As Integer
Public LastChar                As Integer

Private MainDestRect           As RECT
Private MainViewRect           As RECT
Private BackBufferRect         As RECT

Private MainViewWidth          As Integer
Private MainViewHeight         As Integer

' ARRAYS GLOBALES
Public GrhData()               As GrhData 'Guarda todos los grh
Public MapData()               As MapBlock ' Mapa
Public charlist(1 To 10000)    As Char


Private Declare Function StretchBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal nSrcWidth As Long, _
                             ByVal nSrcHeight As Long, _
                             ByVal dwRop As Long) As Long

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

Private Function GetElapsedTime() As Single

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim start_time    As Currency
    Static end_time   As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        Call QueryPerformanceFrequency(timer_freq)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

End Function

Public Sub Grh_Render_To_Hdc(ByVal desthDC As Long, _
                             ByVal grh_index As Long, _
                             ByVal screen_x As Integer, _
                             ByVal screen_y As Integer, _
                             Optional transparent As Boolean = False)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/30/2004
    'This method is SLOW... Don't use in a loop if you care about
    'speed!
    'Modified by Juan Martín Sotuyo Dodero
    '*************************************************************
    
    On Error GoTo ErrorHandler
    
    Dim file_path  As String

    Dim src_x      As Integer
    Dim src_y      As Integer

    Dim src_width  As Integer
    Dim src_height As Integer

    Dim hdcsrc     As Long
    Dim MaskDC     As Long
    Dim PrevObj    As Long
    Dim PrevObj2   As Long

    If grh_index <= 0 Then Exit Sub
    
    With GrhData(grh_index)
    
        'If it's animated switch grh_index to first frame
        If .NumFrames <> 1 Then
            grh_index = .Frames(1)
        End If
    
        file_path = App.Path & "\Graficos\" & .FileNum & ".bmp"
            
        src_x = .SX
        src_y = .SY
        src_width = .pixelWidth
        src_height = .pixelHeight
    
    End With
    
            
    hdcsrc = CreateCompatibleDC(desthDC)
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
        
    If transparent = False Then
        
        Call BitBlt(desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy)
    
    Else
    
        MaskDC = CreateCompatibleDC(desthDC)
        PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
            
        Call Grh_Create_Mask(hdcsrc, MaskDC, src_x, src_y, src_width, src_height)
            
        'Render tranparently
        Call BitBlt(desthDC, screen_x, screen_y, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd)
        Call BitBlt(desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcPaint)
            
        Call DeleteObject(SelectObject(MaskDC, PrevObj2))
            
        Call DeleteDC(MaskDC)

    End If
        
    Call DeleteObject(SelectObject(hdcsrc, PrevObj))
    Call DeleteDC(hdcsrc)
    
    Exit Sub
    
ErrorHandler:
    
End Sub

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
