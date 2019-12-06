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

'Tipo de las celdas del mapa
Public Type MapBlock
    particle_group As Integer
End Type

Public FPS                     As Long

' ARRAYS GLOBALES
Public GrhData()               As GrhData 'Guarda todos los grh
Public MapData()               As MapBlock ' Mapa

Private Declare Function SetPixel _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long) As Long

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
