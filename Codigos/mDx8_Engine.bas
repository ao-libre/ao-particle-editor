Attribute VB_Name = "mDx8_Engine"
'MOTOR GRÁFICO ESCRITO(mayormente) POR MENDUZ@NOICODER.COM
Option Explicit
 
Private Type decoration
    Grh As Grh
    Render_On_Top As Boolean
    subtile_pos As Byte
End Type

Dim base_tile_size As Integer

Public bRunning           As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Dim font_count      As Long
Dim font_last       As Long
Private font_list() As D3DXFont

Dim texture      As Direct3DTexture8
Dim TransTexture As Direct3DTexture8

Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public FPS                     As Integer
Private FramesPerSecCounter    As Integer
Private timerElapsedTime       As Single
Private timerTicksPerFrame     As Double

Private particletimer          As Single

Public engineBaseSpeed         As Single

Private lFrameTimer            As Long
Private lFrameLimiter          As Long

Private ScrollPixelsPerFrameX  As Byte
Private ScrollPixelsPerFrameY  As Byte

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer

Private MainViewTop            As Integer
Private MainViewLeft           As Integer

Private MainDestRect           As RECT
Private MainViewRect           As RECT
Private BackBufferRect         As RECT

Private MainViewWidth          As Integer
Private MainViewHeight         As Integer

Private MouseTileX             As Byte
Private MouseTileY             As Byte

Private iFrameIndex            As Byte  'Frame actual de la LL

Private llTick                 As Long  'Contador

Private WindowTileWidth        As Integer
Private WindowTileHeight       As Integer

Private HalfWindowTileWidth    As Integer
Private HalfWindowTileHeight   As Integer

Dim dimeTex             As Long

Dim tex                 As Direct3DTexture8
Dim D3DbackBuffer       As Direct3DSurface8
Dim zTarget             As Direct3DSurface8
Dim stencil             As Direct3DSurface8
Dim superTex            As Direct3DSurface8
Dim bump_map_texture    As Direct3DTexture8
Dim bump_map_texture_ex As Direct3DTexture8
Dim bump_map_supported  As Boolean
Dim bump_map_powa       As Boolean

Dim char_last           As Long
Dim char_list()         As Char
Dim char_count          As Long

'Sets a Grh animation to loop indefinitely.
Private Declare Function BitBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal x As Long, _
                             ByVal Y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal dwRop As Long) As Long

Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal hObject As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

#Const HARDCODED = False 'True ' == MÁS FPS ^^

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
        QueryPerformanceFrequency timer_freq

    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

End Function

Function MakeVector(ByVal x As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    MakeVector.x = x
    MakeVector.Y = Y
    MakeVector.Z = Z

End Function

Public Sub Engine_Init()
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    'On Error GoTo ErrHandler:

    Dim DispMode    As D3DDISPLAYMODE
    Dim DispModeBK  As D3DDISPLAYMODE
    Dim D3DWindow   As D3DPRESENT_PARAMETERS

    Set SurfaceDB = New clsTexManager
    
    Set DirectX = New DirectX8
    Set D3D = DirectX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.format
        .BackBufferWidth = frmMain.renderer.ScaleWidth
        .BackBufferHeight = frmMain.renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.renderer.hWnd

    End With

    DispMode.format = D3DFMT_X8R8G8B8

    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then

        Dim Caps8 As D3DCAPS8

        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8

        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.format = DispModeBK.format
        End If

    Else
        bump_map_supported = False
        DispMode.format = DispModeBK.format
    End If

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
                                                            
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
    
    TileBufferSize = 9
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    engineBaseSpeed = 0.017
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    ScrollPixelsPerFrameX = 8
    ScrollPixelsPerFrameY = 8
    
    UserPos.x = 50
    UserPos.Y = 50
    
    MinXBorder = XMinMapSize
    MaxXBorder = XMaxMapSize
    MinYBorder = YMinMapSize
    MaxYBorder = YMaxMapSize
    
    'partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    bRunning = True
    Exit Sub
ErrHandler:
    Debug.Print "Error Number Returned: " & Err.Number
    bRunning = False

End Sub

Public Sub Engine_DeInitialize()

    Erase MapData
    Erase charlist
    Erase particle_group_list
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set DirectX = Nothing
    End

End Sub
 
Public Sub Engine_ActFPS()

    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If

End Sub

Public Sub Render()
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************

    Call D3DDevice.BeginScene
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    
    Call RenderScreen(50, 50)
    Call Engine_ActFPS
    
    With frmMain.Label1
        .Caption = "FPS: " & FPS
        .Refresh
    End With
    
    Call D3DDevice.EndScene
    Call D3DDevice.Present(ByVal 0, ByVal 0, 0, ByVal 0)

    lFrameLimiter = GetTickCount
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************
    Dim Y                 As Integer     'Keeps track of where on map we are
    Dim x                 As Integer     'Keeps track of where on map we are

    Dim screenminY        As Integer  'Start Y pos on current screen
    Dim screenmaxY        As Integer  'End Y pos on current screen

    Dim screenminX        As Integer  'Start X pos on current screen
    Dim screenmaxX        As Integer  'End X pos on current screen

    Dim ScreenX           As Integer  'Keeps track of where to place tile on screen
    Dim ScreenY           As Integer  'Keeps track of where to place tile on screen

    Static OffsetCounterX As Single
    Static OffsetCounterY As Single

    Dim PixelOffsetX      As Integer
    Dim PixelOffsetY      As Integer
    
    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    PixelOffsetX = OffsetCounterX
    PixelOffsetY = OffsetCounterY

    For Y = screenminY To screenmaxY
        For x = screenminX To screenmaxX

            With MapData(x, Y)

                '***********************************************
                If .particle_group > 0 Then
                    Call Particle_Group_Render(.particle_group, (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY)
                End If

            End With

            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next Y
        
End Sub

Private Function Geometry_Create_TLVertex(ByVal x As Single, _
                                          ByVal Y As Single, _
                                          ByVal Z As Single, _
                                          ByVal rhw As Single, _
                                          ByVal Color As Long, _
                                          ByVal Specular As Long, _
                                          tu As Single, _
                                          ByVal tv As Single) As TLVERTEX

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '**************************************************************
    
    With Geometry_Create_TLVertex
        .x = x
        .Y = Y
        .Z = Z
        .rhw = rhw
        .Color = Color
        .Specular = Specular
        .tu = tu
        .tv = tv
    End With

End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, _
                                ByRef dest As RECT, _
                                ByRef src As RECT, _
                                ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, _
                                Optional ByRef Textures_Height As Long, _
                                Optional ByVal angle As Single)

    '**************************************************************
    'Author: Aaron Perkins
    'Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 11/17/2002
    '
    ' * v1      * v3
    ' |\        |
    ' |  \      |
    ' |    \    |
    ' |      \  |
    ' |        \|
    ' * v0      * v2
    '**************************************************************
    Dim x_center    As Single
    Dim y_center    As Single

    Dim radius      As Single

    Dim x_Cor       As Single
    Dim y_Cor       As Single

    Dim left_point  As Single
    Dim right_point As Single

    Dim temp        As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal grhindex As Long, _
                                      ByVal dest_x As Integer, _
                                      ByVal dest_y As Integer, _
                                      ByVal src_width As Integer, _
                                      ByVal src_height As Integer, _
                                      ByRef rgb_list() As Long, _
                                      ByVal src_x As Integer, _
                                      ByVal src_y As Integer, _
                                      Optional ByVal alpha_blend As Boolean, _
                                      Optional ByVal angle As Single)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 2/12/2004
    'Just copies the Textures
    '**************************************************************
    Static src_rect            As RECT
    Static dest_rect           As RECT

    Static temp_verts(3)       As TLVERTEX

    Static D3DTextures         As D3D8Textures

    Static light_value(0 To 3) As Long
    
    If grhindex = 0 Then Exit Sub
    
    Set D3DTextures.texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, D3DTextures.texwidth, D3DTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)

    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Call Geometry_Create_Box(temp_verts(), dest_rect, src_rect, light_value(), D3DTextures.texwidth, D3DTextures.texheight, angle)
    
    With D3DDevice
    
        'Set Textures
        Call .SetTexture(0, D3DTextures.texture)
    
        If alpha_blend Then
            'Set Rendering for alphablending
            Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_ONE)
            Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_ONE)

            'Sets up properties for transparency.
            Call D3DDevice.SetRenderState(D3DRS_ALPHAREF, 255)
            Call .SetRenderState(D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL)
        End If

        'Draw the triangles that make up our square Textures
        Call .DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0)))
    
        If alpha_blend Then
            'Set Rendering for colokeying
            Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
            Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
            Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
        End If
    
    End With

End Sub

Public Sub Start()

    Dim f              As Boolean
    Dim ulttick        As Long, esttick As Long
    Dim timers(1 To 2) As Integer
    Dim LoopC          As Integer
    
    With vertList(0)
        .x = 0
        .Y = 0
        .rhw = 1
        .Color = D3DColorXRGB(255, 255, 255)
        .tu = 0
        .tv = 0
    End With
    
    With vertList(1)
        .x = 800
        .Y = 0
        .rhw = 1
        .Color = D3DColorXRGB(255, 255, 255)
        .tu = 1
        .tv = 0
    End With
    
    With vertList(2)
        .x = 0
        .Y = 600
        .rhw = 1
        .Color = D3DColorXRGB(255, 255, 255)
        .tu = 0
        .tv = 1
    End With
    
    With vertList(3)
        .x = 800
        .Y = 600
        .rhw = 1
        .Color = D3DColorXRGB(255, 255, 255)
        .tu = 1
        .tv = 1
    End With
    
    On Error Resume Next

    Do While prgRun

        If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
            Call Render
        Else
            Call Sleep(10&)
        End If

        DoEvents

    Loop
    
    Call Engine_DeInitialize

    EngineRun = False
    frmCargando.Show

    'Destruimos los objetos públicos creados
    Set SurfaceDB = Nothing
    
    Call UnloadAllForms
    
    End

End Sub

Private Function Engine_FToDW(f As Single) As Long

    ' single > long
    Dim buf As D3DXBuffer

    Set buf = D3DX.CreateBuffer(4)
    Call D3DX.BufferSetData(buf, 0, 4, 1, f)
    Call D3DX.BufferGetData(buf, 0, 4, 1, Engine_FToDW)

End Function

Public Sub GrhRenderToHdc(ByVal grh_index As Long, _
                          desthDC As Long, _
                          ByVal screen_x As Integer, _
                          ByVal screen_y As Integer, _
                          Optional transparent As Boolean = False)

    On Error Resume Next

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

    'If it's animated switch grh_index to first frame
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)

    End If

    file_path = DirGraficos & GrhData(grh_index).FileNum & ".bmp"
        
    src_x = GrhData(grh_index).SX
    src_y = GrhData(grh_index).SY
    src_width = GrhData(grh_index).pixelWidth
    src_height = GrhData(grh_index).pixelHeight
            
    hdcsrc = CreateCompatibleDC(desthDC)
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
        
    If transparent = False Then
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
    Else
        MaskDC = CreateCompatibleDC(desthDC)
            
        PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
            
        Grh_Create_Mask hdcsrc, MaskDC, src_x, src_y, src_width, src_height
            
        'Render tranparently
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcPaint
            
        Call DeleteObject(SelectObject(MaskDC, PrevObj2))
            
        DeleteDC MaskDC

    End If
        
    Call DeleteObject(SelectObject(hdcsrc, PrevObj))
    DeleteDC hdcsrc

    Exit Sub

End Sub

Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, _
                            ByRef MaskDC As Long, _
                            ByVal src_x As Integer, _
                            ByVal src_y As Integer, _
                            ByVal src_width As Integer, _
                            ByVal src_height As Integer)

    Dim x          As Integer

    Dim Y          As Integer

    Dim TransColor As Long

    Dim ColorKey   As String

    ColorKey = "0"
    TransColor = &H0

    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For Y = src_y To src_height + src_y
        For x = src_x To src_width + src_x

            If GetPixel(hdcsrc, x, Y) = TransColor Then
                SetPixel MaskDC, x, Y, vbWhite
                SetPixel hdcsrc, x, Y, vbBlack
            Else
                SetPixel MaskDC, x, Y, vbBlack

            End If

        Next x
    Next Y

End Sub

Private Sub Grh_Render(ByRef Grh As Grh, _
                       ByVal screen_x As Integer, _
                       ByVal screen_y As Integer, _
                       ByRef rgb_list() As Long, _
                       Optional ByVal h_centered As Boolean = True, _
                       Optional ByVal v_centered As Boolean = True, _
                       Optional ByVal alpha_blend As Boolean = False, _
                       Optional ByVal KillAnim As Boolean = 0, _
                       Optional angle As Single)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/28/2003
    'Modified by Juan Martín Sotuyo Dodero
    'Added centering
    '**************************************************************
    On Error Resume Next

    Dim tile_width  As Integer
    Dim tile_height As Integer

    Dim grh_index   As Long
   
    If Grh.grhindex = 0 Then Exit Sub
       
    'Animation
    If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhindex).NumFrames / Grh.speed)

        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhindex).NumFrames) + 1

            If Grh.Loops <> -1 Then
                If Grh.Loops > 0 Then
                    Grh.Loops = Grh.Loops - 1
                Else
                    Grh.Started = 0

                End If

            End If

        End If

    End If
 
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    'If Not Grh_Check(Grh.grhindex) Then Exit Sub
    grh_index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)

    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
       
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(Grh.grhindex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * (32 \ 2)) + 32 \ 2

        End If

    End If
   
    If v_centered Then
        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * 32) + 32

        End If

    End If
   
    'Draw it to device
    Device_Box_Textured_Render grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list(), GrhData(grh_index).SX, GrhData(grh_index).SY, alpha_blend, angle
 
End Sub

Private Sub Convert_Heading_to_Direction(ByVal Heading As Long, _
                                         ByRef direction_x As Long, _
                                         ByRef direction_y As Long)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    Dim addy As Long

    Dim addx As Long
    
    'Figure out which way to move
    Select Case Heading
    
        Case 1
            addy = -1
    
        Case 2
            addy = -1
            addx = 1
    
        Case 3
            addx = 1
            
        Case 4
            addx = 1
            addy = 1
    
        Case 5
            addy = 1
        
        Case 6
            addx = -1
            addy = 1
        
        Case 7
            addx = -1
            
        Case 8
            addx = -1
            addy = -1
            
    End Select
    
    direction_x = direction_x + addx
    direction_y = direction_y + addy

End Sub

Private Function Particle_Group_Next_Open() As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim LoopC As Long
    
    LoopC = 1

    Do Until particle_group_list(LoopC).active = False

        If LoopC = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function

        End If

        LoopC = LoopC + 1
    Loop
    
    Particle_Group_Next_Open = LoopC
    Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1

End Function

Public Function Particle_Group_Create(ByVal map_x As Integer, _
                                      ByVal map_y As Integer, _
                                      ByRef grh_index_list() As Long, _
                                      ByRef rgb_list() As Long, _
                                      Optional ByVal particle_count As Long = 20, _
                                      Optional ByVal stream_type As Long = 1, _
                                      Optional ByVal alpha_blend As Boolean, _
                                      Optional ByVal alive_counter As Long = -1, _
                                      Optional ByVal frame_speed As Single = 0.5, _
                                      Optional ByVal id As Long, _
                                      Optional ByVal x1 As Integer, _
                                      Optional ByVal y1 As Integer, _
                                      Optional ByVal angle As Integer, _
                                      Optional ByVal vecx1 As Integer, _
                                      Optional ByVal vecx2 As Integer, _
                                      Optional ByVal vecy1 As Integer, _
                                      Optional ByVal vecy2 As Integer, _
                                      Optional ByVal life1 As Integer, _
                                      Optional ByVal life2 As Integer, _
                                      Optional ByVal fric As Integer, _
                                      Optional ByVal spin_speedL As Single, _
                                      Optional ByVal gravity As Boolean, _
                                      Optional grav_strength As Long, _
                                      Optional bounce_strength As Long, _
                                      Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional ByVal Radio As Integer) As Long
    
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 12/15/2002
    'Returns the particle_group_index if successful, else 0
    '**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey, Radio
        Else
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey, Radio

        End If

    End If

End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True

    End If

End Function

Public Function Particle_Group_Remove_All() As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    Dim index As Long
    
    For index = 1 To particle_group_last

        'Make sure it's a legal index
        If Particle_Group_Check(index) Then
            Particle_Group_Destroy index

        End If

    Next index
    
    Particle_Group_Remove_All = True

End Function

Public Function Particle_Group_Find(ByVal id As Long) As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Find the index related to the handle
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim LoopC As Long
    
    LoopC = 1

    Do Until particle_group_list(LoopC).id = id

        If LoopC = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function

        End If

        LoopC = LoopC + 1
    Loop
    
    Particle_Group_Find = LoopC
    Exit Function
ErrorHandler:
    Particle_Group_Find = 0

End Function

Private Sub Particle_Group_Make(ByVal particle_group_index As Long, _
                                ByVal map_x As Integer, _
                                ByVal map_y As Integer, _
                                ByVal particle_count As Long, _
                                ByVal stream_type As Long, _
                                ByRef grh_index_list() As Long, _
                                ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, _
                                Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, _
                                Optional ByVal id As Long, _
                                Optional ByVal x1 As Integer, _
                                Optional ByVal y1 As Integer, _
                                Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, _
                                Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, _
                                Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, _
                                Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, _
                                Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, _
                                Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional Radio As Integer)

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Makes a new particle effect
    '*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)

    End If

    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y

    End If
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    Rem Lord Fers
    particle_group_list(particle_group_index).Radio = Radio
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False

    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on map
    MapData(map_x, map_y).particle_group = particle_group_index

End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, _
                            ByVal screen_x As Integer, _
                            ByVal screen_y As Integer, _
                            ByVal grh_index As Long, _
                            ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, _
                            Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, _
                            Optional ByVal y1 As Integer, _
                            Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, _
                            Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, _
                            Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, _
                            Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, _
                            Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, _
                            Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, _
                            Optional ByVal x2 As Integer, _
                            Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, _
                            Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional ByVal Radio As Integer, Optional ByVal count As Integer, Optional ByVal index As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 4/24/2003
    '
    '**************************************************************

    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            InitGrh temp_particle.Grh, grh_index, alpha_blend

            If Radio = 0 Then
                temp_particle.x = RandomNumber(x1, x2)
                temp_particle.Y = RandomNumber(y1, y2)
            Else
                temp_particle.x = (RandomNumber(x1, x2) + Radio) + Radio * Cos(PI * 2 * index / count)
                temp_particle.Y = (RandomNumber(y1, y2) + Radio) + Radio * Sin(PI * 2 * index / count)

            End If

            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
        Else

            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength

                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength

                End If

            End If

            'Do rotation
            If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0

            End If
                                
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)

        End If

        'Add in vector
        temp_particle.x = temp_particle.x + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
        temp_particle.alive_counter = temp_particle.alive_counter - 1

    End If
 
    'Draw it
    Grh_Render temp_particle.Grh, temp_particle.x + screen_x, temp_particle.Y + screen_y, rgb_list(), 1, False, True, True, temp_particle.Grh.angle

End Sub

Private Sub Particle_Group_Render(ByVal particle_group_index As Long, _
                                  ByVal screen_x As Integer, _
                                  ByVal screen_y As Integer)

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 12/15/2002
    'Renders a particle stream at a paticular screen point
    '*****************************************************************
    Dim LoopC            As Long

    Dim temp_rgb(0 To 3) As Long

    Dim no_move          As Boolean
    
    'Set colors
    '   If UserMinHP = 0 Then
    '       temp_rgb(0) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    '      temp_rgb(1) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    '       temp_rgb(2) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    '       temp_rgb(3) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    '   Else
    temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
    temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
    temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
    temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    '  End If
        
    If particle_group_list(particle_group_index).alive_counter Then
    
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame

        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True

        End If
    
        'If it's still alive render all the particles inside
        For LoopC = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(LoopC), _
               screen_x, screen_y, _
               particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
               temp_rgb(), _
               particle_group_list(particle_group_index).alpha_blend, no_move, _
               particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).angle, _
               particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
               particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
               particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
               particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
               particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
               particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
               particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
               particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
               particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
               particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
               particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey, _
               particle_group_list(particle_group_index).Radio, particle_group_list(particle_group_index).particle_count, LoopC
                            
        Next LoopC
        
        If no_move = False Then

            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1

            End If

        End If
    
    Else
        'If it's dead destroy it
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1

        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index

    End If

End Sub

Public Function Particle_Type_Get(ByVal particle_index As Long) As Long

    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modify Date: 8/27/2003
    'Returns the stream type of a particle stream
    '*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type

    End If

End Function

Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True

        End If

    End If

End Function

Public Function Particle_Group_Map_Pos_Set(ByVal particle_group_index As Long, _
                                           ByVal map_x As Long, _
                                           ByVal map_y As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 5/27/2003
    'Returns true if successful, else false
    '**************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then

        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
            'Move it
            particle_group_list(particle_group_index).map_x = map_x
            particle_group_list(particle_group_index).map_y = map_y
    
            Particle_Group_Map_Pos_Set = True

        End If

    End If

End Function

Public Function Particle_Group_Move(ByVal particle_group_index As Long, _
                                    ByVal Heading As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 5/27/2003
    'Returns true if successful, else false
    '**************************************************************
    Dim map_x As Long

    Dim map_y As Long

    Dim nX    As Long

    Dim nY    As Long
    
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Particle_Group_Move = False
        Exit Function

    End If
    
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
    
        map_x = particle_group_list(particle_group_index).map_x
        map_y = particle_group_list(particle_group_index).map_y
        
        nX = map_x
        nY = map_y
        
        Convert_Heading_to_Direction Heading, nX, nY
        
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
            'Move it
            particle_group_list(particle_group_index).map_x = nX
            particle_group_list(particle_group_index).map_y = nY
            
            Particle_Group_Move = True

        End If

    End If

End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    Dim temp As particle_group
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group = 0

    End If
    
    particle_group_list(particle_group_index) = temp
            
    'Update array size
    If particle_group_index = particle_group_last Then

        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1

            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub

            End If

        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)

    End If

    particle_group_count = particle_group_count - 1

End Sub

Public Function Map_Particle_Group_Get(ByVal map_x As Integer, _
                                       ByVal map_y As Integer) As Long

    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group
    Else
        Map_Particle_Group_Get = 0

    End If

End Function

