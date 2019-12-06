Attribute VB_Name = "mDx8_Engine"
'MOTOR GRÁFICO ESCRITO(mayormente) POR MENDUZ@NOICODER.COM
Option Explicit

Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public FPS                     As Integer
Public FramesPerSecCounter    As Integer
Public timerElapsedTime       As Single
Public timerTicksPerFrame     As Double

Public engineBaseSpeed         As Single

Private lFrameTimer            As Long

Private HalfWindowTileWidth    As Integer
Private HalfWindowTileHeight   As Integer

Private bump_map_supported As Boolean

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
        Call QueryPerformanceFrequency(timer_freq)
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
    
    With MakeVector
        .x = x
        .Y = Y
        .Z = Z
    End With
    
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
    
    Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispModeBK)
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.renderer.ScaleWidth
        .BackBufferHeight = frmMain.renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.renderer.hWnd
    End With

    DispMode.Format = D3DFMT_X8R8G8B8

    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then

        Dim Caps8 As D3DCAPS8

        Call D3D.GetDeviceCaps(0, D3DDEVTYPE_HAL, Caps8)

        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.Format = DispModeBK.Format
        End If

    Else
        bump_map_supported = False
        DispMode.Format = DispModeBK.Format
    End If

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
                                                            
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
    
    With D3DDevice
    
        Call .SetVertexShader(D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR)
    
        '//Transformed and lit vertices dont need lighting
        '   so we disable it...
        Call .SetRenderState(D3DRS_LIGHTING, False)
        
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        
        'Partículas
        Call .SetRenderState(D3DRS_POINTSIZE, Engine_FToDW(2))
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        Call .SetRenderState(D3DRS_POINTSPRITE_ENABLE, 1)
        Call .SetRenderState(D3DRS_POINTSCALE_ENABLE, 0)
    
    End With
    
    Call SurfaceDB.Init(D3DX, D3DDevice)

    engineBaseSpeed = 0.017
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    Exit Sub
    
ErrHandler:
    Debug.Print "Error: " & Err.Number
    End
    
End Sub

Public Sub Engine_DeInitialize()

    Erase MapData
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
    
    With D3DDevice
    
        Call .BeginScene
        Call .Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    
        Call RenderScreen(50, 50)
        Call Engine_ActFPS
    
        With frmMain.Label1
            .Caption = "FPS: " & FPS
            .Refresh
        End With
    
        Call .EndScene
        Call .Present(ByVal 0, ByVal 0, 0, ByVal 0)
    
    End With

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
                          ByVal desthDC As Long, _
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
    
    With GrhData(grh_index)
    
        'If it's animated switch grh_index to first frame
        If .NumFrames <> 1 Then
            grh_index = .Frames(1)
        End If

        file_path = DirGraficos & .FileNum & ".bmp"
        
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
                Call SetPixel(MaskDC, x, Y, vbWhite)
                Call SetPixel(hdcsrc, x, Y, vbBlack)
            Else
                Call SetPixel(MaskDC, x, Y, vbBlack)
            End If

        Next x
    Next Y

End Sub

Public Sub Grh_Render(ByRef Grh As Grh, _
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

    Dim grh_index As Long
    
    With Grh
    
        If .grhindex = 0 Then Exit Sub
       
        'Animation
        If .Started = 1 Then
            .FrameCounter = .FrameCounter + (timerElapsedTime * GrhData(.grhindex).NumFrames / .speed)

            If .FrameCounter > GrhData(.grhindex).NumFrames Then
                .FrameCounter = (.FrameCounter Mod GrhData(.grhindex).NumFrames) + 1

                If .Loops <> -1 Then
                    
                    If Grh.Loops > 0 Then
                        .Loops = Grh.Loops - 1
                    Else
                        .Started = 0
                    End If

                End If

            End If

        End If
 
        'Figure out what frame to draw (always 1 if not animated)
        If .FrameCounter = 0 Then .FrameCounter = 1
        
        grh_index = GrhData(.grhindex).Frames(.FrameCounter)

        If grh_index <= 0 Then Exit Sub
        If GrhData(grh_index).FileNum = 0 Then Exit Sub
       
        'Modified by Augusto José Rando
        'Simplier function - according to basic ORE engine
        If h_centered Then
        
            If GrhData(.grhindex).TileWidth <> 1 Then
                screen_x = screen_x - Int(GrhData(.grhindex).TileWidth * (32 \ 2)) + 32 \ 2
            End If

        End If
   
        If v_centered Then
        
            If GrhData(.grhindex).TileHeight <> 1 Then
                screen_y = screen_y - Int(GrhData(.grhindex).TileHeight * 32) + 32
            End If

        End If
   
        'Draw it to device
        Call Device_Box_Textured_Render(grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list(), GrhData(grh_index).SX, GrhData(grh_index).SY, alpha_blend, angle)
    
    End With
    
End Sub

Public Sub Convert_Heading_to_Direction(ByVal Heading As Long, _
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
