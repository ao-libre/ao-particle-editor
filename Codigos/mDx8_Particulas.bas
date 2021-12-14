Attribute VB_Name = "mDx8_Particulas"
Option Explicit

'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream

Public Type Stream

    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    speed As Single
    life_counter As Long
    
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
    
    Radio As Integer

End Type

Public Type Particle
    friction As Single
    x As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    Radio As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
 
Public Type particle_group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    live As Long
    liv1 As Integer
    liveend As Long
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    
    'Added by Juan Martín Sotuyo Dodero
    speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
    Radio As Integer
End Type

'Particle system
Public particle_group_list() As particle_group
Public particle_group_count  As Long
Public particle_group_last   As Long
 
Public Sub LoadStreamFile(ByVal StreamFile As String)

    On Error GoTo Error

    Dim LoopC As Long
    Dim i          As Long
    Dim GrhListing As String
    
    '****************************
    'Load stream file via clsIniManager
    '****************************
    Dim FileManager As clsIniManager
    Set FileManager = New clsIniManager
    Call FileManager.Initialize(StreamFile)
    
    '****************************
    'load stream types
    '****************************
    TotalStreams = Val(FileManager.GetValue("INIT", "Total"))
    
    'resize StreamData array
    ReDim StreamData(1 To TotalStreams) As Stream
    
    'clear combo box
    frmMain.List2.Clear
    
    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalStreams

        With StreamData(LoopC)
        
            .Name = FileManager.GetValue(Val(LoopC), "Name")
            .NumOfParticles = FileManager.GetValue(Val(LoopC), "NumOfParticles")
            .x1 = FileManager.GetValue(Val(LoopC), "X1")
            .y1 = FileManager.GetValue(Val(LoopC), "Y1")
            .x2 = FileManager.GetValue(Val(LoopC), "X2")
            .y2 = FileManager.GetValue(Val(LoopC), "Y2")
            .angle = FileManager.GetValue(Val(LoopC), "Angle")
            .vecx1 = FileManager.GetValue(Val(LoopC), "VecX1")
            .vecx2 = FileManager.GetValue(Val(LoopC), "VecX2")
            .vecy1 = FileManager.GetValue(Val(LoopC), "VecY1")
            .vecy2 = FileManager.GetValue(Val(LoopC), "VecY2")
            .life1 = FileManager.GetValue(Val(LoopC), "Life1")
            .life2 = FileManager.GetValue(Val(LoopC), "Life2")
            .friction = FileManager.GetValue(Val(LoopC), "Friction")
            .spin = FileManager.GetValue(Val(LoopC), "Spin")
            .spin_speedL = FileManager.GetValue(Val(LoopC), "Spin_SpeedL")
            .spin_speedH = FileManager.GetValue(Val(LoopC), "Spin_SpeedH")
            .AlphaBlend = FileManager.GetValue(Val(LoopC), "AlphaBlend")
            .gravity = FileManager.GetValue(Val(LoopC), "Gravity")
            .grav_strength = FileManager.GetValue(Val(LoopC), "Grav_Strength")
            .bounce_strength = FileManager.GetValue(Val(LoopC), "Bounce_Strength")
            .XMove = FileManager.GetValue(Val(LoopC), "XMove")
            .YMove = FileManager.GetValue(Val(LoopC), "YMove")
            .move_x1 = FileManager.GetValue(Val(LoopC), "move_x1")
            .move_x2 = FileManager.GetValue(Val(LoopC), "move_x2")
            .move_y1 = FileManager.GetValue(Val(LoopC), "move_y1")
            .move_y2 = FileManager.GetValue(Val(LoopC), "move_y2")
            .Radio = Val(FileManager.GetValue(Val(LoopC), "Radio"))
            .life_counter = FileManager.GetValue(Val(LoopC), "life_counter")
            .speed = Val(FileManager.GetValue(Val(LoopC), "Speed"))
            .grh_resize = Val(FileManager.GetValue(Val(LoopC), "resize"))
            .grh_resizex = Val(FileManager.GetValue(Val(LoopC), "rx"))
            .grh_resizey = Val(FileManager.GetValue(Val(LoopC), "ry"))
            .NumGrhs = FileManager.GetValue(Val(LoopC), "NumGrhs")
        
            ReDim .grh_list(1 To .NumGrhs)
            GrhListing = FileManager.GetValue(Val(LoopC), "Grh_List")
        
            For i = 1 To .NumGrhs
                .grh_list(i) = ReadField(Str$(i), GrhListing, 44)
            Next i
        
            Dim TempSet  As String

            Dim ColorSet As Long
        
            For ColorSet = 1 To 4
                
                TempSet = FileManager.GetValue(Val(LoopC), "ColorSet" & ColorSet)
                
                With .colortint(ColorSet - 1)
                    .r = ReadField(1, TempSet, 44)
                    .g = ReadField(2, TempSet, 44)
                    .B = ReadField(3, TempSet, 44)
                End With
                
            Next ColorSet
        
            'fill stream type combo box
            Call frmMain.List2.AddItem(LoopC & " - " & .Name)
        
        End With
        
    Next LoopC
    
    'set list box index to 1st item
    frmMain.List2.ListIndex = 0
    
    frmMain.CurStreamFile = StreamFile
    
    Set FileManager = Nothing

Exit Sub
Error:

    Call MsgBox("Ha ocurrido un error en la carga de " & StreamFile & ": " & Err.Number & " - " & Err.Description)


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
            Call Particle_Group_Make(Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey, Radio)
        
        Else
            Particle_Group_Create = Particle_Group_Next_Open
            Call Particle_Group_Make(Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey, Radio)

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
        Call Particle_Group_Destroy(particle_group_index)
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
            Call Particle_Group_Destroy(index)
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
    On Error Resume Next
    
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)

    End If

    particle_group_count = particle_group_count + 1
    
    With particle_group_list(particle_group_index)
    
        'Make active
        .active = True
    
        'Map pos
        If (map_x <> -1) And (map_y <> -1) Then
            .map_x = map_x
            .map_y = map_y
        End If
    
        'Grh list
        ReDim .grh_index_list(1 To UBound(grh_index_list))
        .grh_index_list() = grh_index_list()
        .grh_index_count = UBound(grh_index_list)
    
        ' Lord Fers
        .Radio = Radio
    
        'Sets alive vars
        If alive_counter = -1 Then
            .alive_counter = -1
            .never_die = True
        Else
            .alive_counter = alive_counter
            .never_die = False
        End If
    
        'alpha blending
        .alpha_blend = alpha_blend
    
        'stream type
        .stream_type = stream_type
    
        'speed
        .frame_speed = frame_speed
    
        .x1 = x1
        .y1 = y1
        .x2 = x2
        .y2 = y2
        .angle = angle
        .vecx1 = vecx1
        .vecx2 = vecx2
        .vecy1 = vecy1
        .vecy2 = vecy2
        .life1 = life1
        .life2 = life2
        .fric = fric
        .spin = spin
        .spin_speedL = spin_speedL
        .spin_speedH = spin_speedH
        .gravity = gravity
        .grav_strength = grav_strength
        .bounce_strength = bounce_strength
        .XMove = XMove
        .YMove = YMove
        .move_x1 = move_x1
        .move_x2 = move_x2
        .move_y1 = move_y1
        .move_y2 = move_y2
    
        .rgb_list(0) = rgb_list(0)
        .rgb_list(1) = rgb_list(1)
        .rgb_list(2) = rgb_list(2)
        .rgb_list(3) = rgb_list(3)
    
        .grh_resize = grh_resize
        .grh_resizex = grh_resizex
        .grh_resizey = grh_resizey
    
        'create particle stream
        .particle_count = particle_count
        ReDim .particle_stream(1 To particle_count)
    
        'plot particle group on map
        MapData(map_x, map_y).particle_group = particle_group_index
    
    End With

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
    
    With temp_particle
        
        If no_move = False Then
            
            If .alive_counter = 0 Then
                Call InitGrh(.Grh, grh_index, alpha_blend)

                If Radio = 0 Then
                    .x = RandomNumber(x1, x2)
                    .Y = RandomNumber(y1, y2)
                Else
                    .x = (RandomNumber(x1, x2) + Radio) + Radio * Cos(PI * 2 * index / count)
                    .Y = (RandomNumber(y1, y2) + Radio) + Radio * Sin(PI * 2 * index / count)

                End If

                .vector_x = RandomNumber(vecx1, vecx2)
                .vector_y = RandomNumber(vecy1, vecy2)
                .angle = angle
                .alive_counter = RandomNumber(life1, life2)
                .friction = fric
                
            Else

                'Continue old particle
                'Do gravity
                If gravity = True Then
                    .vector_y = .vector_y + grav_strength

                    If .Y > 0 Then
                        'bounce
                        .vector_y = bounce_strength
                    End If

                End If

                'Do rotation
                If spin = True Then .Grh.angle = .Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
                
                'Set angle
                If .angle >= 360 Then .angle = 0
                                
                If XMove = True Then .vector_x = RandomNumber(move_x1, move_x2)
                If YMove = True Then .vector_y = RandomNumber(move_y1, move_y2)

            End If

            'Add in vector
            .x = .x + (.vector_x \ .friction)
            .Y = .Y + (.vector_y \ .friction)
    
            'decrement counter
            .alive_counter = .alive_counter - 1

        End If
 
        'Draw it
        Call Grh_Render(.Grh, .x + screen_x, .Y + screen_y, rgb_list(), 1, False, True, True, .angle)

    End With
    
End Sub

Public Sub Particle_Group_Render(ByVal particle_group_index As Long, _
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
    
    With particle_group_list(particle_group_index)
    
        temp_rgb(0) = .rgb_list(0)
        temp_rgb(1) = .rgb_list(1)
        temp_rgb(2) = .rgb_list(2)
        temp_rgb(3) = .rgb_list(3)
        
        If .alive_counter Then
    
            'See if it is time to move a particle
            .frame_counter = .frame_counter + timerTicksPerFrame

            If .frame_counter > .frame_speed Then
                .frame_counter = 0
                no_move = False
            Else
                no_move = True
            End If
    
            'If it's still alive render all the particles inside
            For LoopC = 1 To .particle_count
        
                'Render particle
                Particle_Render .particle_stream(LoopC), _
                                screen_x, screen_y, _
                                .grh_index_list(Round(RandomNumber(1, .grh_index_count), 0)), _
                                temp_rgb(), _
                                .alpha_blend, no_move, _
                                .x1, .y1, .angle, _
                                .vecx1, .vecx2, _
                                .vecy1, .vecy2, _
                                .life1, .life2, _
                                .fric, .spin_speedL, _
                                .gravity, .grav_strength, _
                                .bounce_strength, .x2, _
                                .y2, .XMove, _
                                .move_x1, .move_x2, _
                                .move_y1, .move_y2, _
                                .YMove, .spin_speedH, _
                                .spin, .grh_resize, .grh_resizex, .grh_resizey, _
                                .Radio, .particle_count, LoopC
                            
            Next LoopC
        
            If no_move = False Then

                'Update the group alive counter
                If .never_die = False Then
                    .alive_counter = .alive_counter - 1
                End If

            End If
    
        Else
            
            'If it's dead destroy it
            .particle_count = .particle_count - 1

            If .particle_count <= 0 Then
                Call Particle_Group_Destroy(particle_group_index)
            End If

        End If
    
    End With

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
            
            With particle_group_list(particle_group_index)
            
                'Move it
                .map_x = map_x
                .map_y = map_y
            
            End With
            
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
        
        With particle_group_list(particle_group_index)
        
            map_x = .map_x
            map_y = .map_y
        
            nX = map_x
            nY = map_y
        
            Call Convert_Heading_to_Direction(Heading, nX, nY)
        
            'Make sure it's a legal move
            If InMapBounds(nX, nY) Then
                
                'Move it
                .map_x = nX
                .map_y = nY
            
                Particle_Group_Move = True

            End If
        
        End With

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



