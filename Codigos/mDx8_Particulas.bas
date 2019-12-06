Attribute VB_Name = "mDx8_Particulas"
Option Explicit

Public Sub LoadStreamFile(ByVal StreamFile As String)

    On Error Resume Next

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
                .grh_list(i) = ReadField(Str(i), GrhListing, 44)
            Next i
        
            Dim TempSet  As String

            Dim ColorSet As Long
        
            For ColorSet = 1 To 4
                TempSet = FileManager.GetValue(Val(LoopC), "ColorSet" & ColorSet)
                .colortint(ColorSet - 1).r = ReadField(1, TempSet, 44)
                .colortint(ColorSet - 1).g = ReadField(2, TempSet, 44)
                .colortint(ColorSet - 1).B = ReadField(3, TempSet, 44)
            Next ColorSet
        
            'fill stream type combo box
            Call frmMain.List2.AddItem(LoopC & " - " & .Name)
        
        End With
        
    Next LoopC
    
    'set list box index to 1st item
    frmMain.List2.ListIndex = 0
    
    Set FileManager = Nothing

End Sub

