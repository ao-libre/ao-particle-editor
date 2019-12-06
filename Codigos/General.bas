Attribute VB_Name = "Mod_General"
Option Explicit

Private Const CANT_GRH_INDEX As Long = 40000

Public Sub LoadGrhData()

    On Error GoTo ErrorHandler

    Dim Grh         As Long
    Dim Frame       As Long
    Dim grhCount    As Long
    Dim handle      As Integer
    Dim fileVersion As Long
   
    'Open files
    handle = FreeFile()

    Open App.Path & "\INIT\Graficos.ind" For Binary Access Read As handle
    Seek #handle, 1
   
    'Get file version
    Get handle, , fileVersion
   
    'Get number of grhs
    Get handle, , grhCount
   
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    Get handle, , Grh

    While Not Grh <= 0

        With GrhData(Grh)
        
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            ReDim .Frames(1 To GrhData(Grh).NumFrames)
           
            If .NumFrames > 1 Then

                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then GoTo ErrorHandler
                Next Frame
               
                Get handle, , .speed
                If .speed <= 0 Then GoTo ErrorHandler
               
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
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
               
                Get handle, , GrhData(Grh).SX
                If .SX < 0 Then GoTo ErrorHandler
               
                Get handle, , .SY
                If .SY < 0 Then GoTo ErrorHandler
               
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
               
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
               
                'Compute width and height
                .TileWidth = .pixelWidth / 32
                .TileHeight = .pixelHeight / 32
               
                .Frames(1) = Grh

            End If

        End With

        Get handle, , Grh
    Wend
   
    Close handle
 
    Exit Sub
 
ErrorHandler:

End Sub

Public Function DirGraficos() As String
    DirGraficos = App.Path & "\GRAFICOS\"

End Function

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, _
                     ByVal Text As String, _
                     Optional ByVal red As Integer = -1, _
                     Optional ByVal green As Integer, _
                     Optional ByVal blue As Integer, _
                     Optional ByVal Bold As Boolean = False, _
                     Optional ByVal Italic As Boolean = False, _
                     Optional ByVal bCrLf As Boolean = False)

    '******************************************
    'Adds text to a Richtext box at the bottom.
    'Automatically scrolls to new text.
    'Text box MUST be multiline and have a 3D
    'apperance!
    '******************************************
    With RichTextBox

        If (Len(.Text)) > 10000 Then .Text = vbNullString
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh

    End With

End Sub

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound

End Function

Sub UnloadAllForms()

    On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms

        Unload mifrm
    Next

End Sub

Sub Main()

    On Error Resume Next
    
    Call ChDrive(App.Path)
    Call ChDir(App.Path)
    
    With frmCargando
        .Show
        .Refresh
        
        Call AddtoRichTextBox(.Status, "Cargando Engine Grafico....")
        
        Call Engine_Init
        Call LoadGrhData
        
        DoEvents
    
        Call AddtoRichTextBox(.Status, "Terminado carga de Engine Grafico con Exito..")
        Call AddtoRichTextBox(.Status, "¡Bienvenido!")
    
    End With
    
    Call AgregaGrH(1)
    Unload frmCargando
                   
    frmMain.Show

    'Inicialización de variables globales
    prgRun = True

    Call Start
    
    Exit Sub
    
ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    Debug.Print "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
    End

End Sub

Public Function General_Particle_Create(ByVal ParticulaInd As Long, _
                                        ByVal x As Integer, _
                                        ByVal Y As Integer, _
                                        Optional ByVal particle_life As Long = 0) As Long

    Dim rgb_list(0 To 3) As Long

    rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).B)
    rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).B)
    rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).B)
    rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).B)

    General_Particle_Create = Particle_Group_Create(x, Y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
       StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
       StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
       StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
       StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
       StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
       StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, _
       StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey, _
       StreamData(ParticulaInd).Radio)

End Function

Public Sub WriteVar(ByVal File As String, _
                    ByVal Main As String, _
                    ByVal var As String, _
                    ByVal Value As String)
                    
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Writes a var to a text file
    '*****************************************************************
    Call writeprivateprofilestring(Main, var, Value, File)

End Sub

Public Function GetVar(ByVal File As String, _
                       ByVal Main As String, _
                       ByVal var As String) As String

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Get a var to from a text file
    '*****************************************************************
    Dim l        As Long
    Dim Char     As String
    Dim sSpaces  As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = vbNullString
    
    sSpaces = Space$(5000)
    
    Call getprivateprofilestring(Main, var, szReturn, sSpaces, Len(sSpaces), File)
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

End Function

Public Function ReadField(ByVal field_pos As Long, _
                          ByVal Text As String, _
                          ByVal delimiter As Byte) As String

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets a field from a delimited string
    '*****************************************************************
    Dim i        As Long
    Dim LastPos  As Long
    Dim FieldNum As Long
    
    LastPos = 0
    FieldNum = 0

    For i = 1 To Len(Text)

        If delimiter = CByte(Asc(mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1

            If FieldNum = field_pos Then
                ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If

            LastPos = i

        End If

    Next i

    FieldNum = FieldNum + 1

    If FieldNum = field_pos Then
        ReadField = mid$(Text, LastPos + 1)
    End If

End Function

Public Sub AgregaGrH(ByVal numgrh As Integer)

    Dim i           As Long
    Dim EsteIndex   As Long
    Dim CuentaIndex As Long
    
    With GrhData(numgrh)
        .FileNum = 1
        .NumFrames = 1
        .pixelHeight = 32
        .pixelWidth = 32
        .Frames(1) = numgrh
    End With
    
    CuentaIndex = -1
    frmMain.lstGrhs.Clear

    For i = 1 To 32000
        
        Select Case GrhData(i).NumFrames
        
            Case Is = 1
                frmMain.lstGrhs.AddItem i
                CuentaIndex = CuentaIndex + 1
                
            Case Is > 1
                frmMain.lstGrhs.AddItem i & " (animacion)"
                CuentaIndex = CuentaIndex + 1
                
        End Select

        If i = numgrh Then
            EsteIndex = CuentaIndex
        End If

    Next i

    frmMain.lstGrhs.ListIndex = EsteIndex

End Sub

Public Function FileExists(ByVal file_path As String, _
                            ByVal file_type As VbFileAttribute) As Boolean

    If LenB(Dir$(file_path, file_type)) = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If

End Function

Public Sub HookSurfaceHwnd(pic As Form)
    Call ReleaseCapture
    Call SendMessage(pic.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Function Grh_Check(ByVal grh_index As Long) As Boolean

    If grh_index > 0 And grh_index <= CANT_GRH_INDEX Then
        Grh_Check = GrhData(grh_index).NumFrames
    End If

End Function
