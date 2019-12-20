Attribute VB_Name = "modCompression"
Option Explicit

Public Const PNG_SOURCE_FILE_EXT As String = ".png"
Public Const BMP_SOURCE_FILE_EXT As String = ".bmp"
Public Const GRH_RESOURCE_FILE As String = "Graphics.AO"
Public Const GRH_PATCH_FILE As String = "Graficos.PATCH"
Public Const MAPS_SOURCE_FILE_EXT As String = ".map"
Public Const MAPS_RESOURCE_FILE As String = "Mapas.AO"
Public Const MAPS_PATCH_FILE As String = "Mapas.PATCH"

Public GrhDatContra() As Byte ' Contrasena
Public GrhUsaContra As Boolean ' Usa Contrasena?

Public MapsDatContra() As Byte ' Contrasena
Public MapsUsaContra As Boolean  ' Usa Contrasena?

'This structure will describe our binary file's
'size, number and version of contained files
Public Type FILEHEADER
    lngNumFiles As Long                 'How many files are inside?
    lngFileSize As Long                 'How big is this file? (Used to check integrity)
    lngFileVersion As Long              'The resource version (Used to patch)
End Type

'This structure will describe each file contained
'in our binary file
Public Type INFOHEADER
    lngFileSize As Long             'How big is this chunk of stored data?
    lngFileStart As Long            'Where does the chunk start?
    strFileName As String * 16      'What's the name of the file this data came from?
    lngFileSizeUncompressed As Long 'How big is the file compressed
End Type

Private Enum PatchInstruction
    Delete_File
    Create_File
    Modify_File
End Enum

Private Declare Function compress Lib "zlib.dll" (dest As Any, destlen As Any, src As Any, ByVal srclen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destlen As Any, src As Any, ByVal srclen As Long) As Long

'BitMaps Strucures
Public Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, bytesTotal As Currency, FreeBytesTotal As Currency) As Long

Public Sub GenerateContra(ByVal Contra As String, Optional Modo As Byte = 0)
'***************************************************
'Author: ^[GS]^
'Last Modification: 17/06/2012 - ^[GS]^
'
'***************************************************

On Error Resume Next

    Dim LoopC As Byte
    Dim Upper_grhDatContra As Long, Upper_mapsDatContra As Long
    
    If Modo = 0 Then
        Erase GrhDatContra
    ElseIf Modo = 1 Then
        Erase MapsDatContra
    End If
    
    If LenB(Contra) <> 0 Then
        If Modo = 0 Then
            ReDim GrhDatContra(Len(Contra) - 1)
            Upper_grhDatContra = UBound(GrhDatContra)
            
            For LoopC = 0 To Upper_grhDatContra
                GrhDatContra(LoopC) = Asc(mid$(Contra, LoopC + 1, 1))
            Next LoopC
            GrhUsaContra = True
        ElseIf Modo = 1 Then
            ReDim MapsDatContra(Len(Contra) - 1)
            Upper_mapsDatContra = UBound(MapsDatContra)
            
            For LoopC = 0 To Upper_mapsDatContra
                MapsDatContra(LoopC) = Asc(mid$(Contra, LoopC + 1, 1))
            Next LoopC
            MapsUsaContra = True
        End If
    Else
        If Modo = 0 Then
            GrhUsaContra = False
        ElseIf Modo = 1 Then
            MapsUsaContra = False
        End If
    End If
    
End Sub

Private Function General_Drive_Get_Free_Bytes(ByVal DriveName As String) As Currency
'**************************************************************
'Author: Juan Martin Sotuyo Dodero
'Last Modify Date: 6/07/2004
'
'**************************************************************
    Dim retval As Long
    Dim FB As Currency
    Dim BT As Currency
    Dim FBT As Currency
    
    retval = GetDiskFreeSpace(Left$(DriveName, 2), FB, BT, FBT)
    
    General_Drive_Get_Free_Bytes = FB * 10000 'convert result to actual size in bytes
End Function

''
' Sorts the info headers by their file name. Uses QuickSort.
'
' @param    InfoHead() The array of headers to be ordered.
' @param    first The first index in the list.
' @param    last The last index in the list.

Private Sub Sort_Info_Headers(ByRef InfoHead() As INFOHEADER, ByVal first As Long, ByVal last As Long)
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/20/2007
'Sorts the info headers by their file name using QuickSort.
'*****************************************************************
    Dim aux As INFOHEADER
    Dim Min As Long
    Dim Max As Long
    Dim comp As String
    
    Min = first
    Max = last
    
    comp = InfoHead((Min + Max) \ 2).strFileName
    
    Do While Min <= Max
        Do While InfoHead(Min).strFileName < comp And Min < last
            Min = Min + 1
        Loop
        Do While InfoHead(Max).strFileName > comp And Max > first
            Max = Max - 1
        Loop
        If Min <= Max Then
            aux = InfoHead(Min)
            InfoHead(Min) = InfoHead(Max)
            InfoHead(Max) = aux
            Min = Min + 1
            Max = Max - 1
        End If
    Loop
    
    If first < Max Then Call Sort_Info_Headers(InfoHead, first, Max)
    If Min < last Then Call Sort_Info_Headers(InfoHead, Min, last)
End Sub

''
' Searches for the specified InfoHeader.
'
' @param    ResourceFile A handler to the data file.
' @param    InfoHead The header searched.
' @param    FirstHead The first head to look.
' @param    LastHead The last head to look.
' @param    FileHeaderSize The bytes size of a FileHeader.
' @param    InfoHeaderSize The bytes size of a InfoHeader.
'
' @return   True if found.
'
' @remark   File must be already open.
' @remark   InfoHead must have set its file name to perform the search.

Private Function BinarySearch(ByRef ResourceFile As Integer, ByRef InfoHead As INFOHEADER, ByVal FirstHead As Long, ByVal LastHead As Long, ByVal FileHeaderSize As Long, ByVal InfoHeaderSize As Long) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/21/2007
'Searches for the specified InfoHeader
'*****************************************************************
    Dim ReadingHead As Long
    Dim ReadInfoHead As INFOHEADER
    
    Do Until FirstHead > LastHead
        ReadingHead = (FirstHead + LastHead) \ 2

        Get ResourceFile, FileHeaderSize + InfoHeaderSize * (ReadingHead - 1) + 1, ReadInfoHead

        If InfoHead.strFileName = ReadInfoHead.strFileName Then
            InfoHead = ReadInfoHead
            BinarySearch = True
            Exit Function
        Else
            If InfoHead.strFileName < ReadInfoHead.strFileName Then
                LastHead = ReadingHead - 1
            Else
                FirstHead = ReadingHead + 1
            End If
        End If
    Loop
End Function

''
' Retrieves the InfoHead of the specified graphic file.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    InfoHead The InfoHead where data is returned.
'
' @return   True if found.

Private Function Get_InfoHeader(ByRef ResourcePath As String, ByRef FileName As String, ByRef InfoHead As INFOHEADER, Optional Modo As Byte = 0) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 16/07/2012 - ^[GS]^
'Retrieves the InfoHead of the specified graphic file
'*****************************************************************
    Dim ResourceFile As Integer
    Dim ResourceFilePath As String
    Dim FileHead As FILEHEADER
    
    Dim ERROR_LEER_ARCHIVO As String
    
On Local Error GoTo ErrHandler

    If Modo = 0 Then
        ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
    End If
    
    'Set InfoHeader we are looking for
    InfoHead.strFileName = UCase$(FileName)
   
    'Open the binary file
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Extract the FILEHEADER
        Get ResourceFile, 1, FileHead
        
        'Check the file for validity
        If LOF(ResourceFile) <> FileHead.lngFileSize Then
            Call MsgBox("El archivo " & ResourceFilePath & " se encuentra corrupto.", , "modCompression")
            Close ResourceFile
            Exit Function
        End If
        
        'Search for it!
        If BinarySearch(ResourceFile, InfoHead, 1, FileHead.lngNumFiles, Len(FileHead), Len(InfoHead)) Then
            Get_InfoHeader = True
        End If
        
    Close ResourceFile
Exit Function

ErrHandler:
    Close ResourceFile
    
    If Err.Number <> 0 Then
        Call MsgBox("Error al leer el archivo descomprimido. Error: " & Err.Number & " - " & Err.Description)
    End If
    
End Function

''
' Compresses binary data avoiding data loses.
'
' @param    data() The data array.

Private Sub Compress_Data(ByRef data() As Byte, Optional Modo As Byte = 0)
'*****************************************************************
'Author: Juan Martin Dotuyo Dodero
'Last Modify Date: 17/07/2012 - ^[GS]^
'Compresses binary data avoiding data loses
'*****************************************************************
    Dim Dimensions As Long
    Dim DimBuffer As Long
    Dim BufTemp() As Byte
    Dim LoopC As Long
    Dim Upper_grhDatContra As Long, Upper_mapsDatContra As Long
    
    Dimensions = UBound(data) + 1
    
    ' The worst case scenario, compressed info is 1.06 times the original - see zlib's doc for more info.
    DimBuffer = Dimensions * 1.06
    
    ReDim BufTemp(DimBuffer)
    
    Call compress(BufTemp(0), DimBuffer, data(0), Dimensions)
    
    Erase data
    
    ReDim data(DimBuffer - 1)
    ReDim Preserve BufTemp(DimBuffer - 1)
    
    data = BufTemp
    
    Erase BufTemp
    
    ' GSZAO - Seguridad
    If Modo = 0 And GrhUsaContra = True Then
        If UBound(GrhDatContra) <= UBound(data) And UBound(GrhDatContra) <> 0 Then
            Upper_grhDatContra = UBound(GrhDatContra)
            
            For LoopC = 0 To Upper_grhDatContra
                data(LoopC) = data(LoopC) Xor GrhDatContra(LoopC)
            Next LoopC
        End If
    ElseIf Modo = 1 And MapsUsaContra = True Then
        If UBound(MapsDatContra) <= UBound(data) And UBound(MapsDatContra) <> 0 Then
            Upper_mapsDatContra = UBound(MapsDatContra)
            
            For LoopC = 0 To Upper_mapsDatContra
                data(LoopC) = data(LoopC) Xor MapsDatContra(LoopC)
            Next LoopC
        End If
    End If
    ' GSZAO - Seguridad
    
End Sub

''
' Decompresses binary data.
'
' @param    data() The data array.
' @param    OrigSize The original data size.

Private Sub Decompress_Data(ByRef data() As Byte, ByVal OrigSize As Long, Optional Modo As Byte = 0)
'*****************************************************************
'Author: Juan Martin Dotuyo Dodero
'Last Modify Date: 16/07/2012 - ^[GS]^
'Decompresses binary data
'*****************************************************************
    Dim BufTemp() As Byte
    Dim LoopC As Integer
    Dim Upper_grhDatContra As Long, Upper_mapsDatContra As Long
    
    ReDim BufTemp(OrigSize - 1)
    
    ' GSZAO - Seguridad
    If Modo = 0 And GrhUsaContra = True Then
        If UBound(GrhDatContra) <= UBound(data) And UBound(GrhDatContra) <> 0 Then
            Upper_grhDatContra = UBound(GrhDatContra)
            
            For LoopC = 0 To Upper_grhDatContra
                data(LoopC) = data(LoopC) Xor GrhDatContra(LoopC)
            Next LoopC
        End If
    ElseIf Modo = 1 And MapsUsaContra = True Then
        If UBound(MapsDatContra) <= UBound(data) And UBound(MapsDatContra) <> 0 Then
            Upper_mapsDatContra = UBound(MapsDatContra)
            
            For LoopC = 0 To Upper_mapsDatContra
                data(LoopC) = data(LoopC) Xor MapsDatContra(LoopC)
            Next LoopC
        End If
    End If
    ' GSZAO - Seguridad
    
    Call uncompress(BufTemp(0), OrigSize, data(0), UBound(data) + 1)
    
    ReDim data(OrigSize - 1)
    
    data = BufTemp
    
    Erase BufTemp
End Sub

''
' Retrieves a byte array with the compressed data from the specified file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   InfoHead must not be encrypted.
' @remark   Data is not desencrypted.

Public Function Get_File_RawData(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef data() As Byte, Optional Modo As Byte = 0) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 16/07/2012 - ^[GS]^
'Retrieves a byte array with the compressed data from the specified file
'*****************************************************************
    Dim ResourceFilePath As String
    Dim ResourceFile As Integer
    
On Local Error GoTo ErrHandler
    If Modo = 0 Then
        ResourceFilePath = ResourcePath & GRH_RESOURCE_FILE
    ElseIf Modo = 1 Then
        ResourceFilePath = ResourcePath & MAPS_RESOURCE_FILE
    End If
    
    'Size the Data array
    ReDim data(InfoHead.lngFileSize - 1)
    
    'Open the binary file
    ResourceFile = FreeFile
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Get the data
        Get ResourceFile, InfoHead.lngFileStart, data
    'Close the binary file
    Close ResourceFile
    
    Get_File_RawData = True
Exit Function

ErrHandler:
    Close ResourceFile
End Function

''
' Extract the specific file from a resource file.
'
' @param    ResourcePath The resource file folder.
' @param    InfoHead The header specifiing the graphic file info.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Extract_File(ByRef ResourcePath As String, ByRef InfoHead As INFOHEADER, ByRef data() As Byte, Optional Modo As Byte = 0) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 14/09/2012 - ^[GS]^
'Extract the specific file from a resource file
'*****************************************************************
On Local Error GoTo ErrHandler
    
    If Get_File_RawData(ResourcePath, InfoHead, data, Modo) Then
        'Decompress all data
        'If InfoHead.lngFileSize < InfoHead.lngFileSizeUncompressed Then ' GSZAO
            Call Decompress_Data(data, InfoHead.lngFileSizeUncompressed, Modo)
        'End If
        
        Extract_File = True
    End If
Exit Function

ErrHandler:
    Call MsgBox("Error al extraer los recursos del Graphics.AO", vbOKOnly, "modCompression")
End Function

''
' Retrieves a byte array with the specified file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.
'
' @remark   Data is desencrypted.

Public Function Get_File_Data(ByRef ResourcePath As String, ByRef FileName As String, ByRef data() As Byte, Optional Modo As Byte = 0) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 16/07/2012 - ^[GS]^
'Retrieves a byte array with the specified file data
'*****************************************************************
    Dim InfoHead As INFOHEADER
    
    If Get_InfoHeader(ResourcePath, FileName, InfoHead, Modo) Then
        'Extract!
        Get_File_Data = Extract_File(ResourcePath, InfoHead, data, Modo)
    Else
        Get_File_Data = False
        'Call MsgBox(JsonLanguage("ERROR_404").Item("TEXTO") & ": " & FileName)
    End If
End Function

''
' Retrieves image file data.
'
' @param    ResourcePath The resource file folder.
' @param    FileName The graphic file name.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   True if no error occurred.

Public Function Get_Image(ByRef ResourcePath As String, ByRef FileName As String, ByRef data() As Byte, Optional SoloBMP As Boolean = False) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 09/10/2012 - ^[GS]^
'Retrieves image file data
'*****************************************************************
    Dim InfoHead As INFOHEADER
    Dim ExistFile As Boolean
    
    ExistFile = False
    
    If SoloBMP = True Then
        If Get_InfoHeader(ResourcePath, FileName & ".BMP", InfoHead, 0) Then ' BMP?
            FileName = FileName & ".BMP"
            ExistFile = True
        End If
    Else
        If Get_InfoHeader(ResourcePath, FileName & ".BMP", InfoHead, 0) Then ' BMP?
            FileName = FileName & ".BMP"
            ExistFile = True
        ElseIf Get_InfoHeader(ResourcePath, FileName & ".PNG", InfoHead, 0) Then ' Existe PNG?
            FileName = FileName & ".PNG" ' usamos el PNG
            ExistFile = True
        End If
    End If
    
    If ExistFile = True Then
        If Extract_File(ResourcePath, InfoHead, data, 0) Then Get_Image = True
    Else
        Call MsgBox("No se encuentra el recurso: " & FileName, vbCritical, "Descompresion de Recursos")
    End If
End Function

''
' Compare two byte arrays to detect any difference.
'
' @param    data1() Byte array.
' @param    data2() Byte array.
'
' @return   True if are equals.

Private Function Compare_Datas(ByRef data1() As Byte, ByRef data2() As Byte) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 02/11/2007
'Compare two byte arrays to detect any difference
'*****************************************************************
    Dim length As Long
    Dim act As Long
    
    length = UBound(data1) + 1
    
    If (UBound(data2) + 1) = length Then
        While act < length
            If data1(act) Xor data2(act) Then Exit Function
            
            act = act + 1
        Wend
        
        Compare_Datas = True
    End If
End Function

''
' Retrieves the next InfoHeader.
'
' @param    ResourceFile A handler to the resource file.
' @param    FileHead The reource file header.
' @param    InfoHead The returned header.
' @param    ReadFiles The number of headers that have already been read.
'
' @return   False if there are no more headers tu read.
'
' @remark   File must be already open.
' @remark   Used to walk through the resource file info headers.
' @remark   The number of read files will increase although there is nothing else to read.
' @remark   InfoHead is encrypted.

Private Function ReadNext_InfoHead(ByRef ResourceFile As Integer, ByRef FileHead As FILEHEADER, ByRef InfoHead As INFOHEADER, ByRef ReadFiles As Long) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 08/24/2007
'Reads the next InfoHeader
'*****************************************************************

    If ReadFiles < FileHead.lngNumFiles Then
        'Read header
        Get ResourceFile, Len(FileHead) + Len(InfoHead) * ReadFiles + 1, InfoHead
        
        'Update
        ReadNext_InfoHead = True
    End If
    
    ReadFiles = ReadFiles + 1
End Function

''
' Retrieves the next bitmap.
'
' @param    ResourcePath The resource file folder.
' @param    ReadFiles The number of bitmaps that have already been read.
' @param    bmpInfo The bitmap info structure.
' @param    data() The byte array to return data.
'
' @return   False if there are no more bitmaps tu get.
'
' @remark   Used to walk through the resource file bitmaps.

Public Function GetNext_Bitmap(ByRef ResourcePath As String, ByRef ReadFiles As Long, ByRef bmpInfo As BITMAPINFO, ByRef data() As Byte, ByRef fileIndex As Long) As Boolean
'*****************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 09/10/2012 - ^[GS]^
'Reads the next InfoHeader
'*****************************************************************
On Error Resume Next

    Dim ResourceFile As Integer
    Dim FileHead As FILEHEADER
    Dim InfoHead As INFOHEADER
    Dim FileName As String
    
    ResourceFile = FreeFile
    Open ResourcePath & GRH_RESOURCE_FILE For Binary Access Read Lock Write As ResourceFile
    Get ResourceFile, 1, FileHead
    
    If ReadNext_InfoHead(ResourceFile, FileHead, InfoHead, ReadFiles) Then
        Call Get_Image(ResourcePath, InfoHead.strFileName, data())
        FileName = Trim$(InfoHead.strFileName)
        fileIndex = CLng(Left$(FileName, Len(FileName) - 4))
        
        GetNext_Bitmap = True
    End If
    
    Close ResourceFile
End Function

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
'*****************************************************************
'Author: Unknown
'Last Modify Date: Unknown
'*****************************************************************
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8
End Function

''
' Retrieves the version number of a given resource file.
'
' @param    ResourceFilePath The resource file complete path.
'
' @return   The version number of the given file.

Public Function GetVersion(ByVal ResourceFilePath As String) As Long
'*****************************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/23/2008
'
'*****************************************************************
    Dim ResourceFile As Integer
    Dim FileHead As FILEHEADER
    
    ResourceFile = FreeFile()
    Open ResourceFilePath For Binary Access Read Lock Write As ResourceFile
        'Extract the FILEHEADER
        Get ResourceFile, 1, FileHead
        
    Close ResourceFile
    
    GetVersion = FileHead.lngFileVersion
End Function

