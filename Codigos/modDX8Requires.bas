Attribute VB_Name = "modDX8Requires"
Option Explicit

Public vertList(3) As TLVERTEX

Public Type D3D8Textures
    texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type

Public DirectX        As DirectX8
Public D3D       As Direct3D8
Public D3DDevice As Direct3DDevice8
Public D3DX      As D3DX8

Public Type TLVERTEX
    x As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

Public Type TLVERTEX2
    x As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu1 As Single
    tv1 As Single
    tu2 As Single
    tv2 As Single
End Type

Public Const PI   As Single = 3.14159265358979

'To get free bytes in drive
Private Declare Function GetDiskFreeSpace _
                Lib "kernel32" _
                Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
                                             FreeBytesToCaller As Currency, _
                                             BytesTotal As Currency, _
                                             FreeBytesTotal As Currency) As Long

'To get free bytes in RAM
Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Function General_Bytes_To_Megabytes(Bytes As Double) As Double

    Dim dblAns As Double

    dblAns = (Bytes / 1024) / 1024
    General_Bytes_To_Megabytes = format(dblAns, "###,###,##0.00")

End Function

Public Function General_Get_Free_Ram() As Double

    'Return Value in Megabytes
    Dim dblAns As Double

    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    General_Get_Free_Ram = General_Bytes_To_Megabytes(dblAns)

End Function

Public Function General_Get_Free_Ram_Bytes() As Long
    GlobalMemoryStatus pUdtMemStatus
    General_Get_Free_Ram_Bytes = pUdtMemStatus.dwAvailPhys

End Function

Public Function ARGB(ByVal r As Long, _
                     ByVal g As Long, _
                     ByVal B As Long, _
                     ByVal a As Long) As Long
        
    Dim c As Long
        
    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or B
    Else
        c = a * 2 ^ 24
        c = c Or r * 2 ^ 16
        c = c Or g * 2 ^ 8
        c = c Or B

    End If
    
    ARGB = c

End Function

