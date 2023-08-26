Attribute VB_Name = "modFileIO_Bitness"
                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||        modFileIO_Bitness (v1.0)       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                            ' _
    AUTHOR:   Kallun Willock                                                                                                                                                                ' _
                                                                                                                                                                                            ' _
    NOTES:    Function used to identify the bitness of a given PE (Portable Executable) files - include executables,                                                                        ' _
              object code, DLLs (Dynamic Link Libraries), and other Windows-based files. Especially useful                                                                                  ' _
              when dealing with third-party DLLs in VBA.
                                                                                                                                                                                            ' _
    LICENSE:  MIT                                                                                                                                                                           ' _
    VERSION:  1.0        27/08/2023         Published v1.0 on Github.                                                                                                                       ' _
                                                                                                                                                                                            ' _                                                                                                                                                   ' _
    TODO:     [ ] Add GetBinaryType API (for completeness)
                                                                                                                                                                                            ' _
                    #If VBA7 Then                                                                                                                                                           ' _
                        Private Declare PtrSafe Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeW" (ByVal lpApplicationName As LongPtr, ByRef lpBinaryType As Long) As Long       ' _
                    #Else                                                                                                                                                                   ' _
                        Private Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeW" (ByVal lpApplicationName As Long, ByRef lpBinaryType As Long) As Long                  ' _
                    #End If                                                                                                                                                                 ' _
                    Private Enum EXE_BINARY_TYPE_ENUM                                                                                                                                       ' _
                        SCS_32BIT_BINARY        ' 32-bit Windows-based application                                                                                                          ' _
                        SCS_DOS_BINARY          ' MS-DOS – based application                                                                                                                ' _
                        SCS_WOW_BINARY          ' 16-bit Windows-based application                                                                                                          ' _
                        SCS_PIF_BINARY          ' PIF file that executes an MS-DOS – based application                                                                                      ' _
                        SCS_POSIX_BINARY        ' POSIX – based application                                                                                                                 ' _
                        SCS_OS216_BINARY        ' 16-bit OS/2-based application                                                                                                             ' _
                        SCS_64BIT_BINARY        ' 64-bit Windows-based application.                                                                                                         ' _
                    End Enum
                    
    Option Explicit
    
    Public Enum FILE_BITNESS_ENUM
        BITNESS_UNKNOWN
        x86_32BIT
        x64_64BIT
    End Enum
    
    Public Sub TestFileBitness()
        
        Dim BitnessType() As String, Result As FILE_BITNESS_ENUM
        BitnessType = Split("Unknown|32-bit|64-bit", "|")
        
        Result = GetFileBitness(ThisWorkbook.Path & Application.PathSeparator & "DemoDLL_win64.dll")
        Debug.Print BitnessType(Result)
        
        Result = GetFileBitness(ThisWorkbook.Path & Application.PathSeparator & "DemoDLL_win32.dll")
        Debug.Print BitnessType(Result)
        
    End Sub
    
    Public Function GetFileBitness(ByVal FilePath As String) As FILE_BITNESS_ENUM
        
        On Error GoTo ErrHandler
        Dim Offset      As Long: Offset = &H3C
        Dim FFile       As Long: FFile = FreeFile
        Dim FileData()  As Byte
        
        Open FilePath For Binary Access Read As #FFile
        
        ' Get the offset to PE signature
        ReDim FileData(3)
        Get #FFile, Offset + 1, FileData
        Offset = FileData(&H0) * 256 ^ 0 Or FileData(&H1) * 256 ^ 1 Or _
                 FileData(&H2) * 256 ^ 2 Or FileData(&H3) * 256 ^ 3
        Erase FileData
        
        ' Read the file bitness signature using the above PE signature offset
        ReDim FileData(5)
        Get #FFile, Offset + 1, FileData
        
        If FileData(&H0) = &H50 And FileData(&H1) = &H45 And FileData(&H2) = &H0 And FileData(&H3) = &H0 Then
            If FileData(&H4) = &H64 And FileData(&H5) = &H86 Then GetFileBitness = x64_64BIT
            If FileData(&H4) = &H4C And FileData(&H5) = &H1 Then GetFileBitness = x86_32BIT
        End If
ErrHandler:
        Close #FFile
        Erase FileData
        
    End Function
    
