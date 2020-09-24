Attribute VB_Name = "mdHookFunction"
Option Explicit

'--- will Debug.Print module imports
#Const SHOW_MODULE_IMPORTS = False

Public Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES   As Long = 16
Public Const IMAGE_DIRECTORY_ENTRY_IMPORT       As Long = 1 ' Import Directory
Public Const IMAGE_ORDINAL_FLAG32               As Long = &H80000000
Public Const PAGE_READWRITE                     As Long = &H4
Public Const VER_PLATFORM_WIN32_NT              As Long = 2

Public Declare Function VirtualQuery Lib "kernel32" (lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Public Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Type IMAGE_IMPORT_DESCRIPTOR
    OriginalFirstThunk As Long              ' RVA to original unbound IAT (PIMAGE_THUNK_DATA)
    TimeDateStamp As Long                   ' 0 if not bound,
                                            ' -1 if bound, and real date\time stamp
                                            '     in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND)
                                            ' O.W. date/time stamp of DLL bound to (Old BIND)

    ForwarderChain As Long                  ' -1 if no forwarders
    Name As Long
    FirstThunk As Long                      ' RVA to IAT (if bound this IAT has actual addresses)
End Type

Public Type IMAGE_DOS_HEADER               ' DOS .EXE header
    e_magic As Integer                     ' Magic number
    e_cblp As Integer                      ' Bytes on last page of file
    e_cp As Integer                        ' Pages in file
    e_crlc As Integer                      ' Relocations
    e_cparhdr As Integer                   ' Size of header in paragraphs
    e_minalloc As Integer                  ' Minimum extra paragraphs needed
    e_maxalloc As Integer                  ' Maximum extra paragraphs needed
    e_ss As Integer                        ' Initial (relative) SS value
    e_sp As Integer                        ' Initial SP value
    e_csum As Integer                      ' Checksum
    e_ip As Integer                        ' Initial IP value
    e_cs As Integer                        ' Initial (relative) CS value
    e_lfarlc As Integer                    ' File address of relocation table
    e_ovno As Integer                      ' Overlay number
    e_res(0 To 3) As Integer                    ' Reserved words
    e_oemid As Integer                     ' OEM identifier (for e_oeminfo)
    e_oeminfo As Integer                   ' OEM information; e_oemid specific
    e_res2(0 To 9) As Integer                  ' Reserved words
    e_lfanew As Long                    ' File address of new exe header
End Type

Public Type IMAGE_FILE_HEADER
    Machine                 As Integer
    NumberOfSections        As Integer
    TimeDateStamp           As Long
    PointerToSymbolTable    As Long
    NumberOfSymbols         As Long
    SizeOfOptionalHeader    As Integer
    Characteristics         As Integer
End Type

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Public Type IMAGE_OPTIONAL_HEADER
    '
    ' Standard fields.
    '

    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode  As Long
    BaseOfData  As Long

    '
    ' NT additional fields.
    '

    ImageBase  As Long
    SectionAlignment  As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To IMAGE_NUMBEROF_DIRECTORY_ENTRIES - 1) As IMAGE_DATA_DIRECTORY
End Type

Public Type IMAGE_NT_HEADERS
    Signature As Long
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type

Public Type IMAGE_THUNK_DATA32
    FunctionOrOrdinalOrAddress As Long
End Type

Public Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Public Function HookImportedFunctionByName( _
            ByVal hModule As Long, _
            szImportMod As String, _
            szImportFunc As String, _
            ByVal pFuncAddress As Long, _
            pOrigAddress As Long) As Boolean
    Dim udtDesc         As IMAGE_IMPORT_DESCRIPTOR
    Dim pOrigThunk      As Long
    Dim pRealThunk      As Long
    Dim udtOrigThunk    As IMAGE_THUNK_DATA32
    Dim udtRealThunk    As IMAGE_THUNK_DATA32
    Dim sBuffer         As String
    Dim udtMem          As MEMORY_BASIC_INFORMATION
    Dim lOldProtect     As Long
    Dim lNotUsed        As Long

    On Error Resume Next
    '--- parameters check
    If hModule = 0 Or pFuncAddress = 0 Or szImportMod = "" Or szImportFunc = "" Then
        Exit Function
    End If
    '--- dll above 2G on 9x -> NOT working!!!!
    If hModule > &H80000000 Then
        If Not IsNT() Then
            Exit Function
        End If
    End If
    '--- get Import Descriptor
    If Not GetNamedImportDescriptor(hModule, szImportMod, udtDesc) Then
        Exit Function
    End If
    '--- guard offset
    If udtDesc.FirstThunk = 0 Or udtDesc.OriginalFirstThunk = 0 Then
        Exit Function
    End If
    '--- loop Real and Original thunks
    pOrigThunk = hModule + udtDesc.OriginalFirstThunk
    pRealThunk = hModule + udtDesc.FirstThunk
    '--- dereference Original Thunk
    CopyMemory udtOrigThunk, ByVal pOrigThunk, LenB(udtOrigThunk)
    Do While udtOrigThunk.FunctionOrOrdinalOrAddress <> 0
        '--- check if imported by name
        If (udtOrigThunk.FunctionOrOrdinalOrAddress And IMAGE_ORDINAL_FLAG32) = 0 Then
#If SHOW_MODULE_IMPORTS Then
            sBuffer = String(1024, 0)
            lstrcpy sBuffer, hModule + udtOrigThunk.FunctionOrOrdinalOrAddress + 2
            Debug.Print Left(sBuffer, InStr(1, sBuffer, Chr(0)))
#End If
            '--- case-insensitive compare
            If lstrcmpi(szImportFunc, hModule + udtOrigThunk.FunctionOrOrdinalOrAddress + 2) = 0 Then
                '--- set read/write access to pRealThunk
                VirtualQuery ByVal pRealThunk, udtMem, LenB(udtMem)
                If VirtualProtect(ByVal udtMem.BaseAddress, udtMem.RegionSize, PAGE_READWRITE, lOldProtect) = 0 Then
                    '--- ooops!
                    Exit Function
                End If
                '--- save orig func address and change to our func address
                CopyMemory udtRealThunk, ByVal pRealThunk, LenB(udtRealThunk)
                pOrigAddress = udtRealThunk.FunctionOrOrdinalOrAddress
                udtRealThunk.FunctionOrOrdinalOrAddress = pFuncAddress
                CopyMemory ByVal pRealThunk, udtRealThunk, LenB(udtRealThunk)
                '--- restore protection
                VirtualProtect ByVal udtMem.BaseAddress, udtMem.RegionSize, lOldProtect, lNotUsed
                '--- success
                HookImportedFunctionByName = True
                Exit Function
            End If
        End If
        '--- check next thunks
        pOrigThunk = pOrigThunk + LenB(udtOrigThunk)
        pRealThunk = pRealThunk + LenB(udtRealThunk)
        CopyMemory udtOrigThunk, ByVal pOrigThunk, LenB(udtOrigThunk)
    Loop
End Function

Public Function GetNamedImportDescriptor( _
            ByVal hModule As Long, _
            szImportMod As String, _
            udtDesc As IMAGE_IMPORT_DESCRIPTOR) As Boolean
    Dim udtDosHeader    As IMAGE_DOS_HEADER
    Dim udtNtHeaders    As IMAGE_NT_HEADERS
    Dim udtImportDesc   As IMAGE_IMPORT_DESCRIPTOR
    Dim pImportDesc     As Long
    
    On Error Resume Next
    '--- dereference DOS Header
    CopyMemory udtDosHeader, ByVal hModule, LenB(udtDosHeader)
    '--- dereference NT Header
    CopyMemory udtNtHeaders, ByVal hModule + udtDosHeader.e_lfanew, LenB(udtNtHeaders)
    '--- check if any imports
    If udtNtHeaders.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress = 0 Then
        Exit Function
    End If
    '--- loop and dereference Import Descriptions
    pImportDesc = hModule + udtNtHeaders.OptionalHeader.DataDirectory(IMAGE_DIRECTORY_ENTRY_IMPORT).VirtualAddress
    CopyMemory udtImportDesc, ByVal pImportDesc, LenB(udtImportDesc)
    Do While udtImportDesc.Name <> 0
        '--- case-insensitive compare
        If lstrcmpi(szImportMod, hModule + udtImportDesc.Name) = 0 Then
            udtDesc = udtImportDesc
            '--- success
            GetNamedImportDescriptor = True
            Exit Function
        End If
        '--- dereference next Import Descriptions in the array
        pImportDesc = pImportDesc + LenB(udtImportDesc)
        CopyMemory udtImportDesc, ByVal pImportDesc, LenB(udtImportDesc)
    Loop
End Function

Public Function IsNT() As Boolean
    Dim udtVer As OSVERSIONINFO
    
    On Error Resume Next
    udtVer.dwOSVersionInfoSize = Len(udtVer)
    If GetVersionEx(udtVer) Then
        If udtVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            IsNT = True
        End If
    End If
End Function

