VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRESreader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 99% of the API extraction is from vbAccelerator with some minor modifications.
' 99% of the IDE Res File Extraction is custom designed from scratch

'//////////////////////////////////////////////////////////////////////////////////////////////////
' The following is used to extract resources from an IDE RES file
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
' used to create a stdPicture from a byte array
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long

Private Enum eResTypes    ' Values < 256 are reserved for system use.
    RT_CURSOR = 1
    RT_BITMAP = 2
    RT_ICON = 3
    RT_MENU = 4
    RT_DIALOG = 5
    RT_STRING = 6
    RT_FONTDIR = 7
    RT_FONT = 8
    RT_ACCELERATOR = 9
    RT_RCDATA = 10
    RT_MESSAGETABLE = 11
    RT_GROUP_CURSOR = 12
    RT_GROUP_ICON = 14
    RT_VERSION = 16
    RT_DLGINCLUDE = 17
    RT_PLUGPLAY = 19
    RT_VXD = 20
    RT_ANICURSOR = 21
    RT_ANIICON = 22
    RT_HTML = 23
    RT_MANIFEST = 24
End Enum
Private Type uResGroupInfo
    ID As String            ' ID or Name (i.e., 101, "MYGIF")
    DataOffset As Long      ' Where in file the data begins for this item
    DataLength As Long      ' The length of the data
End Type
Private Type uResGroup
    Name As String          ' ID or Name (i.e., "Bitmap", "CUSTOM")
    Data() As uResGroupInfo ' collection of child resources
    Attr As Byte            ' resource type: RT_ICON, etc. 0=Custom resource
End Type

' PROTOTYPE RESOURCEHEADER
'    DataSize    AS DWORD ' bytes of resource data following header, not counting any padding used for alignment
'    HeaderSize  AS DWORD
'    ResType(0)  AS WORD  ' variable length: numeric id or UnicodeZ string data
'    ResName(0)  AS WORD  ' variable length: numeric id or UnicodeZ string data
'    Padding     AS ????  ' 0-1 WORDs of padding to DWORD-align next item
'    DataVersion AS DWORD
'    MemoryFlags AS WORD
'    LanguageId  AS WORD
'    Version     AS DWORD
'    Characteristics AS DWORD
' END PROTOTYPE
'-------------------------------------------------------------------------------------------------

'//////////////////////////////////////////////////////////////////////////////////////////////////
' The following are used to extract resources from the application when it is a compiled exe
Private Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal uType As Long, ByVal cX As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare Function LoadImageString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cX As Long, ByVal cy As Long, ByVal uFlags As Long) As Long

Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, lpName As Any, lpType As Any) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long

Private Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumResourceNamesByNum Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Any, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Private Type PictDesc
    Size As Long
    Type As Long
    hHandle As Long
    hPal As Long
End Type
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2
'-------------------------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////////////////////////
' This allows the class to be stand-alone, else would require a module for the Enumeration APIs
' Much appreciation goes towards Paul Caton's self-subclassing thunks; makes some things so much easier
'-Thunking/Callback declarations---------------------------------------------------------------------------
Private z_CbMem   As Long    'Callback allocated memory address
Private z_Cb()    As Long    'Callback thunk array

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'-------------------------------------------------------------------------------------------------


' module-wide variables
Private m_Resources() As uResGroup  ' collection of resource Names/offsets
Private m_FileName As String        ' IDE RES file
Private m_ManagedGroups As Long     ' number of groups we will expose
Private m_callBackAddr As Long      ' thunk function address & flag


Public Function ScanResources(RESfilename As String) As Boolean

    ' Function extracts Resource Group/Item names, IDs and maybe offsets
    ' Depending on whether the RESfilename is null or not, the extraction is from File or Process
    
    ' Parameter: RESfilename
    '   - When vbNullString, then the routine will look for resources from the compiled application
    '   - Otherwise, the parameter must be full path to the resource file (.RES) to be processed.
    
    ' This class does not care about, nor expose, every resource type/group. See AddResGroup for more.


    Dim I As Long
    
    ' reset & clear if class is being reused
    m_FileName = vbNullString
    Erase m_Resources
    m_ManagedGroups = 0
        
    If RESfilename = vbNullString Then
    
        Dim b() As Byte, hMem As Long, lPtr As Long, hMod As Long
        
        m_callBackAddr = cb_AddressOf(2, 3, 1)  ' using Paul Caton's callback thunks, get address for 2nd to last Routine below
        If Not m_callBackAddr = 0 Then
        
            hMod = GetModuleHandleA(vbNullString)
            ' Special Note for Usercontrols:
            '   When wanting to scan resources of the host application, the one hosting your uc,
            '   don't use app.hInstance. Doing so will scan the resources of your compiled UC, not the host.
            
            ' get the list of resource groups (i.e., Bitmap, Icon, Custom, etc)
            If Not EnumResourceTypes(hMod, m_callBackAddr, 0) = 0 Then
                
                ' now we loop thru each of the groups we cached & will get info for its individual child resources
                m_callBackAddr = cb_AddressOf(1, 4, 0) ' using Paul Caton's callback thunks, get address for last Routine below
                For I = 0 To UBound(m_Resources)
                
                    If m_Resources(I).Attr = 0 Then
                    
                        ' when .Attr=0, then we have a custom resource; not a standard one
                        b = StrConv(m_Resources(I).Name, vbFromUnicode)
                        ' custom resources have custom string names: i.e., CUSTOM
                        ' Convert to a byte array for API usage
                        ReDim Preserve b(0 To UBound(b) + 1) As Byte
                        hMem = GlobalAlloc(&H40, UBound(b) + 1)
                        If Not hMem = 0 Then
                           lPtr = GlobalLock(hMem)
                           If Not lPtr = 0 Then
                              ' good; get the child resource data
                              CopyMemory ByVal lPtr, b(0), UBound(b) + 1
                              EnumResourceNamesByNum hMod, lPtr, m_callBackAddr, I '<< I is our UDT reference
                              GlobalUnlock lPtr
                           End If
                           GlobalFree hMem
                        End If
                    Else
                        ' standard resources are like RT_ICON, RT_BITMAP. A bit easier to call
                        EnumResourceNamesByNum hMod, CLng(m_Resources(I).Attr), m_callBackAddr, I '<< I is our UDT reference
                    End If
                    
                    ' The routines called by above enumeration APIs adds an extra item to our group UDT
                    ' so remove it now. Maintenance
                    If UBound(m_Resources(I).Data) > 0 Then
                        ReDim Preserve m_Resources(I).Data(0 To UBound(m_Resources(I).Data) - 1)
                    End If
                Next
            End If
        End If
        cb_Terminate ' clear callback memory. But don't zero out becauase we use it as a flag that we are compiled vs IDE
        
    Else
    
        ' parsing an IDE RES file
        
        Dim nFile   As Integer      ' file number
        Dim sBuffer As String * 32  ' string input
        Dim sResHeader As String * 32
        
        ' magic number: 0,0,0,0,&H20,0,0,0,&HFF,&HFF,0,0,&HFF,&HFF,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
        ' Equates to: 0 data size, 32byte header size, 0 Resource Group ID, & 0 Resource Item ID
        For I = 1 To 32 ' build the resource file header (1st 32 bytes of a RES file)
            Select Case I
            Case 5:
                Mid$(sResHeader, I, 1) = Chr$(&H20)
            Case 9, 10, 13, 14
                Mid$(sResHeader, I, 1) = Chr$(&HFF)
            Case Else
                Mid$(sResHeader, I, 1) = Chr$(&H0)
            End Select
        Next
    
        On Error Resume Next    ' attempt to open file
        If Len(Dir$(RESfilename, vbHidden Or vbReadOnly Or vbReadOnly Or vbSystem)) = 0 Then
            Exit Function
        End If
        
        m_FileName = RESfilename                        ' cache file name for later if needed
        nFile = FreeFile
        Open m_FileName For Binary Access Read As #nFile
        If Err Then Exit Function
    
        If LOF(nFile) < 64 Then ' expecting the magic number from above & also the 1st real header
            Close nFile
            Exit Function
        End If
    
        Get #nFile, 1, sBuffer
        If Not sBuffer = sResHeader Then
            ' first 32bytes better match our magic string above
            Close nFile
            Exit Function
        End If
    
        I = LOF(nFile)
        ' now we just read til there is nothing more to read
        Do Until RetrieveFileData(nFile, I) = False
            ' reading, reading
        Loop
    
        Close nFile
        
    End If
    
    ' here we are going to place our resources we won't expose on the end
    ' of the collection; just to keep them organized and out of the way...
    
    If Not IsArrayEmpty(Not m_Resources) Then
    
        Dim tmpResGroup As uResGroup  ' udt for swapping
        
        m_ManagedGroups = UBound(m_Resources)
        
        For I = m_ManagedGroups To 0 Step -1
            
            Select Case m_Resources(I).Attr
            Case RT_ICON, RT_CURSOR, RT_FONT
                ' we won't be exposing the above, but we need them
                ' Those 3 resource "items" are the raw bytes and by themselves mean very little
                
                If Not I = m_ManagedGroups Then
                    ' move to end of array
                    tmpResGroup = m_Resources(I)
                    m_Resources(I) = m_Resources(m_ManagedGroups)
                    m_Resources(m_ManagedGroups) = tmpResGroup
                End If
                m_ManagedGroups = m_ManagedGroups - 1   ' keep track of nr managed/exposed
            Case Else
            End Select
        Next
        
        m_ManagedGroups = m_ManagedGroups + 1
        ScanResources = Not (m_ManagedGroups = 0)
    
    End If
            
End Function

Public Function ResourceSection(ByVal Index As Long) As String
    
    ' Function returns the group name of the index being passed (i.e., "CUSTOM")
    ' Index is the zero-bound index of the resource groups, continue calling
    ' routine until ResourceSection returns vbNullString
    
    If Index < m_ManagedGroups And Index > -1 Then ResourceSection = m_Resources(Index).Name

End Function

Public Function ResourceID(ResSection As Variant, ByVal Index As Long) As String
        
    ' Function returns the child resource id or name (i.e., 101, "MYGIF")
    ' ResSection can be vbResBitmap , vbResIcon, vbResCursor, Custom
    ' can also be just Bitmap, Icon, Cursor
    ' Index is the zero-bound index of the resource, continue calling
    ' routine until ResourceID returns vbNullString
    
    Dim grpID As Long
    ResolveResNames ResSection, 0, grpID, 0
    If grpID < m_ManagedGroups Then
        If Index <= UBound(m_Resources(grpID).Data) And Index > -1 Then
            ResourceID = m_Resources(grpID).Data(Index).ID
        End If
    End If

End Function

Public Function ExtractResourceItem(ResSection As Variant, ResID As Variant, outStream() As Byte, Picture As StdPicture) As Boolean

    ' Purpose: Extract a passed resource reference
    
    ' Parameters.
    ' ResSection : The numeric or string resource group name (i.e., vbResBitmap, CUSTOM, etc)
    ' ResID : The numeric or string resource item name (i.e., 101, MyGIF, etc)
    ' outStream() : an empty byte array
    '   if the resource is not vbResBitmap,vbResIcon,vbResCursor then the extraction will be placed in the array
    ' Picture : any stdPicture object
    '   if the resource is vbResBitmap,vbResIcon,vbResCursor then the extraction will be converted to stdPicture
    
    ' The return value is True if extraction was successful
    ' Test your passed stdPicture to see if extraction is in stdPicture or array...
    '   Example:  If myStdPicParam Is Nothing ' then extraction is in array else in stdPicture
    
    Dim grpID As Long, Index As Long
    ResolveResNames ResSection, ResID, grpID, Index
    
    If grpID < m_ManagedGroups Then
        If Index <= UBound(m_Resources(grpID).Data) Then
            
            If m_callBackAddr = 0 Then  ' uncompiled; resources in RES file
                ExtractResourceItem = ExtractResourceFile(grpID, Index, Picture, outStream())
            
            Else    ' trying to read from compiled exe
                Select Case m_Resources(grpID).Attr
                Case RT_BITMAP, RT_GROUP_ICON, RT_GROUP_CURSOR
                    ExtractResourceItem = PictureFromResource(grpID, Index, Picture)
                Case Else
                    ExtractResourceItem = ExtractResStreamAPI(grpID, Index, outStream())
                End Select
            End If
        End If
    End If

End Function

Private Function AddResGroup(wID As Variant) As Long

    ' Function creates a group UDT to store only the following
    ' resource types where the Attr < 256
    ' The rest of the types are left in for reference only
    
    ' BTW: Where Attr=256, these aren't available in a standard VB Res file
    ' However, they could be in other RES files, and you are welcome to
    ' remove the 256 Attr flag so they can be enumerated. However, you
    ' should also provide a way of exposing/extracting the resource in
    ' the ExtractResourceItem routine.

    Dim resName As String, Index As Long, Attr As Integer
    If VarType(wID) = vbString Then
        ' non-standard resources (i.e., "CUSTOM")
        resName = wID
    Else
        Attr = CByte(wID)
        Select Case wID
        Case RT_CURSOR: resName = "Application Cursor"
            ' these are the actual cursor image bytes + 4 bytes (X & Y hotspots)
        Case RT_BITMAP: resName = "Bitmap"
            ' these are the actual bitmap bytes, less the BitmapFileHeader
        Case RT_ICON: resName = "Application Icon"
            ' these are the actual icon bytes when combined with the
            ' RT_GROUP_ICON form a complete icon
        Case RT_GROUP_CURSOR: resName = "Cursor"
            ' Basically the CUSORDIR & CURSORDIRENTRY structures, but different than standard structure
        Case RT_GROUP_ICON: resName = "Icon"
            ' This is an ICONDIR & ICONDIRENTRY structures, slightly different than standard structure
            
        Case RT_MENU: resName = "Menu": Attr = 256
        Case RT_DIALOG: resName = "Dialog Box": Attr = 256
        Case RT_STRING: resName = "String Table": Attr = 256
            ' VB has string tables; haven't had time to reverse engineer
            ' how they are stored in VB's Res file
        Case RT_FONTDIR: resName = "Font Directory": Attr = 256
        Case RT_FONT: resName = "Font": Attr = 256
        Case RT_ACCELERATOR: resName = "Accelerator": Attr = 256
        Case RT_RCDATA: resName = "Raw Data": Attr = 256
        Case RT_MESSAGETABLE: resName = "Message Table": Attr = 256
        Case RT_VERSION: resName = "Version information": Attr = 256
        Case RT_DLGINCLUDE: resName = ".RC header filename": Attr = 256
        Case RT_PLUGPLAY: resName = "Plug and Play resource": Attr = 256
        Case RT_VXD: resName = "VXD device driver": Attr = 256
        Case RT_ANICURSOR: resName = "Cursor - Animated": Attr = 256
        Case RT_ANIICON: resName = "Icon - Animated": Attr = 256
        Case RT_HTML: resName = "HTML": Attr = 256
        Case RT_MANIFEST: resName = "XML Manifest for Windows XP": Attr = 256
        Case Else: Attr = 256 ' unknown type
        End Select
    End If
    
    If Attr = 256 Then
        AddResGroup = -1    ' we won't cache this data
    Else
        If IsArrayEmpty(Not m_Resources) Then   ' first group
            ReDim m_Resources(0 To 0)
        Else
            For Index = 0 To UBound(m_Resources)    ' see if group exists
                If m_Resources(Index).Name = resName Then Exit For
            Next
            If Index > UBound(m_Resources) Then     ' new group
                ReDim Preserve m_Resources(0 To Index)
            Else
                AddResGroup = Index ' return index & increment number of child resources
                ReDim Preserve m_Resources(Index).Data(0 To UBound(m_Resources(Index).Data) + 1)
                Exit Function
            End If
        End If
        m_Resources(Index).Name = resName       ' add name to group
        m_Resources(Index).Attr = Attr          ' add resource type
        ReDim m_Resources(Index).Data(0 To 0)   ' prime child resource (see AddResItem)
        AddResGroup = Index                     ' return index
    End If
    
End Function

Private Function AddResItem(gID As Long, wID As Variant) As Long
    
    ' These are the children of the resource groups
    
    If gID = -1 Then
        AddResItem = -1 ' we don't want data for groups we aren't caching
    Else
        Dim Index As Long
        Index = UBound(m_Resources(gID).Data)
        m_Resources(gID).Data(Index).ID = wID
        AddResItem = Index
    End If
End Function

Private Function RetrieveFileData(ByVal nFile As Integer, fileSize As Long) As Boolean

    ' This is only called by the ScanResources routine when reading an IDE RES file
    
    ' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/resources/introductiontoresources/resourcereference/resourcestructures/resourceheader.asp

    Dim dcbDataSize   As Long   ' size of resource data
    Dim dcbHeaderSize As Long   ' size of header record
    Dim FileOffset   As Long    ' offset of header record within file
    ' Long, Integer & String variables for reading the file
    Dim resDataL As Long, resDataI As Integer
    Dim resDataS As String
    
    Dim resGrpID As Long        ' Group collection item
    Dim resItemID As Long       ' Child collection item

    FileOffset = Seek(nFile)    ' Cache current file position

    Get #nFile, , dcbDataSize   ' get the data size of the resource (byte size)
    If Err Then Exit Function

    Get #nFile, , dcbHeaderSize ' the header of the Prototype structure
    ' minimum header size is 32, but can exceed 32 when string names are used for resources
    If fileSize < FileOffset + dcbDataSize + dcbHeaderSize Then Exit Function
    
    ' Note: the header must be word aligned at a specific point which
    ' means we may have to move the file pointer at times. Additionally,
    ' the dcbData must also be word aligned, which means the file pointer
    ' may be moved yet again when we are done reading this resource
    
    Get #nFile, , resDataI      ' get resource ID flag
    If resDataI = &HFFFF Then   ' if -1 then numerical ID
        Get #nFile, , resDataI  ' and next 2bytes are the ID
        resGrpID = AddResGroup(resDataI)
    Else
        ' we have a unicode, null-terminated string...
        Do Until resDataI = 0       ' count characters, unicode=2bytes per char
            resDataL = resDataL + 1
            Get #nFile, , resDataI
        Loop
        ' now go back to the beginning to cache it; we could have concatenated the string
        Seek nFile, Seek(nFile) - resDataL * 2 - 2
        ' build string of appropriate length
        resDataS = String$(resDataL * 2, Chr$(0))
        Get #nFile, , resDataS  ' read the string
        Get #nFile, , resDataI  ' and dispose of the null terminator too
        resDataL = 0             ' reset for the child resource, if needed
        ' add this resource group to our collection
        resGrpID = AddResGroup(StrConv(resDataS, vbFromUnicode))
    End If

    If Not resGrpID = -1 Then ' else a resource we aren't caching
    
        Get #nFile, , resDataI      ' get ID type of child resource
        If resDataI = &HFFFF Then   ' numerical ID
            Get #nFile, , resDataI
            resItemID = AddResItem(resGrpID, resDataI)
        Else
            ' same exact comments as above
            Do Until resDataI = 0
                resDataL = resDataL + 1
                Get #nFile, , resDataI
            Loop
            Seek nFile, Seek(nFile) - resDataL * 2 - 2
            resDataS = String$(resDataL * 2, Chr$(0))
            Get #nFile, , resDataS
            resItemID = AddResItem(resGrpID, StrConv(resDataS, vbFromUnicode))
        End If
    
        ' As mentioned above, the resource data must start on a DWORD
        ' aligned boundary; here we are ensuring that...
        resDataL = (FileOffset + dcbHeaderSize + 2) And Not 3
        m_Resources(resGrpID).Data(resItemID).DataOffset = resDataL + 1
        m_Resources(resGrpID).Data(resItemID).DataLength = dcbDataSize
        
    End If
    
    ' There is more information that follows; specifically 16 bytes of info
    ' But for our purposes, we won't be needing it.
    ' See the PROTOTYPE structure at top of module: .DataVersion would be next
    
    ' The next record must start on a DWORD boundary to. Ensure it...
    FileOffset = ((FileOffset + dcbHeaderSize + dcbDataSize) + 2) And Not 3
    
    ' are we at the end?
    If Not FileOffset + 32 >= fileSize Then
        Seek nFile, FileOffset + 1  ' set position for next resource
        RetrieveFileData = True
    End If

End Function

Private Sub ResolveResNames(grpCommon As Variant, itmCommon As Variant, grpID As Long, ItemID As Long)

    ' PURPOSE: Converts string names into references to our m_Resources() collection
    
    If VarType(grpCommon) = vbString Then
        ' compare string IDs/Names
        For grpID = 0 To m_ManagedGroups - 1
            If StrComp(m_Resources(grpID).Name, grpCommon, vbTextCompare) = 0 Then Exit For
        Next
    Else
        ' compare to VB constants: vbResBitmap , vbResIcon, vbResCursor
        Dim bCriteria As Byte
        Select Case grpCommon
        Case vbResBitmap: bCriteria = RT_BITMAP
        Case vbResIcon: bCriteria = RT_GROUP_ICON
        Case vbResCursor: bCriteria = RT_GROUP_CURSOR
        Case Else
        End Select
        If bCriteria = 0 Then
            grpID = m_ManagedGroups
        Else
            For grpID = 0 To m_ManagedGroups - 1
                If m_Resources(grpID).Attr = bCriteria Then Exit For
            Next
        End If
    End If
    
    If grpID < m_ManagedGroups Then
        ' look for the specific item
        For ItemID = 0 To UBound(m_Resources(grpID).Data)
            If StrComp(m_Resources(grpID).Data(ItemID).ID, itmCommon, vbTextCompare) = 0 Then Exit For
        Next
    End If
        
End Sub

Private Function ExtractResourceFile(GroupID As Long, ItemIndex As Long, stdPic As StdPicture, outStream() As Byte) As Boolean
    
    ' Purpose: This is the function that extracts resources from the IDE RES file. See ExtractResStreamAPI for compiled extraction.
    
    Dim fNR As Integer
    Dim iBitCount As Integer
    Dim lSize As Long, lSize2 As Long
    
    ' we already processed it once; but did the user delete it or move it on us?
    fNR = FreeFile()
    On Error Resume Next
    Open m_FileName For Binary Access Read As #fNR
    If Err Then
        Close #fNR
        Err.Clear
        Exit Function
    End If

    With m_Resources(GroupID).Data(ItemIndex)
        
        ReDim outStream(0 To .DataLength - 1)       ' we already know its offset & length
        Get #fNR, .DataOffset, outStream()          ' simply read it from the file
        ' however, for known images (bmp,ico,cur), we will mimic VB
        ' and return the resource as a stdPicture object
        
        Select Case m_Resources(GroupID).Attr
        
        Case RT_BITMAP:
            ' The vbResBitmap section contains everything one would need to
            ' create a DIB bitmap. I am just supplyng the 14byte BITMAPFILEHEADER
            ' so that the stream can be read by APIs or saved to file
            
            CopyMemory lSize, outStream(20), 4&    ' size Image
            CopyMemory lSize2, outStream(0), 4&    ' BMPinfoHeader Size
            ReDim Preserve outStream(0 To .DataLength + 13)
            CopyMemory outStream(14), outStream(0), .DataLength
            CopyMemory outStream(0), &H4D42, 2& ' bmp magic number
            CopyMemory outStream(2), CLng(.DataLength + 14), 4& ' overall size of image
            CopyMemory outStream(6), 0&, 4&         ' reserved bytes
            CopyMemory iBitCount, outStream(28), 2& ' bit count
            CopyMemory lSize, outStream(18), 4&     ' width
            CopyMemory lSize2, outStream(22), 4&    ' height
            
            ' calculate where in the array the image bytes will start; these will
            ' be after the BITMAPINFOHEADER for >8bpp and after the palette for others
            lSize = .DataLength - (ByteAlignOnWord(iBitCount, lSize) * lSize2) + 14
            CopyMemory outStream(10), lSize, 4& ' image offset from beginning of file
            ' convert to stdPicture
            Set stdPic = CreateStdPicFromArray(outStream(), 0, .DataLength + 14)
            Erase outStream
        
        Case RT_GROUP_ICON: ' icon; send to routine to gather multi-icons if needed
            GatherGroupItems RT_ICON, fNR, .DataOffset, outStream()
            Set stdPic = CreateStdPicFromArray(outStream(), 0, UBound(outStream) + 1)
            Erase outStream
            
        Case RT_GROUP_CURSOR: ' cursor; send to routine to gather multi-cursors if needed
            GatherGroupItems RT_CURSOR, fNR, .DataOffset, outStream()
            Set stdPic = CreateStdPicFromArray(outStream(), 0, UBound(outStream) + 1)
            Erase outStream
            
        Case Else
            ' add code for any other resources you care to extract
            
        End Select
    End With
    Close #fNR
    ExtractResourceFile = True
    
End Function

Private Sub GatherGroupItems(whichGroup As Long, FileNum As Integer, fOffset As Long, rtnArray() As Byte)

    ' Purpose: A resource icon/cursor file may have 1 or several items. This routine gathers the items,
    ' builds their directory structures and then stacks that data neatly onto the passed rtnArray()
    
    Dim itemData() As Long
    Dim fileDat() As Byte
    Dim arraySize As Long
    Dim nrItems As Long
    Dim iValue As Integer, lValue As Long
    Dim rtIcon As Long
    
    ' for each icon and/or cursor, there exists a RT_ICON / RT_CURSOR
    ' section that contains bascially the ICONDIR/CUSORDIR information.
    ' Find that directory
    For rtIcon = 0 To UBound(m_Resources)
        If m_Resources(rtIcon).Attr = whichGroup Then Exit For
    Next
    
    arraySize = 6   ' icons/cursors both begin with a 6 byte header
    Get #FileNum, fOffset + 4, iValue ' retrieve nr of icons/cursors
    arraySize = arraySize + iValue * 16 ' add 16byte header for each one
    ' NOTE: 16bytes is the window standard ICONDIRENTRY structure; however,
    ' resource files use a 14byte header. So we read 14bytes & write 16bytes
    
    ' calculate the size of the array we will need
    ReDim itemData(0 To iValue - 1)
    For nrItems = 0 To iValue - 1
    
        ' retrieve the icon/cursor ID for each one that belongs to
        ' the selected Directory. They will always be numeric
        ' and can be found as the last 2 bytes within that res 14 byte header
        Get #FileNum, fOffset + 6 + (nrItems * 14) + 12, iValue
        
        ' is the ID one of the ones we are looking for?
        For lValue = 0 To UBound(m_Resources(rtIcon).Data)
            If m_Resources(rtIcon).Data(lValue).ID = iValue Then
                ' yep, cache its RT_ICON/CURSOR child reference
                itemData(nrItems) = lValue
                Exit For
            End If
        Next
        ' go on to the next icon/cursor as needed
    Next
    
    ' ok, now add up the amount of bytes each one will use
    For nrItems = 0 To UBound(itemData)
        arraySize = arraySize + m_Resources(rtIcon).Data(itemData(nrItems)).DataLength
    Next
    
    ' here we are building the 6 byte header
    iValue = UBound(itemData) + 1
    ReDim rtnArray(0 To arraySize - 1)
    If whichGroup = RT_CURSOR Then
        CopyMemory rtnArray(2), 2, 2&   ' cursor
    Else
        CopyMemory rtnArray(2), 1, 2&   ' icon
    End If
    CopyMemory rtnArray(4), iValue, 2&  ' nr of icons/cursors
    lValue = iValue * 16 + 6            ' offset for first item
    
    For nrItems = 0 To iValue - 1
        
        ReDim fileDat(1 To 14) ' directories are 16 bytes, but in resources they are 14
        Get #FileNum, fOffset + 6 + (nrItems * 14), fileDat()   ' get next icondirectory entry
        If rtnArray(2) = 1 Then 'icon
            fileDat(2) = fileDat(2) \ 2 ' height is expressed * 2
            iValue = 0                  ' file offset
        Else                    ' cursor
            ' not quite as expected. Whereas the std struture uses bytes for these
            ' properties, the res structure uses Integers; shift array values
            fileDat(2) = fileDat(3) ' height
            fileDat(3) = fileDat(5) ' color count if used
            fileDat(4) = fileDat(6) ' reserved
            
            ' the 1st 4 bytes of the res data for cursors will be the X,Y hotspots.
            ' again unique. Get those values & put them in their place within our array
            Get #FileNum, m_Resources(rtIcon).Data(itemData(nrItems)).DataOffset, iValue
            CopyMemory fileDat(5), iValue, 2&   ' X hotspot
            Get #FileNum, m_Resources(rtIcon).Data(itemData(nrItems)).DataOffset + 2, iValue
            CopyMemory fileDat(7), iValue, 2&   ' Y hotspot
            ' the data length is really 4 more than needed now; we offset
            iValue = 4
            ' copy the new data length into the cursor directory
            CopyMemory fileDat(9), (m_Resources(rtIcon).Data(itemData(nrItems)).DataLength - iValue), 4&
        End If
        
        ' copy the icon/cursor directory to our array
        CopyMemory rtnArray(6 + (nrItems * 16)), fileDat(1), 14&
        ' add the image offset value to the directory
        CopyMemory rtnArray(6 + (nrItems * 16) + 12), lValue, 4&
        
        ' now copy the item data
        ReDim fileDat(1 To m_Resources(rtIcon).Data(itemData(nrItems)).DataLength - iValue)
        Get #FileNum, m_Resources(rtIcon).Data(itemData(nrItems)).DataOffset + iValue, fileDat()
        CopyMemory rtnArray(lValue), fileDat(1), m_Resources(rtIcon).Data(itemData(nrItems)).DataLength - iValue
        
        ' calculate next position in our array for the next icon/cursor
        lValue = lValue + m_Resources(rtIcon).Data(itemData(nrItems)).DataLength - iValue
        
    Next
    Erase fileDat

End Sub

Private Function ExtractResStreamAPI(ByVal grpID As Long, ByVal Index As Long, outStream() As Byte) As Boolean

    ' This routine extracts data from custom resources only.

    Dim hGbl As Long
    Dim hRes As Long
    Dim lSize As Long
    Dim lPtr As Long
    Dim hMod As Long
    Dim sName As String
    Dim gName As String
   
    On Error GoTo ErrorHandler
   
    sName = m_Resources(grpID).Data(Index).ID       ' res item name
    If IsNumeric(sName) Then sName = "#" & sName    ' format for numeric IDs
    
    hMod = GetModuleHandleA(vbNullString) ' see note in ScanResources why we don't use app.hInstance
    If m_Resources(grpID).Attr = 0 Then
        ' custom resource, use string resource name
        hRes = FindResource(hMod, ByVal sName, ByVal m_Resources(grpID).Name)
    Else
        ' standard resource; use numeric reference
        hRes = FindResource(hMod, ByVal sName, ByVal CLng(m_Resources(grpID).Attr))
    End If
    
    If Not hRes = 0 Then                    ' resource found?
        hGbl = LoadResource(hMod, hRes)     ' good, load it & read it
        If Not hGbl = 0 Then
            lPtr = LockResource(hGbl)       ' per MSDN, no need to unlock
            If Not lPtr = 0 Then
                lSize = SizeofResource(hMod, hRes)
                If lSize > 0 Then
                    ReDim outStream(0 To lSize - 1)
                    CopyMemory outStream(0), ByVal lPtr, lSize
                    ExtractResStreamAPI = True
                End If
            End If
        End If
        FreeResource hRes
        hRes = 0
    End If
      
ErrorHandler:
    If Not hRes = 0 Then FreeResource hRes
End Function

Private Function PictureFromResource(ByVal grpID As Long, ByVal Index As Long, stdPic As StdPicture) As Boolean

    ' This routine extracts only images from the resource

    Dim hImage As Long
    Dim lID As Long
    Dim hMod As Long
    Dim sName As String

    hMod = GetModuleHandleA(vbNullString) ' see note in ScanResources why we don't use app.hInstance
    sName = m_Resources(grpID).Data(Index).ID          ' resource item string name
    If IsNumeric(sName) = True Then lID = CLng(sName)  ' value of name when name is numeric
    
    ' vbResBitmap
    If m_Resources(grpID).Attr = RT_BITMAP Then
        If lID = 0 Then
            hImage = LoadImageString(hMod, sName, IMAGE_BITMAP, 0, 0, 0)
        Else
            hImage = LoadImageLong(hMod, lID, IMAGE_BITMAP, 0, 0, 0)
        End If
        If Not hImage = 0 Then
            Set stdPic = CreateStdPicFromHandle(hImage, vbPicTypeBitmap)
            PictureFromResource = True
        End If
    
    ' vbResIcon
    ElseIf m_Resources(grpID).Attr = RT_GROUP_ICON Then
        If lID = 0 Then
            hImage = LoadImageString(hMod, sName, IMAGE_ICON, 0, 0, 0)
        Else
            hImage = LoadImageLong(hMod, lID, IMAGE_ICON, 0, 0, 0)
        End If
        If Not hImage = 0 Then
            Set stdPic = CreateStdPicFromHandle(hImage, vbPicTypeIcon)
            PictureFromResource = True
        End If
    
    ' vbResCursor
    ElseIf m_Resources(grpID).Attr = RT_GROUP_CURSOR Then
        If lID = 0 Then
            hImage = LoadImageString(hMod, sName, IMAGE_CURSOR, 0, 0, 0)
        Else
            hImage = LoadImageLong(hMod, lID, IMAGE_CURSOR, 0, 0, 0)
        End If
        If Not hImage = 0 Then
            Set stdPic = CreateStdPicFromHandle(hImage, vbPicTypeIcon)
            PictureFromResource = True
        End If
    
    End If
   
End Function

Private Function ByteAlignOnWord(ByVal bitDepth As Byte, ByVal Width As Long) As Long
    ' function to align any bit depth on dWord boundaries
    ByteAlignOnWord = (((Width * bitDepth) + &H1F&) And Not &H1F&) \ &H8&
End Function

Private Function IsArrayEmpty(ByVal lArrayPointer As Long) As Boolean
  ' test to see if an array has been initialized
  IsArrayEmpty = (lArrayPointer = -1&)
End Function

Public Function CreateStdPicFromArray(inArray() As Byte, Offset As Long, Size As Long) As IPicture
    
    ' function creates a stdPicture from the passed array
    
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(Offset), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), CreateStdPicFromArray)
            End If
        End If
    End If

End Function

Private Function CreateStdPicFromHandle(ByVal hImage As Long, ByVal imgType As Long) As IPicture
    ' function creates a stdPicture object from a image handle (bitmap or icon)
    Dim lpPictDesc As PictDesc, aGUID(0 To 3) As Long
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = imgType
        .hHandle = hImage
        .hPal = 0
    End With
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGUID(0) = &H7BF80980
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc, aGUID(0), True, CreateStdPicFromHandle)
    
End Function


'-Callback code-----------------------------------------------------------------------------------
Private Function cb_AddressOf(ByVal nOrdinal As Long, _
                              ByVal nParamCount As Long, _
                     Optional ByVal nThunkNo As Long = 0, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
'*************************************************************************************************
'* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
'* nParamCount  - The number of parameters that will callback
'* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
'* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety   - Optional, set to false to disable IDE protection.
'*************************************************************************************************
Const MAX_FUNKS   As Long = 2                                               'Number of simultaneous thunks, adjust to taste
Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
  Dim nAddr       As Long
  
  If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
    MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
  If nAddr = 0 Then
    MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If z_CbMem = 0 Then                                                       'If memory hasn't been allocated
    ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
    z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
  End If
  
  If z_Cb(0, nThunkNo) = 0 Then                                             'If this ThunkNo hasn't been initialized...
    z_Cb(3, nThunkNo) = _
              GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
    z_Cb(4, nThunkNo) = &HBB60E089
    z_Cb(5, nThunkNo) = VarPtr(z_Cb(0, nThunkNo))                           'Set the data address
    z_Cb(6, nThunkNo) = &H73FFC589: z_Cb(7, nThunkNo) = &HC53FF04: z_Cb(8, nThunkNo) = &H7B831F75: z_Cb(9, nThunkNo) = &H20750008: z_Cb(10, nThunkNo) = &HE883E889: z_Cb(11, nThunkNo) = &HB9905004: z_Cb(13, nThunkNo) = &H74FF06E3: z_Cb(14, nThunkNo) = &HFAE2008D: z_Cb(15, nThunkNo) = &H53FF33FF: z_Cb(16, nThunkNo) = &HC2906104: z_Cb(18, nThunkNo) = &H830853FF: z_Cb(19, nThunkNo) = &HD87401F8: z_Cb(20, nThunkNo) = &H4589C031: z_Cb(21, nThunkNo) = &HEAEBFC
  End If
  
  z_Cb(0, nThunkNo) = ObjPtr(oCallback)                                     'Set the Owner
  z_Cb(1, nThunkNo) = nAddr                                                 'Set the callback address
  
  If bIdeSafety Then                                                        'If the user wants IDE protection
    z_Cb(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
  End If
    
  z_Cb(12, nThunkNo) = nParamCount                                          'Set the parameter count
  z_Cb(17, nThunkNo) = nParamCount * 4                                      'Set the number of stck bytes to release on thunk return
  
  nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                   'Calculate where in the allocated memory to copy the thunk
  RtlMoveMemory nAddr, VarPtr(z_Cb(0, nThunkNo)), FUNK_LEN                  'Copy thunk code to executable memory
  cb_AddressOf = nAddr + 16                                                 'Thunk code start address
End Function

'Terminate the callback thunks
Private Sub cb_Terminate()
Const MEM_RELEASE As Long = &H8000&                                         'Release allocated memory flag

  If z_CbMem <> 0 Then                                                      'If memory allocated
    If VirtualFree(z_CbMem, 0, MEM_RELEASE) <> 0 Then                       'Release
      z_CbMem = 0                                                           'Indicate memory released
    End If
  End If
End Sub

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim I     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, I, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, I, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, I, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  I = I + 4                                                                 'Bump to the next entry
  J = I + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While I < J
    RtlMoveMemory VarPtr(nAddr), I, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    I = I + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

'*************************************************************************************************
'* Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
'*************************************************************************************************

'Callback ordinal 2
Private Function EnumResTypesProc(ByVal hMod As Long, _
                            ByVal lpszType As Long, ByVal lParam As Long) As Long
On Error GoTo EH
    
    Dim lType As Long
    Dim sType As String
    Dim lLen As Long
    Dim b() As Byte
    If (lpszType And &HFFFF0000) = 0 Then
       ' standard resource type:
       lType = lpszType And &HFFFF&
       AddResGroup lType    ' add to our m_Resources() collection
    Else
       ' string:
       lLen = lstrlen(lpszType)
       If (lLen > 0) Then
          ReDim b(0 To lLen - 1)
          CopyMemory b(0), ByVal lpszType, lLen
          sType = StrConv(b, vbUnicode)
       End If
       AddResGroup sType    ' add to our m_Resources() collection
    End If
   
    EnumResTypesProc = 1    ' keep enumerating
    
EH:
End Function

'Callback ordinal 1
Private Function EnumResNamesProc(ByVal hMod As Long, _
                                ByVal lpszType As Long, ByVal lpszName As Long, _
                                ByVal lParam As Long) As Long
On Error GoTo EH

    ' Note: we passed a value to lParam when we called the
    ' API that is using this callback. lParam is the parent
    ' m_Resources() item for this resource item

    Dim lName As Long, sName As String
    Dim lLen As Long, b() As Byte
    
    If (lpszName And &HFFFF0000) = 0 Then
       ' resource number:
       lName = lpszName And &HFFFF&
       AddResItem lParam, lName ' add to parent m_Resources() item
    Else
       ' resource string:
       lLen = lstrlen(lpszName)
       If (lLen > 0) Then
          ReDim b(0 To lLen - 1) As Byte
          CopyMemory b(0), ByVal lpszName, lLen
          sName = StrConv(b, vbUnicode)
       End If
       AddResItem lParam, sName ' add to our m_Resources() collection
    End If
    ReDim Preserve m_Resources(lParam).Data(0 To UBound(m_Resources(lParam).Data) + 1)
   
   EnumResNamesProc = 1 ' keep enumerating

EH:
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' CAUTION: DO NOT ADD ANY ADDITIONAL CODE OR COMMENTS PAST THE "END FUNCTION"
'          STATEMENT BELOW. Paul Caton's zProbe routine will read it as a start
'          of a new function/sub and the class timer's will fail every time.
End Function
