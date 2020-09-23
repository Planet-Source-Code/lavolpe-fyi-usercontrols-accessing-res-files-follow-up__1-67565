VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   PropertyPages   =   "UserControl1.ctx":0000
   ScaleHeight     =   1695
   ScaleWidth      =   3690
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1200
      Left            =   135
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   3315
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This project is really wrapped around the cResReader class.
' The other classes are supporting files for images processing.

' THIS IS NOT a functional usercontrol. Its only purpose is to show a how a
' usercontrol can retreive images from a resource file when the host is in IDE
' and/or compiled. Please do not ask for updates to this usercontrol or the
' property page, that is for you to mess with, the project is highlighting the
' cResReader.  BTW: currently, I am using something very similar to this for
' a project I'm working.

' The usercontrol's property page is where the magic kinda happens.

' This concept has one minor drawback but a few advantages.
' Disadvantage:
'   Each UC caches a full path to the resource file. This is wasted space (stored in form's frx file)
'      when the project is compiled.  But without that path, the UC can never display images from the
'      resouce file when the host project is uncompiled & in design mode, unless, you force the user to
'      re-select the path every time they open their project in IDE - thought about it & IMO, that is poor

' Advantages are primarily for IDE but also relate in some extent to a compiled application
'   The major advantage is that UCs do not need to store images if the host app will have them in its resource file
'       -- typically UCs store their images into the host form's .frx file
'       -- also when more than one UC uses same image, each UC stores a copy of the image
'       -- and if images are not stored by the UC, then no image is displayed during IDE
'       -- all of the above are not applicable if resource file is used & we get our images from there
' Note: that the cResReader class must be part of the usercontrol's project too. That class will
'   parse the resource file when host is uncompiled and use APIs to read it when the host is a compiled exe.


' Theoretically, this process could also apply to an ImageList; but will have to play with that idea later



' The following are used to get IDE/run/compiled states of both the UC and its host
'----------------------------------------------------------------------------------
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Const GW_OWNER As Long = 4
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Enum eUCstates  ' this is used so the UC knows to have the cResReader class get resources from uncompiled or compiled resource file
    uStUCcompiled = 1   ' else uc uncompiled : Not 1
    uStHostCompiled = 2 ' else host uncompiled : Not 2
    uStHostRunMode = 4  ' else host in ide : Not 4
    uStUCinHost = 8     ' else uc is separate project : Not 8
End Enum
Private m_UCstate As eUCstates
'----------------------------------------------------------------------------------


' these apis are used only for the sample project
Private Declare Function GetBkColor Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private m_LockAspectRatio As Boolean    ' just for sample purpose only
Private m_ResFile As String             ' path to IDE resource file, if used
Private m_ResImgSec As String           ' resource sectoin, if used
Private m_ResImgID As String            ' resource item ID, if used

Private m_Image As c32bppDIB            ' this usercontrol sample incorporates my c32bppDIB suite; can be modified to use stdPictures instead

Private m_ImgScaleWidth As Long         ' cached scaled width/height of the actual
Private m_ImgScaleHeight As Long        '   image compared to display DC
'


Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "BackColor" Then
        On Error Resume Next
        UserControl.BackColor = UserControl.Extender.Container.BackColor
        If Err Then
            UserControl.BackColor = GetBkColor(GetDC(UserControl.ContainerHwnd))
            Err.Clear
        End If
        Call UserControl_Paint
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels    ' our uc will always be vbPixels & no border
    UserControl.BorderStyle = 0
    UserControl.AutoRedraw = True       ' preference, not required
    m_LockAspectRatio = True            ' force scaled sizing of control
    Set m_Image = New c32bppDIB
End Sub

Private Sub UserControl_Paint()
        
    ' Note: VB does not call this routine when AutoRedraw=True, so we call it
    UserControl.Cls
    If Not m_Image.Handle = 0 Then m_Image.Render UserControl.hDC, , , m_ImgScaleWidth, m_ImgScaleHeight

End Sub

Private Sub UserControl_Resize()

    If m_ImgScaleHeight < 0 Then Exit Sub ' prevent recursion during aspect resizing
    If m_ImgScaleWidth < 0 Then Exit Sub  ' prevent recursion during aspect resizing
    
    If m_Image Is Nothing Then Exit Sub
    If Not m_Image.Handle = 0 Then
    
        ReScaleImage True
        If m_LockAspectRatio = True Then
            If Not m_ImgScaleHeight = UserControl.ScaleHeight Then
                m_ImgScaleHeight = -m_ImgScaleHeight
                UserControl.Extender.Height = Int(ScaleY(-m_ImgScaleHeight, vbPixels, vbContainerSize))
                m_ImgScaleHeight = -m_ImgScaleHeight
            End If
            If Not m_ImgScaleWidth = UserControl.ScaleWidth Then
                m_ImgScaleWidth = -m_ImgScaleWidth
                UserControl.Extender.Width = Int(ScaleX(-m_ImgScaleWidth, vbPixels, vbContainerSize))
                m_ImgScaleWidth = -m_ImgScaleWidth
            End If
        End If
        If Not m_ResImgID = vbNullString Then
            ' we may be able to get a better size from the resource on some image types
            Select Case m_Image.ImageType
            Case imgIcon, imgIconARGB, imgPNGicon, imgWMF, imgEMF, imgCursor, imgCursorARGB
                LoadResPic
            End Select
        End If
        Call UserControl_Paint
    End If
    
End Sub

Private Sub UserControl_Show()
    Call UserControl_Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
            
    Dim imgArray() As Byte, cPNG As cPNGparser
    With PropBag
        .WriteProperty "ResFile", m_ResFile, vbNullString
        .WriteProperty "ResID", m_ResImgID, vbNullString
        .WriteProperty "ResSect", m_ResImgSec, vbNullString
        
        If m_ResImgID = vbNullString Then
            Set cPNG = New cPNGparser
            ' attempt to save image in PNG format (compressed)
            If cPNG.SaveTo(vbNullString, imgArray(), m_Image) = False Then
                m_Image.SaveToStream imgArray() ' fallback if we can't
            End If
            .WriteProperty "Image", imgArray()
        End If
    End With

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Dim imgArray() As Byte
    GetUCstate  ' is host compiled?
    With PropBag
        m_ResFile = .ReadProperty("ResFile", vbNullString)
        m_ResImgID = .ReadProperty("ResID", vbNullString)
        m_ResImgSec = .ReadProperty("ResSect", vbNullString)
        
        ' basically a safety check; does the path to the res file exist?
        If Not m_ResFile = vbNullString Then
            If (m_UCstate And uStHostCompiled) = 0 Then
                If Len(Dir$(m_ResFile, vbHidden Or vbReadOnly Or vbSystem)) = 0 Then m_ResFile = vbNullString
            End If
        End If
    End With
    If m_ResFile = vbNullString Then
        imgArray() = PropBag.ReadProperty("Image", imgArray())
        m_Image.LoadPicture_Stream imgArray
        ReScaleImage True
    Else
        LoadResPic
    End If
    
    On Error Resume Next
    UserControl.BackColor = UserControl.Extender.Container.BackColor
    
End Sub



Private Sub ReScaleImage(bAllowGrowth As Boolean)
    ' function to scale display image as needed in these routines
    If Not m_Image.Handle = 0 Then
        m_Image.ScaleImage UserControl.ScaleWidth, UserControl.ScaleHeight, m_ImgScaleWidth, m_ImgScaleHeight, (scaleDownAsNeeded + CInt(bAllowGrowth))
    End If
End Sub

Private Sub GetUCstate()
    
    ' function returns a set of attributes that can be used to determine how a UC
    ' can access a resource file. Attributes also indicate the state of the UC
    ' and its host container/parent.
    
    Dim oHwnd As Long
    Dim pHwnd As Long
    Dim sClass As String * 20
    Dim nNull As Long
    
    On Error Resume Next
    m_UCstate = 0&
    If UserControl.Ambient.UserMode = True Then
        
        m_UCstate = uStHostRunMode ' the uc is in run mode
    
        ' determine if parent is compiled or not.
        ' There are a few ways but none are truly fool-proof except this one for vb6
        pHwnd = UserControl.ContainerHwnd
        If pHwnd = 0 Then
            ' windowless UC in another windowless UC/container?
            pHwnd = UserControl.Parent.hwnd
            If Err Then Err.Clear
        End If
        Do Until pHwnd = 0
            oHwnd = pHwnd
            pHwnd = GetParent(oHwnd)
        Loop
        If Not oHwnd = 0 Then
            oHwnd = GetWindow(oHwnd, GW_OWNER)
            If Not oHwnd Then
                nNull = GetClassName(oHwnd, sClass, 20&)
                ' vb6. When run mode, uncompiled, the project's true owner is a ThunderMain class. Is this true for vb5?
                If Not LCase(Left$(sClass, nNull)) = "thundermain" Then m_UCstate = m_UCstate Or uStHostCompiled
            End If
        End If
    
    End If
    
    ' is the uc compiled?
    Debug.Print 1 / 0
    If Err Then
        Err.Clear   ' nope
    Else
        m_UCstate = m_UCstate Or uStUCcompiled  ' yep
    End If
    ' now is the uc in host project or is it in its own project?
    If App.StartMode = 0 Then m_UCstate = m_UCstate Or uStUCinHost
    
End Sub

' Friend properties for property page to access usercontrol properties
' These are only used by the property page
Friend Property Get ResourceFileName() As String
    ResourceFileName = m_ResFile
End Property
Friend Property Let ResourceFileName(FName As String)
    m_ResFile = FName
    LoadResPic
    Call UserControl_Resize
    PropertyChanged "ResFile"       ' user selected a res file, save file name
End Property
Friend Property Get ImageResSection() As String
    ImageResSection = m_ResImgSec
End Property
Friend Property Let ImageResSection(Section As String)
    m_ResImgSec = Section           ' user selected image from resource file?
    If Section = vbNullString Then Me.ImageResID = Section ' or from elsewhere?
    PropertyChanged "ResSect"       ' Note: this property does not trigger a redraw, the ImageResID does
End Property
Friend Property Get ImageResID() As String
    ImageResID = m_ResImgID
End Property
Friend Property Let ImageResID(ID As String)
    m_ResImgID = ID                 ' user selected image from resource file?
    If ID = vbNullString Then
        m_ResImgSec = ID            ' or from elsewhere?
        m_Image.DestroyDIB
        Call UserControl_Paint
    Else
        Call LoadResPic             ' get the image from the resource
        Call UserControl_Resize     ' size control to new image, if any, & redraw
    End If
    PropertyChanged "ResID"
End Property
Friend Property Get Image() As c32bppDIB
    Set Image = m_Image
End Property
Friend Property Set Image(NewDIBclass As c32bppDIB)
    If NewDIBclass Is Nothing Then
        If Not m_Image Is Nothing Then m_Image.DestroyDIB
        Me.ImageResID = vbNullString    ' this call will  destroy dib & erase DC
    Else
        m_ResFile = vbNullString
        m_ResImgSec = vbNullString
        NewDIBclass.CopyImageTo m_Image
        Call UserControl_Resize         ' repaint
    End If
End Property

Private Sub LoadResPic()
    
    If (m_UCstate And uStHostCompiled) = uStHostCompiled Then
        m_ResFile = vbNullString    ' image comes from the compiled host, we don't need this any longer
    Else
        If m_ResFile = vbNullString Then Exit Sub
    End If
            
    Dim imgArray() As Byte, tPic As StdPicture
    Dim cR As New cRESreader
    If cR.ScanResources(m_ResFile) = True Then
        ' get the image from resource as array or stdPic
        If cR.ExtractResourceItem(m_ResImgSec, m_ResImgID, imgArray, tPic) = True Then
            If tPic Is Nothing Then
                m_Image.LoadPicture_Stream imgArray(), UserControl.ScaleWidth, UserControl.ScaleHeight
            Else
                m_Image.LoadPicture_StdPicture tPic
            End If
            ReScaleImage True
        End If
    End If

End Sub


' these properties are for reference only during IDE
Public Property Let ImageWidth_Actual(Width As Long)
End Property
Public Property Get ImageWidth_Actual() As Long
    If Not m_Image Is Nothing Then ImageWidth_Actual = m_Image.Width
End Property

Public Property Let ImageHeight_Actual(Height As Long)
End Property
Public Property Get ImageHeight_Actual() As Long
    If Not m_Image Is Nothing Then ImageHeight_Actual = m_Image.Height
End Property

