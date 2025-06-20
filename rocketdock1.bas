Attribute VB_Name = "mdlMain"
' .76 DAEB 28/05/2022 rdIconConfigForm.frm New font code synchronising method with FCW fixing tool not displaying previously chosen font ENDS

Option Explicit
'------------------------------------------------------------
' module1
'
' main APIs, constants defined for querying the registry
' some global variables and a few local subroutines/functions
' pertaining to the main form.
'
'------------------------------------------------------------

Private Const IID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
Private Const LR_LOADFROMFILE As Long = &H10
Private Const DI_NORMAL = 3

'Private Const KEY_QUERY_VALUE = &H1
'Private Const KEY_READ = &H20019
'Private Const KEY_WOW64_64KEY As Long = &H100&

Private Type PictDesc
    cbSizeofStruct  As Long
    PicType         As Long
    hImage          As Long
    xExt            As Long
    yExt            As Long
End Type

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Enum OLE_ERROR_CODES
    S_OK = 0
    E_NOINTERFACE = &H80004002
    E_POINTER = &H80004003
    E_FAIL = &H80004005
    E_UNEXPECTED = &H8000FFFF
    E_INVALIDARG = &H80070057
End Enum
  
Public debugFlg As Integer
Public fileIconListPosition As Integer
Public rdIconNumber As Integer

Public gblIcoSizePreset As Integer
Public thumbArray() As Integer
Public rdIconMaximum As Integer
Public theCount As Integer
Public picFrameThumbsGotFocus As Boolean
Public vScrollThumbsGotFocus As Boolean
Public picRdMapGotFocus As Boolean
Public keyPressOccurred As Boolean
Public previewFrameGotFocus As Boolean
Public filesIconListGotFocus As Boolean
Public gblBaseThumbImageSize As Long
Public storeLeft As Long
Public gblStoredIndex As Integer
Public glLargeIcons() As Long
Public glSmallIcons() As Long
Public lIcons         As Long
Public relativePath As String
Public dotCount As Integer
Public iconChanged As Boolean
Public gblBoxSpacing As Integer
Public gblBusyCounter As Integer

Public thumbIndexNo As Integer
Public gblThumbnailStartPosition As Integer
Public refreshThumbnailView As Boolean
Public displayHourglass As Boolean
Public triggerStartCalc As Boolean
Public triggerRdMapRefresh As Boolean
Public gblClassicTheme As Boolean
Public gblStoreThemeColour As Long

Public CTRL_1 As Boolean
Public CTRL_2 As Boolean
Public captureIconCount As Integer      ' allow the icon count to be accessible to the rest of the program

' .54 DAEB 25/04/2022 rDIConConfig.frm Added rDThumbImageSize saved variable to allow the tool to open the thumbnail explorer in small or large mode
Public rDThumbImageSize As String
Public sFilenameCheck As String ' debug

Public rDIconConfigFormXPosTwips As String
Public rDIconConfigFormYPosTwips As String

'Public sdAppPath As String
'Public steamyDockInstalled As Boolean
'Public SDinstalled As String
'Public SD86installed As String
'Public dockAppPath As String
'Public defaultDock As Integer

'Public rDLockIcons As String
'Public rDOpenRunning As String
'Public rDShowRunning As String
'Public rDManageWindows As String
'Public rDDisableMinAnimation As String

' APIs for querying the registry START
'Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByRef lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, ByRef lpcbClass As Long, ByRef lpftLastWriteTime As FILETIME) As Long
' APIs for querying the registry ENDS

' APIs for drawing icons START
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long

Private Declare Function Ole_CreatePic Lib "olepro32" _
                Alias "OleCreatePictureIndirect" ( _
                ByRef lpPictDesc As PictDesc, _
                ByVal riid As Long, _
                ByVal fPictureOwnsHandle As Long, _
                ByRef iPic As IPicture _
) As Long

Private Declare Function OLE_CLSIDFromString Lib "ole32" Alias "CLSIDFromString" (ByVal lpszProgID As Long, ByVal pCLSID As Long) As Long
Private Declare Function PrivateExtractIcons Lib "user32" _
                Alias "PrivateExtractIconsA" ( _
                ByVal lpszFile As String, _
                ByVal nIconIndex As Long, _
                ByVal cxIcon As Long, _
                ByVal cyIcon As Long, _
                ByRef phIcon As Long, _
                ByRef pIconId As Long, _
                ByVal nIcons As Long, _
                ByVal flags As Long _
) As Long

Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As Any, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus.dll" (ByVal hbm As Long, ByRef pBitMap As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "GdiPlus.dll" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal Callback As Long, ByVal CallbackData As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Interpolation As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As Long, ByRef clsidEncoder As Any, ByRef encoderParams As Any) As Long

Private Const InterpolationModeHighQualityBicubic As Long = &H7&

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

' APIs for drawing icons END

Public rDMonitor      As String

Public origWidth As Long
Public origHeight As Long
Public gblRdEnableBalloonTooltips As Boolean

Public picFrameThumbsLostFocus As Boolean
Public thisRoutine As String
Public lastHighlightedRdMapIndex As Integer

Public srcDragControl As String
Public thumbnailDragTimerCounter As Long
Public rdMapDragTimerCounter As Long
Public picThumbIconMouseDown As Boolean
Public rdMapIconMouseDown As Boolean

Public srcRdIconNumber As Integer
Public trgtRdIconNumber As Integer
Public rdMapIconSrcIndex As Integer

Public SDSuppliedFont As String
Public SDSuppliedFontSize As String
Public SDSuppliedFontItalics As String
Public SDSuppliedFontColour As String
'Public SDSuppliedFontStrength As String
'Public SDSuppliedFontStyle As String


' .76 DAEB 28/05/2022 rdIconConfigForm.frm New font code synchronising method with FCW fixing tool not displaying previously chosen font ENDS
'------------------------------------------------------ STARTS
'constants used to choose a font via the system dialog window
Public Const LOGPIXELSY As Integer = 90
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const LF_FACESIZE As Integer = 32
Private Const CF_EFFECTS  As Long = &H100&
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40&
Private Const CF_SCREENFONTS As Long = &H1

'type declaration used to choose a font via the system dialog window
Public Type FormFontInfo
  Name As String
  Weight As Integer
  Height As Integer
  UnderLine As Boolean
  Italic As Boolean
  Color As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
  lStructSize As Long
  hWnd As Long
  hDC As Long
  lpLogFont As Long
  iPointSize As Long
  flags As Long
  rgbColors As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  hInstance As Long
  lpszStyle As String
  nFontType As Integer
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long
  nSizeMax As Long
End Type

Private Type ChooseColorStruct
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'APIs used to choose a font via the system dialog window
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" _
(pChoosefont As FONTSTRUC) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" _
  (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetDeviceCaps Lib "gdi32" _
  (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
' .76 DAEB 28/05/2022 rdIconConfigForm.frm New font code synchronising method with FCW fixing tool not displaying previously chosen font ENDS

''------------------------------------------------------ STARTS
'' Constants for playing sounds
'Public Const SND_ASYNC As Long = &H1         '  play asynchronously
'Public Const SND_FILENAME  As Long = &H20000     '  name is a file name
'
'' APIs for playing sounds
'Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
''------------------------------------------------------ ENDS

' TBD DAEB 19/04/2021 mdlMain.bas  added a new type link for determining shortcuts
Public Type Link
    Attributes As Long
    FileName As String
    Description As String
    RelPath As String
    WorkingDir As String
    Arguments As String
    CustomIcon As String
End Type

' .91 DAEB 25/06/2022 rDIConConfig.frm Deleting an icon from the icon thumbnail display causes a cache imageList error. Added gblCacheingFlg.
Public gblCacheingFlg As Boolean

Public sdChkToggleDialogs As String ' .70 DAEB 16/05/2022 rDIConConfig.frm Read the chkToggleDialogs value from a file and save the value for next time

Public origSettingsFile As String

Public interimSettingsFile As String

Public gblProgramStatus As String

'------------------------------------------------------ STARTS
' Private Types for determining  sizing
Public gblResizeRatio As Double
Public gblFormResizedInCode As Boolean
'Public gblDoNotResize As Boolean

Public gblStartFormHeight As Long
Public gblStartFormWidth  As Long

Public rDIconConfigFormOldHeight As Long
Public rDIconConfigFormOldWidth As Long

Public gblFormPrimaryHeightTwips As String

Public startupFlg As Boolean
            
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
            (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
            
'Public lngGDI As Long
'Public gdipInit As GDIPLUS_STARTINPUT
'Private Type GDIPLUS_STARTINPUT
'    GdiplusVersion As Long
'    DebugEventCallback As Long
'    SuppressBackgroundThread As Long
'    SuppressExternalCodecs As Long
'End Type
'
'Public Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal mtoken As Long)
'Public Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, gdipInput As GDIPLUS_STARTINPUT, GdiplusStartupOutput As Long) As Long
'Public Type GdiplusStartupInput
'    GdiplusVersion As Long
'    DebugEventCallback As Long
'    SuppressBackgroundThread As Long
'    SuppressExternalCodecs As Long
'End Type
'Public Type GdiplusStartupOutput
'    NotificationHook As Long
'    NotificationUnhook As Long
'End Type
'
'' NOTE: Enums evaluate to a Long
'Public Enum GpStatus   ' aka Status
'   Ok = 0
'   GenericError = 1
'   InvalidParameter = 2
'   OutOfMemory = 3
'   ObjectBusy = 4
'   InsufficientBuffer = 5
'   NotImplemented = 6
'   Win32Error = 7
'   WrongState = 8
'   Aborted = 9
'   FileNotFound = 10
'   ValueOverflow = 11
'   AccessDenied = 12
'   UnknownImageFormat = 13
'   FontFamilyNotFound = 14
'   FontStyleNotFound = 15
'   NotTrueTypeFont = 16
'   UnsupportedGdiplusVersion = 17
'   GdiplusNotInitialized = 18
'   PropertyNotFound = 19
'   PropertyNotSupported = 20
'   ProfileNotFound = 21
'End Enum
    
    
    

'------------------------------------------------------ ENDS


'---------------------------------------------------------------------------------------
' Procedure : displayEmbeddedIcons
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : The program extracts icons embedded within a DLL or an executable
'             you pass the name of the picbox you require and the image is displayed there
'             it should return all and not only the 16 and 32 bit icons as does extractIconEx
'
'             I may not have coded this particularly well - but it works.
'---------------------------------------------------------------------------------------
'
'
Public Sub displayEmbeddedIcons(ByVal FileName As String, ByRef targetPicBox As PictureBox, ByVal IconSize As Integer)
    
    Dim lIconIndex As Long: lIconIndex = 0
    Dim xSize As Long: xSize = 0
    Dim ySize As Long: ySize = 0
    Dim hIcon() As Long

    Dim hIconID() As Long
    Dim nIcons As Long: nIcons = 0
    Dim Result As Long: Result = 0
    Dim flags As Long: flags = 0
    Dim i As Long: i = 0
    
    'Dim pic As IPicture ' interface for a Picture object
    Dim pic As StdPicture ' interface for a Picture object
    
    Dim cICO As New cICOparser  ' cICOParser.cls
    Dim cImage As c32bppDIB  ' c32bppDIB.cls
        
    Dim outputFilename As String: outputFilename = "arse1.png"
    Dim outputFilenameICO As String: outputFilenameICO = "arse1.ico"
    
    Dim GSI As GdiplusStartupInput
    Dim hToken As Long
    Dim uEncCLSID(0 To 3) As Long
    
    Dim SaveToPNG As Boolean

    On Error GoTo displayEmbeddedIcons_Error
    
    GSI.GdiplusVersion = 1
    GdiplusStartup hToken, GSI

    On Error Resume Next

    lIconIndex = 0
    i = 2 ' need some experimentation here
    
    'the boundaries of the icons you wish to extract packed into a 32bit LONG for an API call
    xSize = make32BitLong(CInt("256"), CInt("16")) ' 1048832
    ySize = make32BitLong(CInt("256"), CInt("16")) ' 1048832
    
    ' flags
    '
    '    LR_DEFAULTCOLOR
    '    LR_CREATEDIBSECTION
    '    LR_DEFAULTSIZE
    '    LR_LOADFROMFILE
    '    LR_LfsOADMAP3DCOLORS
    '    LR_LOADTRANSPARENT
    '    LR_MONOCHROME
    '    LR_SHARED
    '    LR_VGACOLOR
    '
    flags = LR_LOADFROMFILE '16

    ' Call PrivateExtractIcons with the 5th param set to nothing solely to obtain the total number of Icons in the file.
    Result = PrivateExtractIcons(FileName, lIconIndex, xSize, ySize, ByVal 0&, ByVal 0&, 0&, 0&) ' 63
    
    ' The Filename is the resource string/filepath.
    ' lIconIndex is the index.
    ' xSize and ySize are the desired sizes.
    ' 5th parameter is a pointer to the returned array of icon handles.
    ' piconid is an ID of each icon that best fits the current display device. The returned identifier is 0 if not obtained.
    ' nicons is the number of icons you wish to extract.
    
    ' If you call it with nicon set to this number and niconindex=0 it will extract ALL your icons in one go.
    
    ' eg. PrivateExtractIcons ('C:\Users\Public\Documents\RAD Studio\Projects\2010\Aero Colorizer\AeroColorizer.exe', 0, 128, 128, @hIcon, @nIconId, 1, LR_LOADFROMFILE)
    ' PrivateExtractIcons(sExeName, lIconIndex, xSize, ySize,  hIcon(LBound(hIcon)), hIconID(LBound(hIconID)), nIcons * 2, LR_LOADFROMFILE)

    nIcons = Result ' 63
    
    ' Dimension the arrays to the number of icons.
    ReDim hIcon(lIconIndex To lIconIndex + nIcons * 2 - 1)
    ReDim hIconID(lIconIndex To lIconIndex + nIcons * 2 - 1)

    ' use the undocumented PrivateExtractIcons to extract the icons we require where the 5th param is a pointer to the returned array of handles to extracted icons
    Result = PrivateExtractIcons(FileName, lIconIndex, xSize, _
                            ySize, hIcon(LBound(hIcon)), _
                            hIconID(LBound(hIconID)), _
                            nIcons * 2, flags)
        
    ' create an Ipicture icon with a handle, no specific size
    Set pic = CreateIcon(hIcon(i + lIconIndex - 1))
    ' do we need the above in order to display the icon on the targe picbox? No.
    ' we do this as a check as to a valid pic before we write directly to the targetPicBox
    ' also, we now create a DIB from the pic with the hope that I can use that to extract the image and save an .ICO file
    
    ' BEGIN next bit might be useful, unsure
    
        ' initialise a 32bpp alpha-compatible image bitmap.
        Set cImage = New c32bppDIB ' c32bppDIB.cls
        Dim res As Boolean
        ' pass to icon/cursor parser, convert a single icon from a stdPicture/iPicture or handle to a 32bpp bitmap
        res = cICO.ConvertstdPicTo32bpp(pic.handle, cImage)
        'Call cICO.ConvertstdPicTo32bpp(hIcon(i + lIconIndex - 1), cImage) ' this also seems to work. compiles
        
        Set cICO = Nothing
    
    ' END
'
    ' cImage.handle
    
    ' because we have the pic as ipicture from the above CreateIcon, we should be able to save the icon to a PNG with a bit of luck using the handle to the ICO and GDI+

    
    
' See BUGLIST.TXT FOR A FULL DESCRIPTION OF THE PROGRESS TO THIS POINT.
    
'    next task is to try ICOwriter class, creating an icon and saving


'    Call NewICO(), then AddImage() one or more times, then BuildICO().
'
'    After that you can use the ICO property to get a Byte Array or the
'    SaveICO() method to save it to disk, or both.

'    Dim icoWriter As New icoWriter
'    Call icoWriter.SaveICO(outputFilenameICO, True)



        
    ' resize and place the target picbox according to the size of the icon
    ' (rather than placing the icon in the middle of the picbox as I should, I can code that later)
    
    Call centrePreviewImage(targetPicBox, IconSize)
    
    Dim hGraphics As Long
    Dim hImage As Long
            
    ' Draw the icon directly onto the respective picturebox control.
    If Not (pic Is Nothing) Then
        With targetPicBox
        
            'ensure the picbox is empty first
            'Set .Picture = LoadPicture(vbNullString)
            .Cls
            .AutoRedraw = True
               
            ' DrawIconEx can draw a transparent image from a good image handle directly onto the target picture box
            ' Call DrawIconEx(.hDC, 0, 0, hIcon(LBound(hIcon)), IconSize, IconSize, 0, 0, DI_NORMAL)
            
            ' centre image using a better method
'            Call DrawIconEx(.hDC, _
'                        ScaleX(x, ScaleMode, vbPixels) - WidthPx \ 2, _
'                        ScaleY(y, ScaleMode, vbPixels) - HeightPx \ 2, _
'                        hIcon(LBound(hIcon)), _
'                        IconSize, _
'                        IconSize, _
'                        0, _
'                        WIN32_NULL, _
'               DI_NORMAL)

            'creates an image bitmap from the icon handle
            GdipCreateBitmapFromHICON hIcon(LBound(hIcon)), hImage
            If hImage <> 0& Then
                ' Creates a Graphics object that is associated with a specified device context
                GdipCreateFromHDC .hDC, hGraphics
                
                'Draws an image at a specified location
                '                      hGraphics, hImage, destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight, UnitPixel, hImgAttr, 0&, 0&
                GdipDrawImageRectRectI hGraphics, hImage, 0, 0, IconSize, IconSize, 0, 0, 256, 256, 2&, 0, 0, 0
                    
                    
'                Dim tSI As GdiplusStartupInput
'                Dim lRes As Long, lGDIP As Long, lBitmap As Long
'                Dim X As Long, Y As Long, wide As Long, high As Long
'                Dim myDIB As Long, myDC As Long, desktopDC As Long
'                Dim bi24BitInfo As BITMAPINFO
'                Dim bitmapData() As Byte
'                Dim pcin As PCURSORINFO
'                Dim piinfo As ICONINFO
'                ' Starting position/Size of capture (full screen)
'                X = 0: Y = 0
'                wide = Screen.Width / Screen.TwipsPerPixelX
'                high = Screen.Height / Screen.TwipsPerPixelY
'                '
'                With bi24BitInfo.bmiHeader
'                  .biBitCount = 24
'                  .biCompression = BI_RGB
'                  .biPlanes = 1
'                  .biSize = Len(bi24BitInfo.bmiHeader)
'                  .biWidth = wide
'                  .biHeight = high
'                  .biDataSize = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
'                  ReDim bitmapData(0 To .biDataSize - 1)
'                End With
'                frmscrcontrol.Caption = UBound(bitmapData)
'                myDC = CreateCompatibleDC(0)
'                myDIB = CreateDIBSection(myDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
'                SelectObject myDC, myDIB
'                desktopDC = GetDC(0)
'                BitBlt myDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, desktopDC, X, Y, vbSrcCopy Or CAPTUREBLT
'                ' Include mouse cursor?
'                If IncludeMouseCursor = True Then
'                    pcin.cbSize = Len(pcin)
'                    GetCursorInfo pcin
'                    GetIconInfo pcin.hCursor, piinfo
'                    DrawIcon myDC, pcin.ptScreenPos.X - piinfo.xHotspot, pcin.ptScreenPos.Y - piinfo.yHotspot, pcin.hCursor
'                    If piinfo.hbmMask Then DeleteObject piinfo.hbmMask
'                    If piinfo.hbmColor Then DeleteObject piinfo.hbmColor
'                End If
'                Call GetDIBits(myDC, myDIB, 0, bi24BitInfo.bmiHeader.biHeight, bitmapData(0), bi24BitInfo, DIB_RGB_COLORS)
'
'
'
'               ' save as JPG
'               '------------
'               'Initialize GDI+
'               tSI.GdiplusVersion = 1
'                lRes = GdiplusStartup(lGDIP, tSI)
'                ' Create the GDI+ bitmap from the image handle
'                lRes = GdipCreateBitmapFromHBITMAP(myDIB, 0, lBitmap)
'                If lRes = 0 Then
'                   Dim tJpgEncoder As GUID
'                   Dim tParams As EncoderParameters
'                   ' Initialize the encoder GUID
'                   CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
'                   ' Initialize the encoder parameters
'                   tParams.Count = 1
'                   With tParams.Parameter ' Quality
'                      ' Set the Quality GUID
'                      CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
'                      .NumberOfValues = 1
'                      .Type = 4
'                      .Value = VarPtr(Quality)
'                   End With
'                   ' Save the image
'                   lRes = GdipSaveImageToFile(lBitmap, StrPtr(FileName), tJpgEncoder, tParams)
'                   ' Destroy the bitmap
'                   GdipDisposeImage lBitmap
'                End If
                SaveToPNG = (GdipSaveImageToFile(hImage, StrPtr(outputFilename), uEncCLSID(0&), ByVal 0&) = 0&)

                GdipDeleteGraphics hGraphics
                GdipDisposeImage hImage: hImage = 0&
            End If
            
            .Refresh

        End With
        
        'save the target box contents as an ico, no alpha channel
        'SavePicture targetPicBox.Image, outputFilenameICO
        
        ' using Dil's picSave class, this will save a PNG into the target app folder but it is non-alpha blend as it is straight from the target picbox
        'Call PicSave.SavePicture(targetPicBox.Image, outputFilename, fmtPNG, 70) ' uses picSave.cls
    End If
    
    ' get rid of the icons we created
    Call DestroyIcon(hIcon(i + lIconIndex - 1))
    
    GdiplusShutdown hToken

   On Error GoTo 0
   Exit Sub

displayEmbeddedIcons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayEmbeddedIcons of Module mdlMain"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : make32BitLong
' Author    : beededea
' Date      : 20/11/2019
' Purpose   : packing variables into a 32bit LONG for an API call
'---------------------------------------------------------------------------------------
'
Private Function make32BitLong(ByVal LoWord As Integer, _
                 Optional ByVal HiWord As Integer = 0) As Long
   On Error GoTo make32BitLong_Error
   If debugFlg = 1 Then debugLog "%make32BitLong"

    make32BitLong = CLng(HiWord) * CLng(&H10000) + CLng(LoWord)

   On Error GoTo 0
   Exit Function

make32BitLong_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure make32BitLong of Module Module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateIcon
' Author    : beededea
' Date      : 14/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function CreateIcon(ByVal hImage As Long) As IPicture
    ' This method creates an icon based on a handle
    Dim pic As IPicture
    Dim dsc As PictDesc
    Dim IID(0 To 15) As Byte
    Dim Result As Long: Result = 0
    
   On Error GoTo CreateIcon_Error

    Set CreateIcon = Nothing
    If hImage <> 0 Then
        With dsc
           .cbSizeofStruct = Len(dsc)
           .hImage = hImage
           .PicType = VBRUN.PictureTypeConstants.vbPicTypeBitmap
        End With
        
        Result = OLE_CLSIDFromString(StrPtr(IID_IPicture), _
                                                        VarPtr(IID(0)))
                                                    
        If (Result = OLE_ERROR_CODES.S_OK) Then
            Result = Ole_CreatePic(dsc, VarPtr(IID(0)), True, pic)
            
            If (Result = OLE_ERROR_CODES.S_OK) Then
                Set CreateIcon = pic
            End If
        End If
    End If

   On Error GoTo 0
   Exit Function

CreateIcon_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateIcon of Module Module1"
End Function

'---------------------------------------------------------------------------------------
' Procedure : displayEmbeddedIconsOld
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : The old method of extracting embedded icons from DLLs and EXEs
'             Retained for informational purposes
'---------------------------------------------------------------------------------------
'
Public Sub displayEmbeddedIconsOld(ByVal FileName As String, ByRef targetPicBox As PictureBox, ByRef IconSize As Integer)
    ' The program extracts icons embedded within a DLL or an executable
    ' you pass the name of the picbox you require and the image is displayed there
    ' unfortunately the ExtractIconEx API only returns 16 and 32 bit icons
    
    Dim sExeName       As String
    Dim lIndex         As Long: lIndex = 0

' eg. FileName = "C:\Program Files (x86)\Microsoft Visual Studio 8\Common7\IDE\vbexpress.exe"
   On Error GoTo displayEmbeddedIcons_Error

    sExeName = FileName

' Get the total number of Icons in the file.
    lIcons = ExtractIconEx(sExeName, -1, 0, 0, 0)

' Dimension the arrays to the number of icons.
    ReDim glLargeIcons(lIcons)
    ReDim glSmallIcons(lIcons)

    lIndex = 0

' Get the handle of the icon indicated by lIndex, in this case the smallest
' this API can only get a maximum of two embedded icons
    Call ExtractIconEx(sExeName, lIndex, glLargeIcons(lIndex), glSmallIcons(lIndex), 1)

    Call centrePreviewImage(targetPicBox, IconSize)
    
    With targetPicBox
        Set .Picture = LoadPicture(vbNullString)
        .AutoRedraw = True
           
        Call DrawIconEx(.hDC, 0, 0, glLargeIcons(lIndex), IconSize, IconSize, 0, 0, DI_NORMAL)
            
        .Refresh
    End With

   On Error GoTo 0
   Exit Sub

displayEmbeddedIcons_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayEmbeddedIcons of Module Module1"
 
End Sub



'---------------------------------------------------------------------------------------
' Procedure : centrePreviewImage
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : place the image correctly within the preview pane
'---------------------------------------------------------------------------------------
' because the icon images are drawn from the top left of the
' preview pictureBox we have to manually set the picbox to size and position for each icon size
' this could be done with padding but it matches the VB6 method (no padding there)
Public Sub centrePreviewImage(ByRef targetPicBox As PictureBox, ByVal IconSize As Integer)

    If targetPicBox.Name = "picPreview" Then
        If gblIcoSizePreset = 16 Then
            targetPicBox.Left = (1900 * gblResizeRatio)
            targetPicBox.Top = (1900 * gblResizeRatio)
            targetPicBox.Width = (200 * gblResizeRatio)
            targetPicBox.Height = (200 * gblResizeRatio)
        ElseIf gblIcoSizePreset = 32 Then
            targetPicBox.Left = (1800 * gblResizeRatio)
            targetPicBox.Top = (1800 * gblResizeRatio)
            targetPicBox.Width = (2000 * gblResizeRatio)
            targetPicBox.Height = (2000 * gblResizeRatio)
        ElseIf gblIcoSizePreset = 64 Then
            targetPicBox.Left = (1450 * gblResizeRatio)
            targetPicBox.Top = (1450 * gblResizeRatio)
            targetPicBox.Width = (2000 * gblResizeRatio)
            targetPicBox.Height = (2000 * gblResizeRatio)
        ElseIf gblIcoSizePreset = 128 Then
            targetPicBox.Left = (1000 * gblResizeRatio)
            targetPicBox.Top = (1000 * gblResizeRatio)
            targetPicBox.Width = (2000 * gblResizeRatio)
            targetPicBox.Height = (2000 * gblResizeRatio)
        ElseIf gblIcoSizePreset = 256 Then
            targetPicBox.Left = (100 * gblResizeRatio)
            targetPicBox.Top = (100 * gblResizeRatio)
            targetPicBox.Width = (4000 * gblResizeRatio)
            targetPicBox.Height = (4000 * gblResizeRatio)
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : chkTheRegistry
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : if the original settings.ini exist then RD is not using the registry
'---------------------------------------------------------------------------------------
'
Public Sub chkTheRegistry()

    On Error GoTo chkTheRegistry_Error
    'If debugFlg = 1 Then debugLog  "%" & "chkTheRegistry"

    'frmRegistry.fraReadConfig.Enabled = True
    'frmRegistry.fraWriteConfig.Enabled = True
    
'    If rocketDockInstalled = True And defaultDock = 0 Then
'
'        If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
'            frmRegistry.chkReadRegistry.Value = 0
'            frmRegistry.chkReadSettings.Value = 1
'            frmRegistry.chkReadConfig.Value = 0
'
'            frmRegistry.chkWriteRegistry.Value = 0
'            frmRegistry.chkWriteSettings.Value = 1
'            frmRegistry.chkWriteConfig.Value = 0
'
'        Else
'            frmRegistry.chkReadRegistry.Value = 1
'            frmRegistry.chkReadSettings.Value = 0
'            frmRegistry.chkReadConfig.Value = 0
'
'            frmRegistry.chkWriteRegistry.Value = 1
'            frmRegistry.chkWriteSettings.Value = 0
'            frmRegistry.chkWriteConfig.Value = 0
'
'        End If
'    End If

    If steamyDockInstalled = True And defaultDock = 1 Then  ' it will always exist even if not used
    
'        If fFExists(interimSettingsFile) Then
'            rDGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", interimSettingsFile)
'            rDGeneralWriteConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralWriteConfig", interimSettingsFile)
'        End If
'        If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
'            frmRegistry.chkReadRegistry.Value = 0
'            frmRegistry.chkReadSettings.Value = 1
'            frmRegistry.chkReadConfig.Value = 0
'
'            frmRegistry.chkWriteRegistry.Value = 0
'            frmRegistry.chkWriteSettings.Value = 1
'            frmRegistry.chkWriteConfig.Value = 0
'        Else
            frmRegistry.chkReadRegistry.Value = 1
            frmRegistry.chkReadSettings.Value = 0
            frmRegistry.chkReadConfig.Value = 0
    
            frmRegistry.chkWriteRegistry.Value = 1
            frmRegistry.chkWriteSettings.Value = 0
            frmRegistry.chkWriteConfig.Value = 0
'        End If
    
        If rDGeneralReadConfig = "True" Then
            frmRegistry.chkReadRegistry.Value = 0
            frmRegistry.chkReadSettings.Value = 0
            frmRegistry.chkReadConfig.Value = 1
        End If
        If rDGeneralWriteConfig = "True" Then
            frmRegistry.chkWriteRegistry.Value = 0
            frmRegistry.chkWriteSettings.Value = 0
            frmRegistry.chkWriteConfig.Value = 1
        End If
    End If
    
'
'    frmRegistry.fraReadConfig.Enabled = False
'    frmRegistry.fraWriteConfig.Enabled = False

    
   On Error GoTo 0
   Exit Sub

chkTheRegistry_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure chkTheRegistry of Form rDIconConfigForm"
End Sub



' .76 DAEB 28/05/2022 rDIConConfig.frm New font code synchronising method with FCW fixing tool not displaying previously chosen font STARTS
'---------------------------------------------------------------------------------------
' Procedure : displayFontSelector
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : select a font for the supplied form
'---------------------------------------------------------------------------------------
'
Private Sub displayFontSelector(ByRef currFont As String, ByRef currSize As Integer, ByRef currWeight As Integer, ByRef currStyle As Boolean, ByRef currColour As Long, ByRef currItalics As Boolean, ByRef currUnderline As Boolean, ByRef fontResult As Boolean)

       
    ' variables declared
    Dim thisFont As FormFontInfo
        
    'initialise the dimensioned variables
    'thisFont =
   
   ' On Error GoTo displayFontSelector_Error
   If debugFlg = 1 Then debugLog "%displayFontSelector"

    With thisFont
      .Color = currColour
      .Height = currSize
      .Weight = currWeight
      '400     Font is normal.
      '700     Font is bold.
      .Italic = currItalics
      .UnderLine = currUnderline
      .Name = currFont
    End With
    
    fontResult = fDialogFont(thisFont)
    If fontResult = False Then Exit Sub
    
    ' some fonts have naming problems and the result is an empty font name field on the font selector
    If thisFont.Name = vbNullString Then thisFont.Name = "times new roman"
    If thisFont.Name = vbNullString Then Exit Sub
    
    With thisFont
        currFont = .Name
        currSize = .Height
        currWeight = .Weight
        currItalics = .Italic
        currUnderline = .UnderLine
        currColour = .Color
        'ctl = .Name & " - Size:" & .Height
    End With

   On Error GoTo 0
   Exit Sub

displayFontSelector_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure displayFontSelector of Form dockSettings"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : fDialogFont
' Author    : beededea
' Date      : 21/08/2020
' Purpose   : display the default windows dialog box that allows the user to select a font
'---------------------------------------------------------------------------------------
'
Public Function fDialogFont(ByRef F As FormFontInfo) As Boolean
      
    ' variables declared
    Dim logFnt As LOGFONT
    Dim ftStruc As FONTSTRUC
    Dim lLogFontAddress As Long: lLogFontAddress = 0
    Dim lMemHandle As Long: lMemHandle = 0
    Dim hWndAccessApp As Long: hWndAccessApp = 0
    
     On Error GoTo fDialogFont_Error
    
    logFnt.lfWeight = F.Weight
    logFnt.lfItalic = F.Italic * -1
    logFnt.lfUnderline = F.UnderLine * -1
    logFnt.lfHeight = -fMulDiv(CLng(F.Height), GetDeviceCaps(GetDC(hWndAccessApp), LOGPIXELSY), 72)
    Call StringToByte(F.Name, logFnt.lfFaceName())
    ftStruc.rgbColors = F.Color
    ftStruc.lStructSize = Len(ftStruc)
    
    lMemHandle = GlobalAlloc(GHND, Len(logFnt))
    If lMemHandle = 0 Then
      fDialogFont = False
      Exit Function
    End If

    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
      fDialogFont = False
      Exit Function
    End If
    
    CopyMemory ByVal lLogFontAddress, logFnt, Len(logFnt)
    ftStruc.lpLogFont = lLogFontAddress
    'ftStruc.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    ftStruc.flags = CF_SCREENFONTS Or CF_INITTOLOGFONTSTRUCT
    If ChooseFont(ftStruc) = 1 Then
      CopyMemory logFnt, ByVal lLogFontAddress, Len(logFnt)
      F.Weight = logFnt.lfWeight
      F.Italic = CBool(logFnt.lfItalic)
      F.UnderLine = CBool(logFnt.lfUnderline)
      F.Name = fByteToString(logFnt.lfFaceName())
      F.Height = CLng(ftStruc.iPointSize / 10)
      F.Color = ftStruc.rgbColors
      fDialogFont = True
    Else
      fDialogFont = False
    End If

   On Error GoTo 0
   Exit Function

fDialogFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fDialogFont of Module modCommon"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fMulDiv
' Author    :
' Date      : 21/08/2020
' Purpose   :  fMulDiv function multiplies two 32-bit values and then divides the 64-bit result by a third 32-bit value.
'---------------------------------------------------------------------------------------
'
Private Function fMulDiv(ByVal In1 As Long, ByVal In2 As Long, ByVal In3 As Long) As Long
    
    ' variables declared
    Dim lngTemp As Long: lngTemp = 0
    
    On Error GoTo fMulDiv_Error

    On Error GoTo fMulDiv_err
    If In3 <> 0 Then
      lngTemp = In1 * In2
      lngTemp = lngTemp / In3
    Else
      lngTemp = -1
    End If
fMulDiv_end:
      fMulDiv = lngTemp
      Exit Function
fMulDiv_err:
      lngTemp = -1
      Resume fMulDiv_err

   On Error GoTo 0
   Exit Function

fMulDiv_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fMulDiv of Module modCommon"
End Function



'---------------------------------------------------------------------------------------
' Procedure : StringToByte
' Author    : beededea
' Date      : 21/08/2020
' Purpose   : convert a provided string to a byte array
'---------------------------------------------------------------------------------------
'
Private Sub StringToByte(ByVal InString As String, ByRef ByteArray() As Byte)
    
    ' variables declared
    Dim intLbound As Integer: intLbound = 0
    Dim intUbound As Integer: intUbound = 0
    Dim intLen As Integer: intLen = 0
    Dim intX As Integer: intX = 0
    
    On Error GoTo StringToByte_Error

    intLbound = LBound(ByteArray)
    intUbound = UBound(ByteArray)
    intLen = Len(InString)
    If intLen > intUbound - intLbound Then intLen = intUbound - intLbound
    For intX = 1 To intLen
        ByteArray(intX - 1 + intLbound) = Asc(Mid$(InString, intX, 1))
    Next

   On Error GoTo 0
   Exit Sub

StringToByte_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StringToByte of Module modCommon"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fByteToString
' Author    :
' Date      : 21/08/2020
' Purpose   : convert a byte array provided to a string
'---------------------------------------------------------------------------------------
'
Private Function fByteToString(ByRef aBytes() As Byte) As String
      
    ' variables declared
    Dim dwBytePoint As Long: dwBytePoint = 0
    Dim dwByteVal As Long: dwByteVal = 0
    Dim szOut As String: szOut = vbNullString
    
    On Error GoTo fByteToString_Error

    dwBytePoint = LBound(aBytes)
    While dwBytePoint <= UBound(aBytes)
      dwByteVal = aBytes(dwBytePoint)
      If dwByteVal = 0 Then
        fByteToString = szOut
        Exit Function
      Else
        szOut = szOut & Chr$(dwByteVal)
      End If
      dwBytePoint = dwBytePoint + 1
    Wend
    fByteToString = szOut

   On Error GoTo 0
   Exit Function

fByteToString_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fByteToString of Module modCommon"
End Function


'---------------------------------------------------------------------------------------
' Procedure : changeFont
' Author    : beededea
' Date      : 12/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub changeFont(ByRef formName As Object, ByRef fntNow As Boolean, ByRef fntFont As String, ByRef fntSize As Integer, ByRef fntWeight As Integer, ByRef fntStyle As Boolean, ByRef fntColour As Long, ByRef fntItalics As Boolean, ByRef fntUnderline As Boolean, ByRef fntFontResult As Boolean)
    Dim useloop As Integer: useloop = 0
    Dim Ctrl As Control
    
    On Error GoTo changeFont_Error
    
    If debugFlg = 1 Then debugLog "%" & "changeFont"
    
    If fntNow = True Then
        displayFontSelector fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult
        If fntFontResult = False Then Exit Sub
    End If
          
    ' a method of looping through all the controls and identifying the labels and text boxes
    ' .TBD DAEB 26/05/2022 rdIconConfig.frm Add listboxes to the types handled
    For Each Ctrl In formName.Controls
         If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is ListBox) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
            If Ctrl.Name <> "lblDragCorner" Then
                If fntFont <> "" Then Ctrl.Font.Name = fntFont
                If fntSize > 0 Then Ctrl.Font.Size = fntSize
                Ctrl.Font.Italic = fntItalics
            End If
        End If
    Next
    
    ' .TBD DAEB 26/05/2022 rdIconConfig.frm Added the specifics for the main form to a separate routine
    If formName.Name = "rDIconConfigForm" Then
        Call rdIconConfigSpecificFonts(formName, fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline)
    End If
    
    ' .TBD DAEB 26/05/2022 rdIconConfig.frm Removed the two new forms from the changeFont tool, now called using the first form parameter in changeFont

'    ' .37 DAEB 05/05/2021 rdIconConfig.frm Added the new form to the changeFont tool
'    For Each Ctrl In formSoftwareList.Controls
'         If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ListBox) Then
'           If fntFont <> "" Then Ctrl.Font.Name = fntFont
'           If fntSize > 0 Then Ctrl.Font.Size = fntSize
'            Ctrl.Font.Italic = fntItalics
'        End If
'    Next
'
'    ' .37 DAEB 05/05/2021 rdIconConfig.frm Added the new form to the changeFont tool
'    For Each Ctrl In frmConfirmDock.Controls
'         If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is FileListBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is ComboBox) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Or (TypeOf Ctrl Is ListBox) Then
'           If fntFont <> "" Then Ctrl.Font.Name = fntFont
'           If fntSize > 0 Then Ctrl.Font.Size = fntSize
'           'If suppliedStyle <> "" Then Ctrl.Font.Style = suppliedStyle
'            Ctrl.Font.Italic = fntItalics
'        End If
'    Next
        
   
   On Error GoTo 0
   Exit Sub

changeFont_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeFont of Form rDIconConfigForm"
    
End Sub

' .76 DAEB 28/05/2022 rDIConConfig.frm New font code synchronising method with FCW fixing tool not displaying previously chosen font ENDS



'---------------------------------------------------------------------------------------
' Procedure : rdIconConfigSpecificFonts
' Author    : beededea
' Date      : 27/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub rdIconConfigSpecificFonts(ByRef formName As Object, ByRef fntFont As String, ByRef fntSize As Integer, ByRef fntWeight As Integer, ByRef fntStyle As Boolean, ByRef fntColour As Long, ByRef fntItalics As Boolean, ByRef fntUnderline As Boolean)
    Dim useloop As Integer: useloop = 0
    
    ' change the size of the two labels beneath the preview image
   On Error GoTo rdIconConfigSpecificFonts_Error

    formName.lblFileInfo.Font.Size = 7
    formName.lblWidthHeight.Font.Size = 7
    
    ' change the font size of the large number
    formName.lblRdIconNumber.Font.Name = "Trebuchet MS"
    formName.lblRdIconNumber.Font.Size = 45
    
    ' change the font size of the large blank
    formName.lblBlankText.Font.Name = "Trebuchet MS"
    formName.lblBlankText.Font.Size = 45

    'loop through the 12 dynamic icon thumbnails, they all exist by the time this routine is called
    For useloop = 0 To 11
        formName.picThumbIcon(useloop).Font.Name = fntFont 'array
        If fntSize > 0 Then formName.picThumbIcon(useloop).Font.Size = fntSize 'array
        
        formName.fraThumbLabel(useloop).Font.Name = fntFont 'array
        If fntSize > 0 Then formName.fraThumbLabel(useloop).Font.Size = fntSize 'array
        
        formName.lblThumbName(useloop).Font.Name = fntFont 'array
        If fntSize > 0 Then formName.lblThumbName(useloop).Font.Size = fntSize 'array
    Next useloop
    
    ' then the treeview that is picky about .fontname or .font.name where the others are not.
    formName.folderTreeView.Font.Name = fntFont
    If fntSize > 0 Then formName.folderTreeView.Font.Size = fntSize
    
    ' The comboboxes all autoselect when the font is changed, we need to reset this afterwards to remove highlighting
    
    formName.cmbIconTypesFilter.SelLength = 0
    formName.cmbDefaultDock.SelLength = 0
    formName.cmbRunState.SelLength = 0
    formName.cmbOpenRunning.SelLength = 0
   
    ' after changing the font, sometimes the filelistbox changes height arbitrarily
    formName.filesIconList.Height = 3450

   On Error GoTo 0
   Exit Sub

rdIconConfigSpecificFonts_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure rdIconConfigSpecificFonts of Module mdlMain"
End Sub


' .89 DAEB 13/06/2022 rDIConConfig.frm Moved backup-related private routines to modules to make them public
'---------------------------------------------------------------------------------------
' Procedure : backupDockSettings
' Author    : beededea
' Date      : 13/06/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub backupDockSettings(Optional ByVal askQuestion As Boolean = False)
    Dim ans As VbMsgBoxResult: ans = vbNo
    Dim iconPath As String: iconPath = vbNullString
    Dim dllPath As String: dllPath = vbNullString
    Dim dialogInitDir As String: dialogInitDir = vbNullString
    Dim bkpSettingsFile As String: bkpSettingsFile = vbNullString
    Dim bkpFilename As String: bkpFilename = vbNullString
    
    Const x_MaxBuffer = 256
    
    On Error GoTo backupDockSettings_Error

    If debugFlg = 1 Then debugLog "%" & "btnBackup_Click"

    bkpFilename = fbackupSettings()
    If askQuestion = True Then
        ans = msgBoxA("Created an incremental backup of the Dock settings file - " & vbCr & vbCr & bkpFilename & vbCr & vbCr & "Would you like to review ALL the backup files? ", vbQuestion + vbYesNo, "Backing up settings.", False, "none")
        If ans = 6 Then
    
            On Error Resume Next
    
            ' set the default folder to the existing reference
            If fDirExists(App.Path & "\backup") Then
                ' set the default folder to the existing reference
                dialogInitDir = App.Path & "\backup" 'start dir, might be "C:\" or so also
            Else
                MsgBox "Backup folder " & App.Path & "\backup" & " has been removed. Backup cancelled"
                Exit Sub
            End If
    
            With x_OpenFilename
            '    .hwndOwner = Me.hWnd
            .hInstance = App.hInstance
            .lpstrTitle = "Select a backup INI file to restore - or cancel"
            .lpstrInitialDir = dialogInitDir
    
            .lpstrFilter = "Ini Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
            .nFilterIndex = 2
    
            .lpstrFile = String$(x_MaxBuffer, 0)
            .nMaxFile = x_MaxBuffer - 1
            .lpstrFileTitle = .lpstrFile
            .nMaxFileTitle = x_MaxBuffer - 1
            .lStructSize = Len(x_OpenFilename)
            End With
            
            Dim retFileName As String
            Dim retfileTitle As String
            Call getFileNameAndTitle(retFileName, retfileTitle)
            bkpSettingsFile = retFileName
            
            If Not bkpSettingsFile = "" Then
            
                ans = msgBoxA("Do you wish to restore this file?  " & bkpSettingsFile & "? ", vbQuestion + vbYesNo, "Restore a backup")
                If ans = 6 Then
                    ' take the backup file and copy it into the app's folder
                    ' refresh the map using the restored setings.ini file
                    ' restart rocketdock
                    
                    ' .94 DAEB 26/06/2022 rDIConConfig.frm Backup and restore - fix the problem with dock entries being zeroed after a restore.
                    FileCopy bkpSettingsFile, dockSettingsFile
                    
                    Call btnSaveRestart_Click_event(rDIconConfigForm.hWnd)
                End If
            End If
    
            'ShellExecute 0, vbNullString, App.path & "\backup", vbNullString, vbNullString, 1
        End If
    Else
        ans = msgBoxA("Created an incremental backup of the Dock settings file - " & vbCr & vbCr & bkpFilename & vbCr & " Just in case of failure.", vbExclamation + vbOKOnly, "Backing up settings.")
    End If


    On Error GoTo 0
    Exit Sub

backupDockSettings_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure backupDockSettings of Form rDIconConfigForm"
            Resume Next
          End If
    End With
   
End Sub

'---------------------------------------------------------------------------------------
' Procedure : btnSaveRestart_Click_event
' Author    : beededea
' Date      : 27/03/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub btnSaveRestart_Click_event(ByRef handle As Long)

    ' variables declared

    Dim NameProcess As String: NameProcess = ""
    Dim useloop As Integer: useloop = 0
    Dim ans As Boolean: ans = False
    Dim answer As VbMsgBoxResult: answer = vbNo
    Dim itis As Boolean: itis = False
    
    'If moreConfigVisible = True Then Call picMoreConfigDown_Click ' .nn cause the new expanding section to close
    
'    If defaultDock = 0 Then ' .14 DAEB 27/02/2021 rdIConConfigForm.frm Added default dock check to ensure it works without RD installed
'        origSettingsFile = rdAppPath & "\settings.ini"
'    Else
'        origSettingsFile = sdAppPath & "\settings.ini"
''    End If

   On Error GoTo btnSaveRestart_Click_event_Error

    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
    
    If fFExists(interimSettingsFile) Then
        rDGeneralReadConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralReadConfig", interimSettingsFile)
        rDGeneralWriteConfig = GetINISetting("Software\SteamyDock\DockSettings", "GeneralWriteConfig", interimSettingsFile)
    End If
            
'    If defaultDock = 0 Then
'        NameProcess = dockAppPath & "\" & "RocketDock.exe" ' .07 DAEB 01/02/2021 rDIconConfigForm.frm Modified the parameter passed to isRunning to include the full path, otherwise it does not correlate with the found processes' folder
'    Else
        NameProcess = dockAppPath & "\" & "SteamyDock.exe" ' .07 DAEB 01/02/2021 rDIconConfigForm.frm Modified the parameter passed to isRunning to include the full path, otherwise it does not correlate with the found processes' folder
'    End If
        
    
    '.02 DAEB 26/10/2020 rDIconConfigForm.frm   Added function isRunning and changed the logic to fix a bug where the config. would not be saved if the dock was not running. STARTS.
    itis = IsRunning(NameProcess, vbNull) ' this is the check to see if the process is running
    ' kill the rocketdock /steamydock process first
    If itis = True Then
        ' .09 DAEB 07/02/2021 rDIconConfigForm.frm use the fullprocess variable without adding path again - duh!
        ans = checkAndKill(NameProcess, False, False, False) ' kill a running process
        ' if the process has died then
        If ans = True Then ' only proceed if the kill has succeeded
            PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
            FileCopy interimSettingsFile, dockSettingsFile
            'Call readInterimAndWriteConfig ' save the config.
            ' restart rocketdock /steamydock
            If fFExists(NameProcess) Then ' .09 DAEB 07/02/2021 rDIconConfigForm.frm use the fullprocess variable without adding path again - duh!
                ans = ShellExecute(handle, "Open", NameProcess, vbNullString, App.Path, 1)
            End If
        End If
    Else
        ' save the config.
        PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile
        FileCopy interimSettingsFile, dockSettingsFile
        'Call readInterimAndWriteConfig ' save the config.
        ' say not found     ' .11 DAEB 26/10/2020 rDIconConfigForm.frm No longer pops up the question if the dialog boxes are suppressed.
        If Val(sdChkToggleDialogs) = 1 Then
           answer = msgBoxA("Could not find a " & NameProcess & " process, would you like me to restart " & NameProcess & "?", vbQuestion + vbYesNo, "Restarting SteamyDock")
           If answer = vbNo Then
                msgBoxA "Current Icon Settings Saved.", vbInformation + vbYes, "Restarting SteamyDock"
                Exit Sub
            End If
        End If

        ' restart rocketdock /steamydock
        If fFExists(NameProcess) Then
            ans = ShellExecute(handle, "Open", NameProcess, vbNullString, App.Path, 1)
        End If
    End If
    '.02 DAEB 26/10/2020   Added function isRunning and changed the logic to fix a bug where the config. would not be saved if the dock was not running. ENDS.

   On Error GoTo 0
   Exit Sub

btnSaveRestart_Click_event_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnSaveRestart_Click_event of Module mdlMain"
End Sub


' .89 DAEB 13/06/2022 rDIConConfig.frm Moved backup-related private routines to modules to make them public
'---------------------------------------------------------------------------------------
' Procedure : fbackupSettings
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : Creates an incrementally named backup of the settings.ini
'---------------------------------------------------------------------------------------
' .40 DAEB 09/05/2021 rdIconConfig.frm turned into a function as it returns a value

Public Function fbackupSettings() As String

    Dim bkpSettingsFile As String
    Dim useloop As Integer: useloop = 0
    Dim srchSettingsFile As String
    Dim versionNumberAvailable As Integer
    Dim bkpfileFound As Boolean: bkpfileFound = False
    Dim dockSettingsDir As String: dockSettingsDir = vbNullString
            
    On Error GoTo fbackupSettings_Error
   
    If debugFlg = 1 Then debugLog "%" & "fbackupSettings"
    
    dockSettingsDir = SpecialFolder(SpecialFolder_AppData) & "\steamyDock" '

    ' set the name of the bkp file
    bkpSettingsFile = dockSettingsDir & "\backup\bkpSettings.ini"
            
    'check for any version of the ini file with a suffix exists
    
    For useloop = 1 To 32767
        srchSettingsFile = bkpSettingsFile & "." & useloop
      
        If fFExists(srchSettingsFile) Then
            ' found a file
            bkpfileFound = True
        Else
            ' no file found use this entry
            GoTo l_exit_bkp_loop
        End If
    Next useloop
            
l_exit_bkp_loop:

    'MsgBox "check for any version of the ini file with a suffix exists - DONE"
        
    If bkpfileFound = True Then
        bkpfileFound = False
        versionNumberAvailable = useloop
        
        'if versionNumberAvailable >= 32767 then
            'versionNumberAvailable = 1
            'If fFExists(bkpSettingsFile) Then
                'delete bkpSettingsFile
            'endif
        'endif
    Else
         versionNumberAvailable = 1
    End If
    
    bkpSettingsFile = bkpSettingsFile & "." & Trim$(Str(versionNumberAvailable))
    If Not fFExists(bkpSettingsFile) Then
        If fFExists(dockSettingsFile) Then ' .41 DAEB 09/05/2021 rdIconConfig.frm fix copying the dock settings file for backups
            FileCopy dockSettingsFile, bkpSettingsFile
        End If
    End If
    
    fbackupSettings = bkpSettingsFile

   On Error GoTo 0
   Exit Function

fbackupSettings_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fbackupSettings of Form rDIconConfigForm"
        
End Function





'.02 DAEB 26/10/2020   Created new sub readInterimAndWriteConfig to allow the save to be called more than once on a btnSaveRestart_Click
'---------------------------------------------------------------------------------------
' Procedure : readInterimAndWriteConfig
' Author    : beededea
' Date      : 26/10/2020
' Purpose   : save the current fields to the settings file or registry
'---------------------------------------------------------------------------------------
'
Public Sub readInterimAndWriteConfig()
    Dim useloop As Integer: useloop = 0
    On Error GoTo readInterimAndWriteConfig_Error
    
    PutINISetting "Software\SteamyDock\DockSettings", "lastChangedByWhom", "icoSettings", interimSettingsFile

        
    'use of the 3rd config file in the user data area first
        If steamyDockInstalled = True And defaultDock = 1 And rDGeneralWriteConfig = "True" Then ' note it will always exist even if not used
            If fFExists(interimSettingsFile) Then ' does the temporary settings.ini exist?
                ' read the registry values for each of the icons and write them to the settings.ini
                
                For useloop = 0 To rdIconMaximum
                    
                    'readSettingsIni (useloop)
                    readIconSettingsIni "Software\SteamyDock\IconSettings\Icons", useloop, interimSettingsFile

                    ' write the steamydock config file
                    
                    Call writeIconSettingsIni("Software\SteamyDock\IconSettings" & "\Icons", useloop, dockSettingsFile)
    
                    writeRegistryOnce (useloop)
                Next useloop
                
                'amend the count in the steamydock config file
                PutINISetting "Software\SteamyDock\IconSettings" & "\Icons", "count", theCount, dockSettingsFile

            End If
        End If
'        'Either of Rocketdock's two methods of saving data
'        If rDGeneralReadConfig = "False" Then
'            If fFExists(origSettingsFile) Then ' does the original settings.ini exist?
''                chkReadRegistry.Value = 0
''                chkReadSettings.Value = 1
''                chkReadConfig.Value = 0
'
'                ' write the rocketdock settings.ini
'                'writeSettingsIni (rdIconNumber) ' the settings.ini only exists when RD is set to use it
'                Call writeIconSettingsIni("Software\RocketDock" & "\Icons", rdIconNumber, interimSettingsFile)
'
'                ' copy the duplicate settings file to the original
'                FileCopy interimSettingsFile, origSettingsFile
'            Else ' Rocketdock is using the registry
''                chkReadRegistry.Value = 1
''                chkReadSettings.Value = 0
''                chkReadConfig.Value = 0
'
'                ' if the rocketdock process has died then
'                For useloop = 0 To rdIconMaximum
'
'                     'readSettingsIni (useloop)
'                    readIconSettingsIni "Software\RocketDock\Icons", useloop, interimSettingsFile
'
'                     ' write the rocketdock registry
'                    writeRegistryOnce (useloop)
'                 Next useloop
'                 '0-IsSeparator
'                 'now write the count to the registry
'                 Call savestring(HKEY_CURRENT_USER, "Software\RocketDock\Icons", "count", Str$(theCount))
'
'                 'now save the current icon folder to the registry
'                 Call savestring(HKEY_CURRENT_USER, "Software\RocketDock", "rDCustomIconFolder", rDCustomIconFolder)
'
'                 Sleep (1000) ' this is required as the o/ses final commit of the data to the registry can be delayed
'                              ' and without the pause the restart does not pick up the committed data.
'            End If
'        End If

   On Error GoTo 0
   Exit Sub

readInterimAndWriteConfig_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readInterimAndWriteConfig of Form rDIconConfigForm"

End Sub


' 17/11/2020 rDIconConfigForm.frm .05 DAEB Added the missing code to read/write the current theme to the tool's own settings file
'---------------------------------------------------------------------------------------
' Procedure : setThemeSkin
' Author    : beededea
' Date      : 13/06/2020
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub setThemeSkin(ByRef thisForm As Form)
   On Error GoTo setThemeSkin_Error

    If rDSkinTheme = "dark" Then
        Call setThemeShade(thisForm, 212, 208, 199)
    Else
        Call setThemeShade(thisForm, 240, 240, 240)
    End If

   On Error GoTo 0
   Exit Sub

setThemeSkin_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeSkin of Form dockSettings"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : setThemeShade
' Author    : beededea
' Date      : 26/09/2019
' Purpose   : if running on Win7 with the classic theme setting the theme to dark should do nothing
'             if running on any other theme then setting the theme to dark should replace the visual elements
'---------------------------------------------------------------------------------------
'
Public Sub setThemeShade(ByRef thisForm As Form, ByVal redC As Integer, ByVal greenC As Integer, ByVal blueC As Integer)

    Dim a As String: a = vbNullString
    Dim firstRun As Boolean: firstRun = False
    Dim Ctrl As Control
    Dim useloop As Integer: useloop = 0
    
    firstRun = False
    
    On Error GoTo setThemeShade_Error
    If debugFlg = 1 Then debugLog "setThemeShade"
    
    ' RGB(redC, greenC, blueC) is the background colour used by the classic theme
    
    thisForm.BackColor = RGB(redC, greenC, blueC)
    ' a method of looping through all the controls that require reversion of any background colouring
    For Each Ctrl In thisForm.Controls
        a = Ctrl.Name
        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
          Ctrl.BackColor = RGB(redC, greenC, blueC)
        End If
    Next
    
    If thisForm.Name = "rDIconConfigForm" Then

        ' exclude the label frame from any themeing
        For useloop = 0 To 11
            thisForm.fraThumbLabel(useloop).BackColor = vbWhite
        Next useloop
        
        ' the first of the thumbnail labels goes white when themed, a quick fix
        thisForm.lblThumbName(0).ForeColor = vbBlack
        
        'these buttons must be styled as they are graphical buttons with images that conform to a classic theme
        
        If redC = 212 Then
            gblClassicTheme = True
            thisForm.mnuLight.Checked = False
            thisForm.mnuDark.Checked = True
            If fFExists(App.Path & "\resources\arrowDown.jpg") Then thisForm.btnArrowDown.Picture = LoadPicture(App.Path & "\resources\arrowDown.jpg") ' imageList candidates
            If fFExists(App.Path & "\resources\leftArrow.jpg") Then thisForm.btnMapPrev.Picture = LoadPicture(App.Path & "\resources\leftArrow.jpg")
            If fFExists(App.Path & "\resources\rightArrow.jpg") Then thisForm.btnMapNext.Picture = LoadPicture(App.Path & "\resources\rightArrow.jpg")
            If fFExists(App.Path & "\resources\arrowUp.jpg") Then thisForm.btnArrowUp.Picture = LoadPicture(App.Path & "\resources\arrowUp.jpg")
            ' .52 DAEB 24/04/2022 rDIConConfig.frm Added up button to the two down buttons, theme them and add another at the bottom left
            If fFExists(App.Path & "\resources\arrowDown.jpg") Then thisForm.btnSettingsDown.Picture = LoadPicture(App.Path & "\resources\arrowDown.jpg")
            If fFExists(App.Path & "\resources\arrowUp.jpg") Then thisForm.btnSettingsUp.Picture = LoadPicture(App.Path & "\resources\arrowUp.jpg")
            If fFExists(App.Path & "\resources\arrowDown.jpg") Then thisForm.picMoreConfigDown.Picture = LoadPicture(App.Path & "\resources\arrowDown.jpg")
            'If fFExists(App.Path & "\resources\arrowUp.jpg") Then thisForm.picHideConfig.Picture = LoadPicture(App.Path & "\resources\arrowUp.jpg")
        Else
            gblClassicTheme = False
            thisForm.mnuLight.Checked = True
            thisForm.mnuDark.Checked = False
            If fFExists(App.Path & "\resources\arrowDown10.jpg") Then thisForm.btnArrowDown.Picture = LoadPicture(App.Path & "\resources\arrowDown10.jpg")
            If fFExists(App.Path & "\resources\leftArrow10.jpg") Then thisForm.btnMapPrev.Picture = LoadPicture(App.Path & "\resources\leftArrow10.jpg")
            If fFExists(App.Path & "\resources\rightArrow10.jpg") Then thisForm.btnMapNext.Picture = LoadPicture(App.Path & "\resources\rightArrow10.jpg")
            If fFExists(App.Path & "\resources\arrowUp10.jpg") Then thisForm.btnArrowUp.Picture = LoadPicture(App.Path & "\resources\arrowUp10.jpg")
            ' .52 DAEB 24/04/2022 rDIConConfig.frm Added up button to the two down buttons, theme them and add another at the bottom left
            If fFExists(App.Path & "\resources\arrowDown10.jpg") Then thisForm.btnSettingsDown.Picture = LoadPicture(App.Path & "\resources\arrowDown10.jpg")
            If fFExists(App.Path & "\resources\arrowUp10.jpg") Then thisForm.btnSettingsUp.Picture = LoadPicture(App.Path & "\resources\arrowUp10.jpg")
            If fFExists(App.Path & "\resources\arrowDown10.jpg") Then thisForm.picMoreConfigDown.Picture = LoadPicture(App.Path & "\resources\arrowDown10.jpg")
            'If fFExists(App.Path & "\resources\arrowUp10.jpg") Then thisForm.picHideConfig.Picture = LoadPicture(App.Path & "\resources\arrowUp10.jpg")
        End If
        
        ' these elements are normal elements that should have their styling removed on a classic theme
        
        ' we don't want all pictureboxes to be themed, only this one
        thisForm.picPreview.BackColor = RGB(redC, greenC, blueC)
        
        ' all other buttons go here, note we can colour buttons on VB6 succesfully without losing the theme,
        ' whilst VB.NET loses the bleeding theme deliberately and VB6 is superior in this respect.
        
        thisForm.picRdThumbFrame.BackColor = RGB(redC, greenC, blueC)
        thisForm.btnRemoveFolder.BackColor = RGB(redC, greenC, blueC)
        thisForm.picCover.BackColor = RGB(redC, greenC, blueC)
        thisForm.back.BackColor = RGB(redC, greenC, blueC)
        thisForm.sliPreviewSize.BackColor = RGB(redC, greenC, blueC)
        
        thisForm.picBusy.BackColor = RGB(redC, greenC, blueC)
        thisForm.btnCloseMoreConfig.BackColor = RGB(redC, greenC, blueC)
        thisForm.picMoreConfigDown.BackColor = RGB(redC, greenC, blueC)
        thisForm.btnArrowUp.BackColor = RGB(redC, greenC, blueC)
        thisForm.btnArrowDown.BackColor = RGB(redC, greenC, blueC)
    End If
    
    PutINISetting "Software\IconSettings", "SkinTheme", rDSkinTheme, toolSettingsFile ' now saved to the toolsettingsfile ' 17/11/2020 rDIconConfigForm.frm .05 DAEB Added the missing code to read/write the current theme to the tool's own settings file

    ' on NT6 plus using the MSCOMCTL slider with the lighter default theme, the slider
    ' fails to pick up the new theme colour fully
    ' the following lines triggers a partial colour change on the treeview that has no backcolor property
    ' this also causes a refresh of the preview pane - so don't remove it.
    ' I will have to create a new slider to overcome this - not yet tested the VB.NET version
    ' do not remove - essential

    'a = sliPreviewSize.Value
    'sliPreviewSize.Value = 1
    'sliPreviewSize.Value = a
    
    ' the above no longer required with Krool's replacement controls
    
   On Error GoTo 0
   Exit Sub

setThemeShade_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setThemeShade of Form rDIconConfigForm"

End Sub


