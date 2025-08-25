VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IcoWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'=========
'IcoWriter
'=========
'
'Version 1.0 (1.0.0009) 13-Sep-2015
'
'Accepts one or more hIcon/ColorDepth pairs to create an ICO image that
'can be retrieved as a Byte array or saved to disk as a .ICO file.
'
'USAGE
'-----
'
'Call NewICO(), then AddImage() one or more times, then BuildICO().
'
'After that you can use the ICO property to get a Byte Array or the
'SaveICO() method to save it to disk, or both.
'
'NOTES
'-----
'
'   o When run on Windows XP or later 32-bpp images support alpha channel
'     transparency.
'
'   o Does not handle "Vista" icon images (256x256).
'
'   o Handles changing color depth (see notes below) but does not resize
'     images from the source image's size to a target size.
'
'   o When decreasing color depth from the actual source image that the
'     handle holds GetDIBits() will distort the colors.  Here's a quote:
'
'      "If the requested format for the DIB matches its internal format,
'       the RGB values for the bitmap are copied. If the requested format
'       doesn't match the internal format, a color table is synthesized.
'       The following table describes the color table synthesized for
'       each format.
'
'       1_BPP   The color table consists of a black and a white entry.
'
'       4_BPP   The color table consists of a mix of colors identical to
'               the standard VGA palette.
'
'       8_BPP   The color table consists of a general mix of 256 colors
'               defined by GDI. (Included in these 256 colors are the 20
'               colors found in the default logical palette.)"
'
'     You could request the DIB for the original color depth and optimize
'     for a reduced color depth yourself and generally get better results.
'     However color reduction (color quantization) isn't a trivial process
'     and there are several algorithms that give varying results.  Or you
'     might generate a palette from the original colors used.  Or you might
'     even combine the two strategies.
'
'     That is what we do here, though color reduction is done by taking a
'     number of crude stabs at it.  A real color quantization algorithm
'     might give far better results.  If color reduction attempts fail to
'     create a usable palette we fall back on GetDIBits()'s palettization.
'
'     THIS MAY MAKE SOME NASTY MISTAKES but it generally gives far better
'     results than letting GetDIBits() apply its color tables.
'
'     Something similar might also be done to compensate for the "flattening"
'     of an original 32-bpp alpha channel bitmap to 24-bpp or less.  As
'     things are this just gets stripped out (effectively set to "solid
'     color" which is often black).  Something that examined alpha channel
'     values and had a threshold for setting each pixel transparent might
'     produce much better results.
'
'   o Special case: ForceWhite = True passed to AddImage() forces white as
'     well as black into the palette for ColorDepth values <= 8.
'

'===== API DECLARATIONS =======================================================

Private Const API_NULL As Long = 0

Private Const DIB_RGB_COLORS = 0 'Color table in RGBs.
Private Const BI_RGB = 0& 'Compression.

Private Type BITMAPINFOHEADER
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

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As Long 'As RGBQUAD.
End Type

'We can't use this due to the inability to statically allocate varying
'numbers of elements for the array members:
'Private Type ICONIMAGE
'    'Only the following members of icHeader are used:
'    '
'    '   biSize, biWidth, biHeight, biPlanes, biBitCount, biSizeImage.
'    '
'    '   All other members must be 0.
'    '
'    '   The biHeight member specifies the combined height of the XOR
'    '   and AND masks.
'    '
'    'The AND mask is a monochrome DIB, with a color depth of 1 bpp.
'    icHeader As BITMAPINFOHEADER 'DIB header.
'    icColors(n) As Long    'Color table.
'    icXOR(m) As Byte       'DIB bits for XOR mask.
'    icAND(l) As Byte       'DIB bits for AND mask.
'End Type

'Instead we'll cobble it up from its parts, and we'll hold these in a
'Collection.

Private Type ICONDIRENTRY
    bWidth As Byte         'Width of the image in pixels.
    bHeight As Byte        'Height of the image in pixels.
    bColorCount As Byte    'Number of colors in image (0 if >=8bpp).
    bReserved As Byte      'Reserved (must be 0).
    wPlanes As Integer     'Number of Color Planes.
    wBitCount As Integer   'Bits per pixel.
    dwBytesInRes As Long   'Count of bytes in this ICONIMAGE resource.
    dwImageOffset As Long  'Image data offset within the file.
End Type

Private Const ICONDATA_TYPE_ICON As Integer = 1
Private Const ICONDATA_TYPE_CURSOR As Integer = 2 'Not supported here.

'We can't use this due to UDT-padding and the inability to statically
'allocate varying numbers of idEntries() elements:
'Private Type ICONDIR
'    idReserved As Integer  'Reserved (must be 0).
'    idType As Integer      'Resource Type (1 for icons).
'    idCount As Integer     'Icon image count.
'    idEntries(n) As ICONDIRENTRY 'One entry for each image.
'End Type

'So instead we'll use this and a separate idEntries() array:
Private Type ICONDIR_HEADER
    idReserved As Integer  'Reserved (must be 0).
    idType As Integer      'Resource Type (1 for icons).
    idCount As Integer     'Icon image count.
End Type

Private Type ICONINFO
    fIcon As Long     'API TRUE for icon, API FALSE for cursor.
    xHotspot As Long  'The hotspot X-coordinate for cursor.
    yHotspot As Long  'The hotspot Y-coordinate for cursor.
    hbmMask As Long   'HBITMAP handle to monochrome AND mask bitmap.
    hbmColor As Long  'HBITMAP handle to device dependent XOR mask bitmap.
End Type

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    ByVal lpBits As Long, _
    ByVal lpBI As Long, _
    ByVal wUsage As Long) As Long

Private Declare Function GetIconInfo Lib "user32" ( _
    ByVal hIcon As Long, _
    ByRef II As ICONINFO) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByVal pSource As Long, _
    ByVal Length As Long)

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

'===== PRIVATE DATA ===========================================================

'-Part of a "binary string builder" ---------------------------
Private Const ALLOCATION_CHUNK As Long = 8192

Private Buffer() As Byte
Private Cursor As Long
'--------------------------------------------------------------

Private hDC As Long 'Screen hDC.
Private ICONINFO As ICONINFO
Private ICONDIR_HEADER As ICONDIR_HEADER
Private idEntries() As ICONDIRENTRY 'Dimensioned from 1 to count.
Private Images As Collection        'Byte array pseudo-ICONIMAGE structures.

Private mICO() As Byte 'Finished ICO file image.

'===== PUBLIC MEMBERS =========================================================

Public Property Get ICO() As Byte()
    ICO = mICO
End Property

Public Property Get ImageCount() As Long
    If Not Images Is Nothing Then ImageCount = UBound(idEntries)
End Property

Public Property Get ColorDepth(ByVal ImageIndex As Integer) As Integer
    If 1 > ImageIndex Or ImageIndex > ImageCount Then Err.Raise 5 'Invalid procedure call or argument.
    ColorDepth = idEntries(ImageIndex).wBitCount
End Property

Public Property Get Width(ByVal ImageIndex As Integer) As Long
    If 1 > ImageIndex Or ImageIndex > ImageCount Then Err.Raise 5 'Invalid procedure call or argument.
    Width = idEntries(ImageIndex).bWidth
End Property

Public Property Get Height(ByVal ImageIndex As Integer) As Long
    If 1 > ImageIndex Or ImageIndex > ImageCount Then Err.Raise 5 'Invalid procedure call or argument.
    Height = idEntries(ImageIndex).bHeight
End Property

Public Sub SaveICO(ByVal File As String, Optional ByVal Overwrite As Boolean = True)
    Dim BackPos As Long
    Dim DotPos As Long
    Dim F As Integer
    
    'Force file extension to ".ico" here:
    If LCase$(Right$(File, 4)) <> ".ico" Then
        BackPos = InStrRev(File, "\")
        DotPos = InStrRev(File, ".")
        If DotPos > BackPos Then
            File = Left$(File, DotPos - 1)
        End If
        File = File & ".ico"
    End If
    
    If Overwrite Then
        On Error Resume Next
        Kill File
        On Error GoTo 0
    End If
    F = FreeFile(0)
    Open File For Binary Access Write As #F
    Put #F, , mICO
    Close #F
End Sub

Public Sub ReleaseMemory()
    'Only called when you want to free up space when holding an instance
    'of this class for long periods.  Optional.
    
    Buffer = "" 'Empty but valid array.
    mICO = ""
    Erase idEntries
    Set Images = Nothing
End Sub

Public Sub NewICO()
    With ICONDIR_HEADER
        .idType = ICONDATA_TYPE_ICON
    End With
    Set Images = New Collection
End Sub

Public Sub AddImage( _
    ByVal hIcon As Long, _
    ByVal ColorDepth As Long, _
    Optional ByVal ForceWhite As Boolean = False)
    
    Dim LastDllError As Long
    Dim ColorCount As Long
    Dim MaskBITMAPINFO As BITMAPINFO
    Dim MaskBits() As Long
    Dim ColorBITMAPINFO As BITMAPINFO
    Dim ColorBits() As Long
    Dim ImageIndex As Long
    Dim ColorBMISize As Long
    Dim ColorSizeImage As Long
    Dim ImageData() As Byte
    
    Select Case ColorDepth
        Case 1, 4, 8, 24, 32
            'Ok.
        Case Else
            Err.Raise &H80048FFE, _
                      TypeName(Me), _
                      "Unsupported ColorDepth"
    End Select
    
    If Images Is Nothing Then
        Err.Raise &H80048FFF, _
                  TypeName(Me), _
                  "Unsupported ColorDepth"
    End If

    If GetIconInfo(hIcon, ICONINFO) = 0 Then
        Err.Raise &H80048000 Or (Err.LastDllError And &H7FFF&), _
                  TypeName(Me), _
                  "GetIconInfo() error " & CStr(Err.LastDllError)
    End If
    
    If ColorDepth <= 8 Then
        ColorCount = 2 ^ ColorDepth
    End If
    
    hDC = GetDC(API_NULL)
    With ICONINFO
        GetImageInfoAndBits False, .hbmMask, 1, 2, False, MaskBITMAPINFO, MaskBits
        If .hbmColor = API_NULL Then
            'We go through some gyrations here because a 1-bpp black and white icon does
            'not give us a separate hBitmap for the color and mask images.  Instead we
            'only get a mask hBitmap and it has the .biHeight doubled and both sets of
            'bits come back from that in one piece!
            With MaskBITMAPINFO.bmiHeader
                .biHeight = .biHeight \ 2
            End With
            ColorBITMAPINFO = MaskBITMAPINFO
            With ColorBITMAPINFO.bmiHeader
                .biSizeImage = 0
                .biClrUsed = 2
                .biClrImportant = 2
                .biBitCount = 1
                .biCompression = BI_RGB
            End With
        Else
            GetImageInfoAndBits True, _
                                .hbmColor, _
                                ColorDepth, _
                                ColorCount, _
                                ForceWhite, _
                                ColorBITMAPINFO, _
                                ColorBits
        End If
    End With
    GdiCleanup
    
    ImageIndex = Images.Count + 1
    ReDim Preserve idEntries(1 To ImageIndex)
    ColorBMISize = LenB(ColorBITMAPINFO.bmiHeader) + 4& * ColorCount
    
    'Build idEntries(ImageIndex):
    With idEntries(ImageIndex)
        .bWidth = ColorBITMAPINFO.bmiHeader.biWidth
        .bHeight = ColorBITMAPINFO.bmiHeader.biHeight
        .bColorCount = ColorCount And &HFF&
        .wPlanes = ColorBITMAPINFO.bmiHeader.biPlanes
        .wBitCount = ColorBITMAPINFO.bmiHeader.biBitCount
        .dwBytesInRes = ColorBMISize _
                      + ColorBITMAPINFO.bmiHeader.biSizeImage _
                      + MaskBITMAPINFO.bmiHeader.biSizeImage
    End With
    
    'Build an ICONIMAGE structure in bits and pieces:
    With ColorBITMAPINFO.bmiHeader
        ColorSizeImage = .biSizeImage 'Retain value before we reuse the structure here.
        .biSizeImage = 0
        .biHeight = .biHeight * 2 'Both bitmap images.
        .biCompression = BI_RGB
    End With
    OpenBuf
    AppendBuf VarPtr(ColorBITMAPINFO), ColorBMISize
    If ColorSizeImage Then
        AppendBuf VarPtr(ColorBits(0)), ColorSizeImage
    'Else
        'MaskBits() contains the color bits.
    End If
    AppendBuf VarPtr(MaskBits(0)), MaskBITMAPINFO.bmiHeader.biSizeImage
    ImageData = ReadBuf()
    
    Images.Add ImageData
End Sub

Public Sub BuildICO()
    Dim Length As Long
    Dim ImageIndex As Long
    Dim ImageData() As Byte
    
    OpenBuf
    
    '---Construct and append our ICONDIR in bits and pieces: ------------------
    
    'Finalize and append the ICONDIR_HEADER structure:
    ICONDIR_HEADER.idCount = ImageCount
    AppendBuf VarPtr(ICONDIR_HEADER), LenB(ICONDIR_HEADER)
    
    'Finalize and append the idEntries() structures:
    Length = LenB(idEntries(1))
    idEntries(1).dwImageOffset = Cursor + Length * Images.Count
    AppendBuf VarPtr(idEntries(1)), Length
    For ImageIndex = 2 To Images.Count
        With idEntries(ImageIndex - 1) 'Previous image.
            idEntries(ImageIndex).dwImageOffset = .dwImageOffset + .dwBytesInRes
        End With
        AppendBuf VarPtr(idEntries(ImageIndex)), Length
    Next
    '--------------------------------------------------------------------------
    
    'Append our saved ICONIMAGE structures:
    For ImageIndex = 1 To Images.Count
        ImageData = Images.Item(ImageIndex)
        AppendBuf VarPtr(ImageData(0)), idEntries(ImageIndex).dwBytesInRes
    Next
    
    mICO = ReadBuf()
End Sub

'===== PRIVATE METHODS ========================================================

Private Sub GetImageInfoAndBits( _
    ByVal FetchingColorBitmap As Boolean, _
    ByVal hBitmap As Long, _
    ByVal ColorDepth As Long, _
    ByVal ColorCount As Long, _
    ByVal ForceWhite As Boolean, _
    ByRef BITMAPINFO As BITMAPINFO, _
    ByRef Bits() As Long)
    
    Dim LastDllError As Long
    Dim NoSynthesizedPalette As Boolean

    With BITMAPINFO.bmiHeader
        .biSize = LenB(BITMAPINFO.bmiHeader)
        .biBitCount = 0 'Don't fetch color table.
        If GetDIBits(hDC, _
                     hBitmap, _
                     0, _
                     0, _
                     API_NULL, _
                     VarPtr(BITMAPINFO), _
                     DIB_RGB_COLORS) = 0 Then
            LastDllError = Err.LastDllError
            GdiCleanup
            Err.Raise &H80048000 Or (LastDllError And &H7FFF&), _
                      TypeName(Me), _
                      "GetDIBits() error " & CStr(LastDllError)
        End If
        
        NoSynthesizedPalette = True
        If FetchingColorBitmap Then
            If ColorDepth <= 8 Then
                NoSynthesizedPalette = _
                    SynthesizePalette(hBitmap, ColorDepth, ColorCount, ForceWhite, BITMAPINFO, Bits)
            End If
            If NoSynthesizedPalette Then
                .biBitCount = ColorDepth
                .biCompression = BI_RGB
            End If
        End If
        
        If NoSynthesizedPalette Then
            ReDim Bits((((.biWidth * .biBitCount + 7) \ 8& + 3) \ 4&) * .biHeight - 1)
            If GetDIBits(hDC, _
                         hBitmap, _
                         0, _
                         .biHeight, _
                         VarPtr(Bits(0)), _
                         VarPtr(BITMAPINFO), _
                         DIB_RGB_COLORS) = 0 Then
                LastDllError = Err.LastDllError
                GdiCleanup
                Err.Raise &H80048000 Or (LastDllError And &H7FFF&), _
                          TypeName(Me), _
                          "GetDIBits() error " & CStr(LastDllError)
            End If
        End If
    End With
End Sub

Private Function SynthesizePalette( _
    ByVal hBitmap As Long, _
    ByVal ColorDepth As Long, _
    ByVal ColorCount As Long, _
    ByVal ForceWhite As Boolean, _
    ByRef ColorBITMAPINFO As BITMAPINFO, _
    ByRef ColorBits() As Long) As Boolean
    'Returns True on too many colors (failure).
    
    Dim LastDllError As Long
    Dim BITMAPINFO As BITMAPINFO
    Dim Bits() As Long
    Dim PalettedBitsSize As Long
    Dim PalettedBits() As Byte 'Not Long here but scanlines padded to DWORD boundaries.
    Dim PScanLineBytes As Long
    
    BITMAPINFO = ColorBITMAPINFO 'Copy previously probed specs.
    With BITMAPINFO.bmiHeader
        .biBitCount = 32 'We need True Color values but want easy Long access to ColorBits().
        .biCompression = BI_RGB
        ReDim Bits(.biWidth * .biHeight - 1)
        If GetDIBits(hDC, _
                     hBitmap, _
                     0, _
                     .biHeight, _
                     VarPtr(Bits(0)), _
                     VarPtr(BITMAPINFO), _
                     DIB_RGB_COLORS) = 0 Then
            LastDllError = Err.LastDllError
            GdiCleanup
            Err.Raise &H80048000 Or (LastDllError And &H7FFF&), _
                      TypeName(Me), _
                      "GetDIBits() error " & CStr(LastDllError)
        End If
        .biClrUsed = ColorCount
        .biBitCount = ColorDepth
        .biCompression = BI_RGB
        PScanLineBytes = (((.biWidth * .biBitCount + 7) \ 8& + 3) \ 4&) * 4&
        PalettedBitsSize = PScanLineBytes * .biHeight
        ReDim PalettedBits(PalettedBitsSize - 1)
    End With
    BITMAPINFO.bmiColors(0) = vbBlack 'Always must have black.
    If ForceWhite Then
        BITMAPINFO.bmiColors(1) = vbWhite '1-bpp black & white.
    End If
    
    'Several approaches are tried in turn here, each one further restricting
    'the color space.  We want to consider as many bits of the image colors
    'as we can and still fit the palette size.
    
    'First try all 8 bits of R, G, and B:
    If BuildPalette(&HFFFFFF, BITMAPINFO, Bits, ColorDepth, ColorCount, ForceWhite, PScanLineBytes, PalettedBits) Then
        'Try dropping lower two bits of R, G, and B:
        If BuildPalette(&HFCFCFC, BITMAPINFO, Bits, ColorDepth, ColorCount, ForceWhite, PScanLineBytes, PalettedBits) Then
            'Try dropping lower four bits of R, G, and B:
            If BuildPalette(&HF0F0F0, BITMAPINFO, Bits, ColorDepth, ColorCount, ForceWhite, PScanLineBytes, PalettedBits) Then
                'Try dropping lower five bits of R, G, and B:
                If BuildPalette(&HE0E0E0, BITMAPINFO, Bits, ColorDepth, ColorCount, ForceWhite, PScanLineBytes, PalettedBits) Then
                    If ColorDepth < 8 Then
                        'Try dropping lower six bits of R, G, and B:
                        If BuildPalette(&HC0C0C0, BITMAPINFO, Bits, ColorDepth, ColorCount, ForceWhite, PScanLineBytes, PalettedBits) Then
                            'Try dropping another bit-pattern from R, G, and B:
                            If BuildPalette(&H808888, BITMAPINFO, Bits, ColorDepth, ColorCount, ForceWhite, PScanLineBytes, PalettedBits) Then
                                'Try dropping lower seven bits of R, G, and B:
                                If BuildPalette(&H808080, BITMAPINFO, Bits, ColorDepth, ColorCount, ForceWhite, PScanLineBytes, PalettedBits) Then
                                    SynthesizePalette = True 'We have failed.
                                End If
                            End If
                        End If
                    Else
                        SynthesizePalette = True 'We have failed.
                    End If
                End If
            End If
        End If
    End If
    If Not SynthesizePalette Then
        'Success:
        ReDim ColorBits(PalettedBitsSize \ 4& - 1)
        MoveMemory ColorBits(0), VarPtr(PalettedBits(0)), PalettedBitsSize
        BITMAPINFO.bmiHeader.biSizeImage = PalettedBitsSize
        ColorBITMAPINFO = BITMAPINFO
    End If
End Function

Private Function BuildPalette( _
    ByVal QuantMask As Long, _
    ByRef BITMAPINFO As BITMAPINFO, _
    ByRef Bits() As Long, _
    ByVal ColorDepth As Long, _
    ByVal ColorCount As Long, _
    ByVal ForceWhite As Boolean, _
    ByVal PScanLineBytes As Long, _
    ByRef PalettedBits() As Byte) As Boolean
    'QuantMask is used to reduce color depth.  It must drop the alpha channel,
    'possibly along with other bits of the RGB colors.
    '
    'Returns True on too many colors (failure).
    
    Dim ScanLine As Long
    Dim ScanStartX As Long 'Index to start of current scan line of Bits().
    Dim PScanStartX As Long 'Index to start of current scan line of PalettedBits().
    Dim POffset As Long 'Pixel index within scan line.
    Dim Pixel As Long 'Color from Bits() under examination.
    Dim CX As Long 'Color index.
    Dim PX As Long 'Byte index of current pixel within PalettedBits().
    Dim PixelBit As Long

    With BITMAPINFO.bmiHeader
        'Note: We use .biClrImportant as a base-0 color index value below
        'and then corect it upon successful palette construction.
        If ForceWhite Then
            .biClrImportant = 1 'Reset to two colors (black and white).
        Else
            .biClrImportant = 0 'Reset to one color (black).
        End If
        For ScanLine = 0 To .biHeight - 1
            ScanStartX = ScanLine * .biWidth
            PScanStartX = ScanLine * PScanLineBytes
            For POffset = 0 To .biWidth - 1
                Pixel = Bits(ScanStartX + POffset)
                'Seach color table for a quant-match:
                For CX = 0 To .biClrImportant
                    If (Pixel And QuantMask) = (BITMAPINFO.bmiColors(CX) And QuantMask) Then Exit For
                Next
                If CX > .biClrImportant Then
                    'Add new color to table if room:
                    If CX + 1 > ColorCount Then
                        'Too many colors:
                        If ColorDepth > 1 Then
                            BuildPalette = True 'Failure.
                            Exit Function 'Too many colors!
                        Else
                            'We can't let this fail or we can lose and get black and white
                            'every time.
                            '
                            'Dice roll, but perhaps better than just picking color one
                            'for every 1-bpp image.
                            If Abs((Pixel And &HFFFFFF) - vbBlack) < _
                               Abs((Pixel And &HFFFFFF) - BITMAPINFO.bmiColors(1)) Then
                                CX = 0 'Closer to black.
                            Else
                                CX = 1
                            End If
                        End If
                    Else
                        'Add this color to palette:
                        .biClrImportant = CX
                        'Use actual color, not QuantMask'ed color:
                        BITMAPINFO.bmiColors(CX) = Pixel And &HFFFFFF 'Strip alpha channel.
                    End If
                End If
                'Set paletted pixel to color index:
                Select Case ColorDepth
                    Case 8
                        PalettedBits(PScanStartX + POffset) = CX
                    Case 4
                        PX = PScanStartX + POffset \ 2
                        If POffset Mod 2 Then
                            PalettedBits(PX) = PalettedBits(PX) And &HF0 Or CX
                        Else
                            PalettedBits(PX) = PalettedBits(PX) And &HF Or (CX * &H10&)
                        End If
                    Case Else '0 or 1, i.e. 1-bpp.
                        PX = PScanStartX + POffset \ 8
                        PixelBit = POffset Mod 8
                        'This looks ugly, but does fewer comparisons than a Select Case here:
                        If PixelBit > 3 Then
                            If PixelBit > 5 Then
                                If PixelBit = 7 Then
                                    PalettedBits(PX) = PalettedBits(PX) And &HFE Or CX
                                Else
                                    PalettedBits(PX) = PalettedBits(PX) And &HFD Or (CX * &H2&)
                                End If
                            Else
                                If PixelBit = 5 Then
                                    PalettedBits(PX) = PalettedBits(PX) And &HFB Or (CX * &H4&)
                                Else
                                    PalettedBits(PX) = PalettedBits(PX) And &HF7 Or (CX * &H8&)
                                End If
                            End If
                        Else
                            If PixelBit > 1 Then
                                If PixelBit = 3 Then
                                    PalettedBits(PX) = PalettedBits(PX) And &HEF Or (CX * &H10&)
                                Else
                                    PalettedBits(PX) = PalettedBits(PX) And &HDF Or (CX * &H20&)
                                End If
                            Else
                                If PixelBit = 1 Then
                                    PalettedBits(PX) = PalettedBits(PX) And &HBF Or (CX * &H40&)
                                Else
                                    PalettedBits(PX) = PalettedBits(PX) And &H7F Or (CX * &H80&)
                                End If
                            End If
                        End If
                End Select
            Next
        Next
        .biClrImportant = .biClrImportant + 1 'Correct to base-1 value.
    End With
End Function

'-Part of a "binary string builder" ---------------------------
Private Sub OpenBuf()
    Cursor = 0
End Sub

Private Function ReadBuf() As Byte()
    If Cursor = 0 Then
        ReadBuf = "" 'Empty array.
    Else
        If UBound(Buffer) > Cursor - 1 Then ReDim Preserve Buffer(Cursor - 1)
    End If
    ReadBuf = Buffer
End Function

Private Sub AppendBuf(ByVal Pointer As Long, ByVal Length As Long)
    Dim NeedUB As Long
    Dim BufferUB As Long
    
    NeedUB = Cursor + Length - 1
    BufferUB = UBound(Buffer)
    If NeedUB > BufferUB Then
        ReDim Preserve Buffer(NeedUB + ALLOCATION_CHUNK)
    End If
    
    MoveMemory Buffer(Cursor), Pointer, Length
    Cursor = Cursor + Length
End Sub
'--------------------------------------------------------------

Private Sub GdiCleanup()
    If hDC Then
        DeleteObject ICONINFO.hbmColor
        DeleteObject ICONINFO.hbmMask
        ReleaseDC API_NULL, hDC
        hDC = 0
    End If
End Sub

'===== EVENT HANDLERS =========================================================

Private Sub Class_Initialize()
    Buffer = "" 'Empty but valid array.
    mICO = ""
End Sub

Private Sub Class_Terminate()
    GdiCleanup
End Sub
