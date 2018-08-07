<script language="VBScript">

'ImgX_FileSaveTypes Enumeration
Const ixfsJPG = 1
Const ixfsPNG = 2
Const ixfsTIF = 3
Const ixfsPCX = 4
Const ixfsTGA = 5
Const ixfsBMP = 6
Const ixfsWMF = 7
Const ixfsEMF = 8
Const ixfsPSD = 9
Const ixfsWBMP = 10
Const ixfsGIF = 11
Const ixfsTLA = 12

'ImgXTwain_Capabilities Enumeration
Const ixtcXFerCount = &H1
Const ixtcPixelType = &H101
Const ixtcUnits = &H102
Const ixtcFeederEnabled = &H1002
Const ixtcFeederLoaded = &H1003
Const ixtcFeederControl = &H1006
Const ixtcAutoFeed = &H1007
Const ixtcIndicators = &H100B
Const ixtcFeederOrder = &H102E
Const ixtcAutoBright = &H1100
Const ixtcBrightness = &H1101
Const ixtcContrast = &H1103
Const ixtcGamma = &H1108
Const ixtcPhysicalWidth = &H1111
Const ixtcPhysicalHeight = &H1112
Const ixtcResolution = &H1118
Const ixtcScaling = &H1124
Const ixtcBitDepth = &H112B
Const ixtcZoomFactor = &H113E

'ImgXTwain_PixelTypes Enumeration
Const TWPT_BW = 0
Const TWPT_GRAY = 1
Const TWPT_RGB = 2
Const TWPT_PALETTE = 3

'ImgXTwain_Units Enumeration
Const TWUN_INCHES = 0
Const TWUN_CENTIMETERS = 1
Const TWUN_PICAS = 2
Const TWUN_POINTS = 3
Const TWUN_TWIPS = 4
Const TWUN_PIXELS = 5

'ImageTypes Enumeration
Const ixitPalette_1 = 0
Const ixitPalette_8 = 1
Const ixitPalWithPalAlpha_8 = 2
Const ixitGrayscale_8 = 3
Const ixitGrayscaleWithAlpha_16 = 4
Const ixitRGB_24 = 5
Const ixitRGBWithAlphaChannel_32 = 6

'TIFCompressionTypes Enumeration
Const ixtcNoCompression = 0
Const ixtcGroup3FAXEncoding = 1
Const ixtcGroup4FAXEncoding = 2
Const ixtcJPEGCompression = 3
Const ixtcMacintoshPackbits = 4
Const ixtcDeflate = 5
Const ixtcLZW = 6
Const ixtcModifiedHuffman = 7

'MouseTools Enumeration
Const ixmtNone = 0
Const ixmtSelect = 1
Const ixmtPan = 2
Const ixmtZoomArea = 3
Const ixmtZoom = 4
Const ixmtZoomIn = 5
Const ixmtZoomOut = 6
Const ixmtLine = 7
Const ixmtLines = 8
Const ixmtRectangle = 9
Const ixmtEllipse = 10
Const ixmtFreehand = 11
Const ixmtMagnifier = 12

</script>