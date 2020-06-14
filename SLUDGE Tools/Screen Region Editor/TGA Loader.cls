VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TGALoaderClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This module contains the procedures for loading and displaying True Vision Targa (.TGA) images.
Option Explicit

'The Microsoft Windows API constants, functions, and structures used by this class:
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

Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Private Type RGBTRIPLE
   rgbtBlue As Byte
   rgbtGreen As Byte
   rgbtRed As Byte
End Type

Private Const BI_RGB As Long = &H0&
Private Const CBM_INIT As Long = &H4&
Private Const DIB_RGB_COLORS As Long = &H0&
Private Const SRCCOPY As Long = &HCC0020

Private Declare Function CreateCompatibleDC Lib "Gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBitmap Lib "Gdi32.dll" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function DeleteDC Lib "Gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "Gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "User32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SelectObject Lib "Gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "Gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Sub RtlMoveMemory Lib "Kernel32.dll" (Destination As Byte, Source As Byte, ByVal Length As Long)

'The constants, functions, and structures used by this class:

'This structure defines the bitmap information header and pallete.
Private Type BitmapStr
   Header As BITMAPINFOHEADER    'The bitmap header.
   Palette(255) As RGBQUAD       'The bitmap palette.
End Type

'This structure defines the TGA image header.
Private Type TGAHeaderStr
   InformationBlockSize As Byte  'Defines the length of the image information block.
   ColorType As Byte             'Indicates DAC table or BGR format.
   ImageType As Byte             'Defines the image type.
   Origin As Integer             'Defiens the first entry in the DAC table.
   ColorCount As Integer         'Defines the number of colors in the DAC table.
   EntryBits As Byte             'Defines the number of bits per color in the DAC table.
   LowerLeftCornerX As Integer   'Defines the x coordinate of the lower left corner.
   LowerLeftCornerY As Integer   'Defines the y coordinate of the lower left corner.
   ImageWidth As Integer         'Defines the image width.
   ImageHeight As Integer        'Defines the image height.
   BitsPerPixel As Byte          'Defines the number of bits per pixel.
   Descriptor As Byte            'Defines the image descriptor.
End Type

Private Const COLOR_MAPPED As Byte = 1                   'Indicates that the image is color mapped.
Private Const NOT_COLOR_MAPPED As Byte = 0               'Indicates that the image is not color mapped.
Private Const ORIENTATION_BIT As Long = &H3&             'This bit specifies a TGA image's orientation.
Private Const TGA_COLOR_MAP As Byte = 1                  'Indicates uncompressed color-mapped image.
Private Const TGA_COMPRESSED_COLOR_MAP_A As Byte = 32    'Indicates compressed color-mapped data.
Private Const TGA_COMPRESSED_COLOR_MAP_B As Byte = 33    'Indicates compressed color-mapped data.
Private Const TGA_COMPRESSED_MONOCHROME As Byte = 11     'Indicates compressed black and white image.
Private Const TGA_MONOCHROME As Byte = 3                 'Indicates uncompressed black and white image.
Private Const TGA_RGB As Byte = 2                        'Indicates uncompressed RGB image.
Private Const TGA_RLE_COLOR_MAP As Byte = 9              'Indicates RLE encoded color-mapped image.
Private Const TGA_RLE_RGB As Byte = 10                   'Indicates RLE encoded RGB image.
'This procedure creates a bitmap from the specified pixel data.
Private Function CreateBitmap(Bitmap As BitmapStr, BitmapData() As Byte) As Long
Dim BitmapDC As Long
Dim BitmapHandle As Long

   With Bitmap.Header
      .biSize = Len(Bitmap.Header)
      .biPlanes = 1
      .biCompression = BI_RGB
      .biSizeImage = 0
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
         
   BitmapDC = GetDC(0)
   BitmapHandle = CreateDIBitmap(BitmapDC, Bitmap.Header, CBM_INIT, BitmapData(0), Bitmap.Header, DIB_RGB_COLORS)
   DeleteDC BitmapDC
   
   CreateBitmap = BitmapHandle
End Function


'This procedure decompresses the specified TGA image data if necessary.
Private Function DecompressTGA(CompressedData() As Byte, Bitmap As BitmapStr, TGAHeader As TGAHeaderStr) As Byte()
Dim BitmapSize As Long
Dim BytesPerPixel As Long
Dim BytesProcessed As Long
Dim DecompressedData() As Byte
Dim Index As Long
Dim Pixel As Long
Dim PixelBytesCount As Long
Dim PixelCount As Long

   If TGAHeader.ImageType = TGA_RLE_COLOR_MAP Or TGAHeader.ImageType = TGA_RLE_RGB Then
      BytesPerPixel = Bitmap.Header.biBitCount / 8
      BitmapSize = CLng(Bitmap.Header.biHeight * Bitmap.Header.biWidth * BytesPerPixel)
      Index = 0
      
      ReDim DecompressedData(0 To BitmapSize) As Byte
      Do Until (BytesProcessed = BitmapSize) Or (Index > UBound(CompressedData()))
         If CompressedData(Index) < 128 Then
            PixelCount = CompressedData(Index) + 1
            PixelBytesCount = PixelCount * BytesPerPixel
            RtlMoveMemory DecompressedData(BytesProcessed), CompressedData(Index + 1), PixelBytesCount
            BytesProcessed = BytesProcessed + PixelBytesCount
            
            Index = Index + (PixelCount * BytesPerPixel) + 1
            PixelCount = PixelCount + PixelBytesCount
         ElseIf CompressedData(Index) > 127 Then
            PixelCount = CompressedData(Index) - 127
            
            For Pixel = 0 To PixelCount - 1
               RtlMoveMemory DecompressedData(BytesProcessed), CompressedData(Index + 1), BytesPerPixel
               BytesProcessed = BytesProcessed + BytesPerPixel
            Next Pixel
            
            Index = Index + BytesPerPixel + 1
            PixelCount = (PixelCount * BytesPerPixel) + PixelBytesCount + 1
         End If
      Loop
   Else
      DecompressedData() = CompressedData()
   End If
   
   DecompressTGA = DecompressedData()
End Function

'This procedure draws the specified TGA image on the specified canvas.
Public Function DrawTGA(Canvas As Object, FileName As String, Optional ResizeCanvas As Boolean = True) As Boolean
Dim Bitmap As BitmapStr
Dim BitmapHandle As Long
Dim CanvasDC As Long
Dim PreviousScaleMode As Long
Dim PreviousParentScaleMode As Long

   Bitmap = LoadTGA(FileName, BitmapHandle)
   
   If Not BitmapHandle = 0 Then
      Canvas.AutoRedraw = True
      Canvas.Cls
      PreviousScaleMode = Canvas.ScaleMode
      Canvas.ScaleMode = vbPixels
       
      If TypeOf Canvas Is Form Then
         If ResizeCanvas And Not Canvas.WindowState = vbMaximized Then
            Canvas.Width = Bitmap.Header.biWidth * Screen.TwipsPerPixelX
            Canvas.Height = Bitmap.Header.biHeight * Screen.TwipsPerPixelY
         End If
      Else
         PreviousParentScaleMode = Canvas.Parent.ScaleMode
         Canvas.Parent.ScaleMode = vbPixels
   
         If ResizeCanvas Then
            Canvas.Width = Bitmap.Header.biWidth
            Canvas.Height = Bitmap.Header.biHeight
         End If
      End If
         
      CanvasDC = CreateCompatibleDC(Canvas.hdc)
      SelectObject CanvasDC, BitmapHandle
      StretchBlt Canvas.hdc, 0, 0, Canvas.ScaleWidth, Canvas.ScaleHeight, CanvasDC, 0, 0, Bitmap.Header.biWidth, Bitmap.Header.biHeight, SRCCOPY
      DeleteDC CanvasDC
      DeleteObject BitmapHandle
       
      Canvas.ScaleMode = PreviousScaleMode
      Canvas.Picture = Canvas.Image
      If Not TypeOf Canvas Is Form Then Canvas.Parent.ScaleMode = PreviousParentScaleMode
   End If
   
   DrawTGA = Not (BitmapHandle = 0)
End Function

'This procedure indicates whether the current image is a TGA image.
Private Function IsTGA(TGAHeader As TGAHeaderStr) As Boolean
Dim BitsPerPixelIsTGA As Boolean
Dim ColorTypeIsTGA As Boolean
Dim ImageTypeIsTGA As Boolean

   IsTGA = False
   
   Select Case TGAHeader.BitsPerPixel
      Case 8, 16, 24, 32
         BitsPerPixelIsTGA = True
      Case Else
         BitsPerPixelIsTGA = False
   End Select
   
   Select Case TGAHeader.ColorType
      Case COLOR_MAPPED, NOT_COLOR_MAPPED
         ColorTypeIsTGA = True
      Case Else
         ColorTypeIsTGA = False
   End Select
   
   Select Case TGAHeader.ImageType
      Case TGA_COLOR_MAP, TGA_COMPRESSED_COLOR_MAP_A, TGA_COMPRESSED_COLOR_MAP_B, TGA_COMPRESSED_MONOCHROME, TGA_MONOCHROME, TGA_RGB, TGA_RLE_RGB, TGA_RLE_COLOR_MAP
         ImageTypeIsTGA = True
      Case Else
         ImageTypeIsTGA = False
   End Select
    
   IsTGA = (ColorTypeIsTGA And ImageTypeIsTGA And BitsPerPixelIsTGA)
End Function


'This procedure loads the specified TGA image.
Private Function LoadTGA(FileName As String, ByRef BitmapHandle As Long) As BitmapStr
Dim Bitmap As BitmapStr
Dim BitmapData() As Byte
Dim FileHandle As Long
Dim Index As Long
Dim RGBPalette() As RGBTRIPLE
Dim TGAHeader As TGAHeaderStr

   BitmapHandle = 0
   FileHandle = FreeFile()
   Open FileName For Input Lock Read Write As FileHandle: Close FileHandle
   Open FileName For Binary Lock Read Write As FileHandle
      Seek FileHandle, 1
      Get FileHandle, , TGAHeader
   
      If IsTGA(TGAHeader) Then
         If TGAHeader.EntryBits = 24 Then
            ReDim RGBPalette(0 To TGAHeader.ColorCount - 1) As RGBTRIPLE
            Get FileHandle, , RGBPalette()
            
            With Bitmap
               For Index = LBound(.Palette()) To UBound(.Palette())
                  .Palette(Index).rgbBlue = RGBPalette(Index).rgbtBlue
                  .Palette(Index).rgbGreen = RGBPalette(Index).rgbtGreen
                  .Palette(Index).rgbRed = RGBPalette(Index).rgbtRed
                  .Palette(Index).rgbReserved = 0
               Next Index
            End With
         End If
      
         ReDim BitmapData(0 To LOF(FileHandle) - Len(TGAHeader)) As Byte
         Get FileHandle, , BitmapData()
      End If
   Close FileHandle
   
   If IsTGA(TGAHeader) Then
      With TGAHeader
         Bitmap.Header.biWidth = .ImageWidth - .LowerLeftCornerX
         Bitmap.Header.biHeight = .ImageHeight - .LowerLeftCornerY
         If (.Descriptor And ORIENTATION_BIT) = ORIENTATION_BIT Then Bitmap.Header.biHeight = -Bitmap.Header.biHeight
         Bitmap.Header.biBitCount = TGAHeader.BitsPerPixel
      End With
   
      BitmapData() = DecompressTGA(BitmapData(), Bitmap, TGAHeader)
      BitmapData() = PaddBitmap(BitmapData(), Bitmap, CLng(TGAHeader.BitsPerPixel))
      BitmapHandle = CreateBitmap(Bitmap, BitmapData())
   End If
   
   LoadTGA = Bitmap
End Function

'This procedure padds the specified bitmap data if necessary.
Private Function PaddBitmap(UnpaddedData() As Byte, Bitmap As BitmapStr, BitsPerPixel As Long) As Byte()
Dim BytesPerPaddedRow As Long
Dim BytesPerUnpaddedRow As Long
Dim Index As Long
Dim PaddedData() As Byte
Dim PaddedIndex As Long

   BytesPerUnpaddedRow = Bitmap.Header.biWidth * (BitsPerPixel \ 8)
   
   If BytesPerUnpaddedRow Mod 4 = 0 Then
      PaddedData() = UnpaddedData()
   Else
      BytesPerPaddedRow = ((BytesPerUnpaddedRow \ 4) * 4) + 4
      
      PaddedIndex = 0
      ReDim PaddedData(0 To BytesPerPaddedRow * Bitmap.Header.biHeight) As Byte
      For Index = LBound(PaddedData()) To UBound(UnpaddedData()) - BytesPerUnpaddedRow Step BytesPerUnpaddedRow
         RtlMoveMemory PaddedData(PaddedIndex), UnpaddedData(Index), BytesPerUnpaddedRow
         PaddedIndex = PaddedIndex + BytesPerPaddedRow
      Next Index
   End If
   
   PaddBitmap = PaddedData()
End Function


