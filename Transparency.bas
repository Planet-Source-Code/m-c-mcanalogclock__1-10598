Attribute VB_Name = "Transparency"
   'this module:
   ' (c) 1999 Hobbit (hobbz@ncweb.com)
   
   'API Constants, Types, and Functions (Declares)
   Public Const SRCCOPY = &HCC0020
   Private Const SRCINVERT = &H660046
   Private Const SRCAND = &H8800C6
   Private Const CCHDEVICENAME = 32
   Private Const CCHFORMNAME = 32


   Private Type DEVMODE
       dmDeviceName As String * CCHDEVICENAME
       dmSpecVersion As Integer
       dmDriverVersion As Integer
       dmSize As Integer
       dmDriverExtra As Integer
       dmFields As Long
       dmOrientation As Integer
       dmPaperSize As Integer
       dmPaperLength As Integer
       dmPaperWidth As Integer
       dmScale As Integer
       dmCopies As Integer
       dmDefaultSource As Integer
       dmPrintQuality As Integer
       dmColor As Integer
       dmDuplex As Integer
       dmYResolution As Integer
       dmTTOption As Integer
       dmCollate As Integer
       dmFormName As String * CCHFORMNAME
       dmUnusedPadding As Integer
       dmBitsPerPel As Long
       dmPelsWidth As Long
       dmPelsHeight As Long
       dmDisplayFlags As Long
       dmDisplayFrequency As Long
       End Type


   Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


   Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


   Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long


   Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long


   Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long


   Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long


   Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long


   Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long


   Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

                                                   

   Public Function TransBitBlt(ByVal hDstDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, Optional ByVal dwRop As Long = 0) As Boolean


       'Purpose Extends capabilites of BitBlt to inlude
       ' transparency. This function will treat all
       ' pure black pixels in the source image as
       ' transparent.
       '
       
       'NotesLike BitBlt, it is necessary to call the "Refresh"
       ' method of your destination control after using this
       ' function. Until you call "Refresh", the transparent
       ' image will not appear.
       ' For Example, if your souce was called picSprite,
       'and your destination was picBack:
       'your code might look like this
       ' Call TransBitBlt(Form1.PicBack.hdc, 0, 0, 40, 40, Form1.PicSpri
       '     te.hdc, 320, 210)
       ' PicBack.Refresh
       
       'Inputs
       ' hDstDC -- The destination hDC to copy to
       ' X,Y-- The top-left point in destination to copy to
       ' nWidth, nHeight -- The size of the area to be copied
       ' hSrcDC -- The source hDC to copy from
       ' xSrc, ySrc-- The top-left point in the source to start copying
       '     from
       ' dwRop -- NOT USED (Included for compatibility w/ BitBlt code)
       
       'Outputs
       ' True -- Operation was successful
       ' False -- Operation failed
       '
       
       'Variables
       Dim MaskDC As Long 'Holds the DC For the mask
       Dim MaskBitmap As Long 'Holds the bitmap reference For the mask
       
       
       On Error GoTo Err 'If there is any error, goto Err
       MaskDC = CreateCompatibleDC(hSrcDC) 'Get a DC
      

       If MaskDC Then 'If successful in getting a DC...
           MaskBitmap = CreateBitmap(nWidth, nHeight, 1, 1, 0&) 'Get a bitmap, same size as Src, 1 bit/pixel, 1 colour plane, don't initialize.


           If MaskBitmap Then 'If successful in getting a bitmap...
               
               MaskBitmap = SelectObject(MaskDC, MaskBitmap) 'Select 2 colour bitmap into DC.
               
               Call SetBkColor(hSrcDC, QBColor(0)) 'Set the sources background color to black
               Call BitBlt(MaskDC, 0, 0, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCCOPY) 'Copy the source to the monochrome mask
               Call SetBkColor(hDstDC, QBColor(15)) 'Set the destinations background color to white
               
           Else 'If unsuccessful in getting a bitmap..
               
               Call DeleteDC(MaskDC) 'Free the DC
               MaskDC = 0 'Set the DC reference to 0
               GoTo Err 'Goto Error handler
               
           End If 'End bitmap success conditional


       Else 'If unsuccessful in getting a DC
           MaskDC = 0 'Set the reference to 0
           GoTo Err 'Goto Error handler
       End If 'End DC Success conditional


       
       Call BitBlt(hDstDC, X, Y, nWidth, nHeight, MaskDC, 0, 0, SRCAND) 'AND mask With Dst.
       Call BitBlt(hDstDC, X, Y, nWidth, nHeight, hSrcDC, xSrc, ySrc, SRCINVERT) 'XOR Src With Dst
       iMaskBitmap = SelectObject(MaskDC, MaskBitmap) 'Select the bitmap into the DC.
       DeleteObject MaskBitmap  'Free the bitmap
       DeleteDC MaskDC 'Free the DC
    
   
           
       TransBitBlt = True 'Return True (No Error)
       Exit Function 'Exit the function
       
Err:        'Error Handler
       TransBitBlt = False 'Return False (Error)
 
   End Function


