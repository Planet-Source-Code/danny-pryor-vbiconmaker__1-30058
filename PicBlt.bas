Attribute VB_Name = "PicBlt"

Option Explicit
'
' Win32 API Declarations, Structures, and Constants
'
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xsrc As Long, ByVal ysrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetObj Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'
' Ternary raster operations
'
Private Const SRCCOPY = &HCC0020         ' (DWORD) dest = source
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Private Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest
Private Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)

Public Sub pic_blt(ByRef frmSource As PictureBox, _
                    ByRef picSource As PictureBox, _
                    ByRef picDest As PictureBox, _
                    Optional ByVal transcolor As Long)
                    
   '
   '*************************************************************************
   '  Sub pic_blt
   '*************************************************************************
   '
   '
   ' Parameters
   '
   '    frmSource:  Source form on which animation occurs
   '    picSource:  Source picturebox containing bitmap picture
   '    picDest:    Destination picturebox
   '    transcolor: Color to use as transparent color
   
   Dim OrigColor As Long        ' Holds original background color
   Dim xDest As Long            ' location value in pixels of destination picturebox (x)
   Dim yDest As Long            ' location value in pixels of destination picturebox (y)
   Dim X As Long                ' location value in pixels of source pixel blt box
   Dim Y As Long                ' location value in pixels of source pixel blt box
   Dim nWidth As Long           ' width of destination bitmap
   Dim nHeight As Long          ' height of destimation bitmap
   Dim saveDC As Long           ' Backup copy of source bitmap
   Dim maskDC As Long           ' Mask bitmap (monochrome)
   Dim invDC As Long            ' Inverse of mask bitmap (monochrome)
   Dim resultDC As Long         ' Combination of source bitmap & background
   Dim origDC As Long           ' DC for original place cuttout on form before blitting
      
   Dim hSaveBmp As Long         ' Bitmap stores backup copy of source bitmap
   Dim hMaskBmp As Long         ' Bitmap stores mask (monochrome)
   Dim hInvBmp As Long          ' Bitmap holds inverse of mask (monochrome)
   Dim hResultBmp As Long       ' Bitmap combination of source & background
   Dim hOrigBmp As Long         ' Bitmap that stores the form cuttout bitmap
   Dim hSavePrevBmp As Long     ' Holds previous bitmap in saved DC
   Dim hMaskPrevBmp As Long     ' Holds previous bitmap in the mask DC
   Dim hInvPrevBmp As Long      ' Holds previous bitmap in inverted mask DC
   Dim hDestPrevBmp As Long     ' Holds previous bitmap in destination DC
   Dim hOrigPrevBmp As Long     ' Holds previous bitmap in original DC
   
   Dim Start As Single
   Dim I, J As Integer
   
'
' set up properties for the animation objects
' NOTE:
' the following permanently
' changes the properties of the form and picbox objects.  You may need to insert
' code at the end to set the properties back to what they were originally
'
    With frmSource
        .ScaleMode = vbPixels
        .AutoRedraw = False
    End With
   
    With picSource
        .ScaleMode = vbPixels
        .AutoRedraw = True
    End With
   
    With picDest
        .ScaleMode = vbPixels
        .AutoRedraw = False
        nWidth = .ScaleWidth
        nHeight = .ScaleHeight
        xDest = .Left
        yDest = .Top
    End With
  '============================================================
  ' Create the DC's and bitmaps for regular and transparent blt
  '============================================================
  '
  ' Create Device Handles (DC's)
  '
    resultDC = CreateCompatibleDC(picDest.hDC)
    origDC = CreateCompatibleDC(picDest.hDC)
  '
  ' Creat color bitmaps
  '
    hResultBmp = CreateCompatibleBitmap(picSource.hDC, nWidth, nHeight)
    hOrigBmp = CreateCompatibleBitmap(picSource.hDC, nWidth, nHeight)
  '
  ' select in objects
  '
    hDestPrevBmp = SelectObject(resultDC, hResultBmp)
    hOrigPrevBmp = SelectObject(origDC, hOrigBmp)
  '
  
  '===========================================================
  ' Create DC's and bitmaps for transparent blt only
  '===========================================================
    saveDC = CreateCompatibleDC(picDest.hDC)
    maskDC = CreateCompatibleDC(picDest.hDC)
    invDC = CreateCompatibleDC(picDest.hDC)
    
  '
  ' Create monochrome bitmaps for the mask-related bitmaps.
  '
    hMaskBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
    hInvBmp = CreateBitmap(nWidth, nHeight, 1, 1, ByVal 0&)
    
  '
  ' Create a color bitmap for intermediate transparent blt
  '
    hSaveBmp = CreateCompatibleBitmap(picSource.hDC, nWidth, nHeight)
  '
  ' Select bitmaps into DCs.
  '
    hSavePrevBmp = SelectObject(saveDC, hSaveBmp)
    hMaskPrevBmp = SelectObject(maskDC, hMaskBmp)
    hInvPrevBmp = SelectObject(invDC, hInvBmp)
    
  ' ==================================================================
  ' create transparent bitmap
  ' ==================================================================
  '
  ' Create mask: set background color of source to transparent color.
  '
    OrigColor = SetBkColor(picSource.hDC, transcolor)
    Call BitBlt(maskDC, 0, 0, nWidth, nHeight, picSource.hDC, 0, 0, vbSrcCopy)
    transcolor = SetBkColor(picSource.hDC, OrigColor)
  '
  ' Create inverse of mask to AND w/ source & combine w/ background.
  '
    Call BitBlt(invDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbNotSrcCopy)
  '
  ' Copy background bitmap to result and original
  '
    picDest.Visible = False
    'Form1.picContainer.Refresh
    picDest.Refresh
    frmSource.Refresh
    Call BitBlt(resultDC, 0, 0, nWidth, nHeight, frmSource.hDC, xDest, yDest, vbSrcCopy)
    Call BitBlt(origDC, 0, 0, nWidth, nHeight, frmSource.hDC, xDest, yDest, vbSrcCopy)
  '
  ' AND mask bitmap w/ result DC to punch hole in the background by
  ' painting black area for non-transparent portion of source bitmap.
  '
    Call BitBlt(resultDC, 0, 0, nWidth, nHeight, maskDC, 0, 0, vbSrcAnd)
  '
  ' get overlapper
  '
    Call BitBlt(saveDC, 0, 0, nWidth, nHeight, picSource.hDC, 0, 0, vbSrcCopy)
  '
  ' AND with inverse monochrome mask
  '
    Call BitBlt(saveDC, 0, 0, nWidth, nHeight, invDC, 0, 0, vbSrcAnd)
  '
  ' XOR these two
  '
    Call BitBlt(resultDC, 0, 0, nWidth, nHeight, saveDC, 0, 0, vbSrcInvert)
    
  '
  ' blt in the whole background into the transparent picturebox
  '
    picDest.Visible = True
    picDest.Refresh
    
    Call BitBlt(picDest.hDC, 0, 0, nWidth, nHeight, origDC, 0, 0, vbSrcCopy)

  ' blt the final bitmap in to the whole pic box to display the whole thing
  '
    Call BitBlt(picDest.hDC, 0, 0, nWidth, nHeight, resultDC, 0, 0, vbSrcCopy)
 picDest.Visible = True
 '' picDest.Refresh
  '====================================================================
 
  '==========================================================================
  ' Cleanup by deallocating memory resources
  '==========================================================================
  '
  Call SelectObject(resultDC, hDestPrevBmp)
  Call SelectObject(origDC, hOrigPrevBmp)
  Call DeleteObject(hResultBmp)
  Call DeleteObject(hOrigBmp)
  Call DeleteDC(resultDC)
  Call DeleteDC(origDC)
  '
    Call SelectObject(saveDC, hSavePrevBmp)
    Call SelectObject(maskDC, hMaskPrevBmp)
    Call SelectObject(invDC, hInvPrevBmp)
    
  
    Call DeleteObject(hSaveBmp)
    Call DeleteObject(hMaskBmp)
    Call DeleteObject(hInvBmp)
    
      
    Call DeleteDC(saveDC)
    Call DeleteDC(maskDC)
    Call DeleteDC(invDC)
  
End Sub

