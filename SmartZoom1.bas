Attribute VB_Name = "SmartZoom1"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Public swApp                   As SldWorks.SldWorks
Public swPart                  As SldWorks.ModelDoc2
Public swDraw                  As SldWorks.DrawingDoc
Public swSheet                 As SldWorks.Sheet

'SW does everything in meters, but the sheetsize is in inches, and so is our
'BORDER_SIZE
Public CONVERSION_FACTOR As Double

'how much do we want exposed around the edge of the sheet
Public BORDER_SIZE As Double
    
Sub main()
    'These vars are meant to define a square by two points: upper left, and lower
    'right.
    Dim ulCornX As Double
    Dim ulCornY As Double
    Dim lrCornX As Double
    Dim lrCornY As Double
    
    Dim vSheetProps             As Variant
    
    Set swApp = Application.SldWorks
    Set swPart = swApp.ActiveDoc
    
    CONVERSION_FACTOR = 0.0254

    If swPart.GetType = swDocumentTypes_e.swDocDRAWING Then
        Set swDraw = swApp.ActiveDoc
        Set swSheet = swDraw.GetCurrentSheet
        vSheetProps = swSheet.GetProperties
        
        BORDER_SIZE = 0.1

        'Calculating the two points that will describe our zoom box
        ulCornX = (0 - BORDER_SIZE) * CONVERSION_FACTOR
        ulCornY = vSheetProps(6) + (BORDER_SIZE * CONVERSION_FACTOR)
        lrCornX = vSheetProps(5) + (BORDER_SIZE * CONVERSION_FACTOR)
        lrCornY = (0 - BORDER_SIZE) * CONVERSION_FACTOR
    
        'Actually do the zooming
        swPart.ViewZoomTo2 ulCornX, ulCornY, 0, lrCornX, lrCornY, 0
    Else
        'if we're looking at a model, let's do something useful, rather than
        'throwing an error.
        swPart.ViewZoomtofit2
    End If
End Sub
