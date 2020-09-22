Attribute VB_Name = "color_mod"

Option Explicit

Public Const HSLMAX As Integer = 240 '***
    'H, S and L values can be 0 - HSLMAX
    '240 matches what is used by MS Win;
    'any number less than 1 byte is OK;
    'works best if it is evenly divisible by 6
Const RGBMAX As Integer = 255 '***
    'R, G, and B value can be 0 - RGBMAX
Const UNDEFINED As Integer = (HSLMAX * 2 / 3) '***
    'Hue is undefined if Saturation = 0 (greyscale)

Public Type HSLCol 'Datatype used to pass HSL Color values
    Hue As Integer
    Sat As Integer
    Lum As Integer
End Type

Public lSelected As Long '//variable for selected color
Public selIndex As Integer
Public bLoad As Boolean
Public lExportCol As String


Public Function HexRGB(lCdlColor As Long)
Dim lCol As Long
Dim iRed As Integer
Dim iGreen As Integer
Dim iBlue As Integer
Dim vHexR As String
Dim vHexG As String
Dim vHexB As String
    'Break out the R, G, B values from the c
    '     ommon dialog color
    lCol = lCdlColor
    iRed = lCol Mod &H100
    lCol = lCol \ &H100
    iGreen = lCol Mod &H100
    lCol = lCol \ &H100
    iBlue = lCol Mod &H100
    
    'Determine Red Hex
    vHexR = Hex(iRed)

    If Len(vHexR) < 2 Then
        vHexR = "0" & vHexR
    End If
    'Determine Green Hex
    vHexG = Hex(iGreen)

    If Len(vHexG) < 2 Then
        vHexG = "0" & iGreen
    End If
    'Determine Blue Hex
    vHexB = Hex(iBlue)

    If Len(vHexB) < 2 Then
        vHexB = "0" & vHexB
    End If
    'Add it up, return the function value
    HexRGB = vHexR & vHexG & vHexB
End Function

Public Function RGBtoHSL(RGBCol As Long) As HSLCol '***
'Returns an HSLCol datatype containing Hue, Luminescence
'and Saturation; given an RGB Color value

Dim R As Integer, g As Integer, b As Integer
Dim cMax As Integer, cMin As Integer
Dim RDelta As Double, GDelta As Double, _
    BDelta As Double
Dim h As Double, s As Double, L As Double
Dim cMinus As Long, cPlus As Long
    
    R = rgbRed(RGBCol)
    g = rgbGreen(RGBCol)
    b = rgbBlue(RGBCol)
    
    cMax = iMax(iMax(R, g), b) 'Highest and lowest
    cMin = iMin(iMin(R, g), b) 'color values
    
    cMinus = cMax - cMin 'Used to simplify the
    cPlus = cMax + cMin  'calculations somewhat.
    
    'Calculate luminescence (lightness)
    L = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)
    
    If cMax = cMin Then 'achromatic (r=g=b, greyscale)
        s = 0 'Saturation 0 for greyscale
        h = UNDEFINED 'Hue undefined for greyscale
    Else
        'Calculate color saturation
        If L <= (HSLMAX / 2) Then
            s = ((cMinus * HSLMAX) + 0.5) / cPlus
        Else
            s = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
        End If
    
        'Calculate hue
        RDelta = (((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus
        GDelta = (((cMax - g) * (HSLMAX / 6)) + 0.5) / cMinus
        BDelta = (((cMax - b) * (HSLMAX / 6)) + 0.5) / cMinus
    
        Select Case cMax
            Case CLng(R)
                h = BDelta - GDelta
            Case CLng(g)
                h = (HSLMAX / 3) + RDelta - BDelta
            Case CLng(b)
                h = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select
        
        If h < 0 Then h = h + HSLMAX
    End If
    
    RGBtoHSL.Hue = CInt(h)
    RGBtoHSL.Lum = CInt(L)
    RGBtoHSL.Sat = CInt(s)
End Function

Private Function iMax(A As Integer, b As Integer) _
    As Integer
'Return the Larger of two values
    iMax = IIf(A > b, A, b)
End Function
Private Function iMin(A As Integer, b As Integer) _
    As Integer
'Return the smaller of two values
    iMin = IIf(A < b, A, b)
End Function
Public Function rgbGreen(RGBCol As Long) As Integer
'Return the Green component from an RGB Color
    rgbGreen = ((RGBCol And &H100FF00) / &H100)
End Function
Public Function rgbBlue(RGBCol As Long) As Integer
'Return the Blue component from an RGB Color
    rgbBlue = (RGBCol And &HFF0000) / &H10000
End Function
Public Function rgbRed(RGBCol As Long) As Integer
'Return the Red component from an RGB Color
    rgbRed = RGBCol And &HFF
End Function


