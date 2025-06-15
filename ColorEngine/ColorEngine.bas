Attribute VB_Name = "ColorEngine"
' ========================================
' ColorEngine v1.0.0 RGB to other stuff ig
' RokyBeast/@rokybeast - GitHub
' ========================================

' === Required for Color Picker.

#If VBA7 Then
    Private Declare PtrSafe Function ChooseColorA Lib "comdlg32.dll" (pChoosecolor As CHOOSECOLOR) As Long
    Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

    Private Type CHOOSECOLOR
        lStructSize As Long
        hwndOwner As LongPtr
        hInstance As LongPtr
        rgbResult As Long
        lpCustColors As LongPtr
        Flags As Long
        lCustData As LongPtr
        lpfnHook As LongPtr
        lpTemplateName As String
    End Type
#Else
    Private Declare Function ChooseColorA Lib "comdlg32.dll" (pChoosecolor As CHOOSECOLOR) As Long
    Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long

    Private Type CHOOSECOLOR
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        Flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
#End If


Public Function toRGB(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte) As Long
    toRGB = CLng(b) * 65536 + CLng(g) * 256 + CLng(r)
End Function

' Convert OLE to RGB Array
Public Function OLECol(ByVal oleColor As Long) As Variant
    Dim r As Byte, g As Byte, b As Byte
    r = oleColor And 255
    g = (oleColor \ 256) And 255
    b = (oleColor \ 65536) And 255
    OLECol = Array(r, g, b)
End Function

' Convert HEX string to OLE
Public Function HexToOLE(ByVal hexString As String) As Long
    Dim r As Long, g As Long, b As Long
    Dim tempHexString As String
    
    tempHexString = Replace(hexString, "#", "") ' Work with a copy
    
    If Len(tempHexString) <> 6 Then Err.Raise vbObjectError + 513, "HexToOLE", "Invalid HEX string. Must be 6 characters after removing '#'."
    
    On Error GoTo HexConversionError
    r = CLng("&H" & Mid(tempHexString, 1, 2))
    g = CLng("&H" & Mid(tempHexString, 3, 2))
    b = CLng("&H" & Mid(tempHexString, 5, 2))
    On Error GoTo 0 ' Reset error handler

    If r < 0 Or r > 255 Or g < 0 Or g > 255 Or b < 0 Or b > 255 Then
        Err.Raise vbObjectError + 516, "HexToOLE", "Hex color components out of range (0-255)."
    End If
    
    HexToOLE = toRGB(CByte(r), CByte(g), CByte(b))
    Exit Function

HexConversionError:
    Err.Raise vbObjectError + 517, "HexToOLE", "Error converting HEX components to numbers. Ensure valid hex characters (0-9, A-F)."
End Function

' Convert OLE to HEX string
Public Function oleToHex(ByVal oleColor As Long) As String
    Dim rgb As Variant
    rgb = OLECol(oleColor)
    
    oleToHex = "#" & _
        Right("0" & VBA.HEX(rgb(0)), 2) & _
        Right("0" & VBA.HEX(rgb(1)), 2) & _
        Right("0" & VBA.HEX(rgb(2)), 2)
End Function

' Convert CMYK to OLE
Public Function CMYK(ByVal c As Double, ByVal m As Double, ByVal y As Double, ByVal k As Double) As Long
    Dim rD As Double, gD As Double, bD As Double
    
    If c < 0 Or c > 1 Or m < 0 Or m > 1 Or y < 0 Or y > 1 Or k < 0 Or k > 1 Then
        Err.Raise vbObjectError + 514, "CMYK", "CMYK values must be between 0 and 1 (inclusive)."
    End If
    
    rD = 255 * (1 - c) * (1 - k)
    gD = 255 * (1 - m) * (1 - k)
    bD = 255 * (1 - y) * (1 - k)
    
    CMYK = toRGB(CByte(Round(rD)), CByte(Round(gD)), CByte(Round(bD)))
End Function

' Convert HSL to OLE
' h = 0-360, s = 0-1, l = 0-1 (inclusive for s and l)
Public Function HSL(ByVal h As Double, ByVal s As Double, ByVal l As Double) As Long
    Dim cVal As Double, xVal As Double, mVal As Double
    Dim r1 As Double, g1 As Double, b1 As Double
    Dim r As Byte, g As Byte, b As Byte
    Dim hh As Double
    
    If h < 0 Or h > 360 Or s < 0 Or s > 1 Or l < 0 Or l > 1 Then
        Err.Raise vbObjectError + 515, "HSL", "HSL values out of range (h:0-360, s:0-1, l:0-1)."
    End If
    
    ' Normalize h to be within [0, 360)
    If h = 360 Then h = 0
    
    If s = 0 Then
        ' Achromatic (grey)
        r1 = l: g1 = l: b1 = l
    Else
        cVal = (1 - Abs(2 * l - 1)) * s
        hh = h / 60
        xVal = cVal * (1 - Abs(hh Mod 2 - 1))
        
        Select Case Int(hh) ' Floor of hh
            Case 0
                r1 = cVal: g1 = xVal: b1 = 0
            Case 1
                r1 = xVal: g1 = cVal: b1 = 0
            Case 2
                r1 = 0: g1 = cVal: b1 = xVal
            Case 3
                r1 = 0: g1 = xVal: b1 = cVal
            Case 4
                r1 = xVal: g1 = 0: b1 = cVal
            Case 5 ' Covers 5 to <6
                r1 = cVal: g1 = 0: b1 = xVal
            Case Else ' Should not happen if h is 0-359.99...
                r1 = 0: g1 = 0: b1 = 0
        End Select
    End If
    
    mVal = l - (cVal / 2)

    If s = 0 Then
        r = CByte(Round(l * 255))
        g = CByte(Round(l * 255))
        b = CByte(Round(l * 255))
    Else
        r = CByte(Round((r1 + mVal) * 255))
        g = CByte(Round((g1 + mVal) * 255))
        b = CByte(Round((b1 + mVal) * 255))
    End If
        
    HSL = toRGB(r, g, b)
End Function

' Returns selected OLE color, or -1 if cancelled
Public Function PickColor(Optional ByVal defaultColor As Long = -1) As Long
    Dim cc As CHOOSECOLOR
    Dim customColors(16) As Long
    Dim success As Long

    cc.lStructSize = LenB(cc)
    cc.hwndOwner = 0 ' Set this to Application.hwnd for Excel
    cc.hInstance = 0
    cc.lpCustColors = VarPtr(customColors(0))
    cc.rgbResult = IIf(defaultColor = -1, RGB(0, 0, 0), defaultColor)
    cc.Flags = 0

    success = ChooseColorA(cc)

    If success <> 0 Then
        PickColor = cc.rgbResult
    Else
        PickColor = -1
    End If
End Function