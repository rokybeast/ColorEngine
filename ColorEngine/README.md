# ColorEngine v1.0.0 Documentation

> A comprehensive VBA color conversion and manipulation library

## Overview

ColorEngine is a VBA module that provides color conversion functions between different color spaces (RGB, HEX, CMYK, HSL) and includes a native Windows color picker dialog. The module is compatible with both VBA6 and VBA7 environments.

## Installation

1. Open your VBA editor (Alt + F11)
2. Right-click on your project in the Project Explorer
3. Select **Insert > Module**
4. Copy and paste the ColorEngine code into the new module
5. Save your project

## API Reference

### Core Functions

#### `toRGB(r As Byte, g As Byte, b As Byte) As Long`

Converts individual RGB components to OLE color format.

**Parameters:**
- `r` - Red component (0-255)
- `g` - Green component (0-255)  
- `b` - Blue component (0-255)

**Returns:** Long - OLE color value

**Example:**
```vb
Dim oleColor As Long
oleColor = toRGB(255, 128, 64) ' Orange color
```

#### `OLECol(oleColor As Long) As Variant`

Converts OLE color to RGB array.

**Parameters:**
- `oleColor` - OLE color value

**Returns:** Variant - Array containing [R, G, B] values

**Example:**
```vb
Dim rgbArray As Variant
rgbArray = OLECol(255) ' Returns [255, 0, 0] for red
Debug.Print "Red: " & rgbArray(0)
Debug.Print "Green: " & rgbArray(1)
Debug.Print "Blue: " & rgbArray(2)
```

### Color Space Conversions

#### `HexToOLE(hexString As String) As Long`

Converts HEX color string to OLE color format.

**Parameters:**
- `hexString` - HEX color string (with or without #)

**Returns:** Long - OLE color value

**Throws:**
- Error 513: Invalid HEX string length
- Error 516: Color components out of range
- Error 517: Invalid HEX characters

**Example:**
```vb
Dim oleColor As Long
oleColor = HexToOLE("#FF8040") ' Orange
oleColor = HexToOLE("FF8040")  ' Also works without #
```

#### `oleToHex(oleColor As Long) As String`

Converts OLE color to HEX string format.

**Parameters:**
- `oleColor` - OLE color value

**Returns:** String - HEX color string with # prefix

**Example:**
```vb
Dim hexColor As String
hexColor = oleToHex(255) ' Returns "#FF0000" for red
```

#### `CMYK(c As Double, m As Double, y As Double, k As Double) As Long`

Converts CMYK values to OLE color format.

**Parameters:**
- `c` - Cyan component (0.0-1.0)
- `m` - Magenta component (0.0-1.0)
- `y` - Yellow component (0.0-1.0)
- `k` - Black component (0.0-1.0)

**Returns:** Long - OLE color value

**Throws:**
- Error 514: CMYK values out of range

**Example:**
```vb
Dim oleColor As Long
oleColor = CMYK(0.0, 0.5, 0.75, 0.0) ' Orange color
```

#### `HSL(h As Double, s As Double, l As Double) As Long`

Converts HSL values to OLE color format.

**Parameters:**
- `h` - Hue (0-360 degrees)
- `s` - Saturation (0.0-1.0)
- `l` - Lightness (0.0-1.0)

**Returns:** Long - OLE color value

**Throws:**
- Error 515: HSL values out of range

**Example:**
```vb
Dim oleColor As Long
oleColor = HSL(30, 1.0, 0.625) ' Orange color
```

### Color Picker

#### `PickColor([defaultColor As Long]) As Long`

Opens the native Windows color picker dialog.

**Parameters:**
- `defaultColor` (Optional) - Initial color selection (default: black)

**Returns:** 
- Long - Selected OLE color value
- -1 if user cancelled the dialog

**Example:**
```vb
Dim selectedColor As Long
selectedColor = PickColor() ' Opens with black default

If selectedColor <> -1 Then
    Debug.Print "Selected color: " & oleToHex(selectedColor)
Else
    Debug.Print "User cancelled color selection"
End If

' With default color
selectedColor = PickColor(RGB(255, 0, 0)) ' Opens with red default
```

## Usage Examples

### Basic Color Conversion Workflow

```vb
Sub ColorConversionExample()
    Dim originalColor As Long
    Dim hexColor As String
    Dim rgbArray As Variant
    
    ' Create a color from RGB components
    originalColor = toRGB(255, 128, 64) ' Orange
    
    ' Convert to HEX
    hexColor = oleToHex(originalColor)
    Debug.Print "HEX: " & hexColor ' Output: #FF8040
    
    ' Extract RGB components
    rgbArray = OLECol(originalColor)
    Debug.Print "RGB: " & rgbArray(0) & ", " & rgbArray(1) & ", " & rgbArray(2)
    
    ' Convert from different color spaces
    Dim cmykColor As Long
    Dim hslColor As Long
    
    cmykColor = CMYK(0.0, 0.5, 0.75, 0.0) ' Orange in CMYK
    hslColor = HSL(30, 1.0, 0.625)        ' Orange in HSL
End Sub
```

### Interactive Color Selection

```vb
Sub InteractiveColorSelection()
    Dim userColor As Long
    Dim hexValue As String
    
    ' Let user pick a color
    userColor = PickColor(RGB(128, 128, 128)) ' Default to gray
    
    If userColor <> -1 Then
        hexValue = oleToHex(userColor)
        MsgBox "You selected: " & hexValue
        
        ' Use the color (example: set cell background)
        ActiveCell.Interior.Color = userColor
    Else
        MsgBox "No color selected"
    End If
End Sub
```

### Excel Integration Example

```vb
Sub ColorizeExcelCells()
    Dim ws As Worksheet
    Dim i As Integer
    Dim cellColor As Long
    
    Set ws = ActiveSheet
    
    ' Create a rainbow effect
    For i = 1 To 10
        cellColor = HSL(i * 36, 1.0, 0.5) ' Full saturation, medium lightness
        ws.Cells(1, i).Interior.Color = cellColor
        ws.Cells(1, i).Value = oleToHex(cellColor)
    Next i
End Sub
```

## Error Handling

The module includes comprehensive error handling with specific error codes:

- **513**: Invalid HEX string format
- **514**: CMYK values out of range
- **515**: HSL values out of range  
- **516**: HEX color components out of range
- **517**: Invalid HEX characters

Example error handling:
```vb
Sub SafeColorConversion()
    On Error GoTo ErrorHandler
    
    Dim result As Long
    result = HexToOLE("InvalidHex")
    
    Exit Sub
    
ErrorHandler:
    Select Case Err.Number - vbObjectError
        Case 513
            MsgBox "Invalid HEX format. Use 6-character HEX codes."
        Case 517
            MsgBox "Invalid HEX characters. Use 0-9 and A-F only."
        Case Else
            MsgBox "Error: " & Err.Description
    End Select
End Sub
```

## Compatibility

- **VBA7**: Supported (Office 2010+)
- **VBA6**: Supported (Office 2007 and earlier)
- **Applications**: Excel, Word, Access, PowerPoint, and other VBA-enabled applications

## Color Space Information

### RGB (Red, Green, Blue)
- Range: 0-255 for each component
- Additive color model
- Default for computer displays

### HEX (Hexadecimal)
- Format: #RRGGBB
- Range: 00-FF for each component
- Common in web development

### CMYK (Cyan, Magenta, Yellow, Black)
- Range: 0.0-1.0 for each component
- Subtractive color model
- Used in printing

### HSL (Hue, Saturation, Lightness)
- Hue: 0-360 degrees
- Saturation: 0.0-1.0 (0% to 100%)
- Lightness: 0.0-1.0 (0% to 100%)
- Intuitive for color adjustments

## License

MIT License - see LICENSE file for details.

## Author

Created by **RokyBeast** ([@rokybeast](https://github.com/rokybeast))

## Contributing

Issues and pull requests are welcome! Please feel free to contribute to this project.
