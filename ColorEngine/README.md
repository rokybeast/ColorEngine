# 🎨 ColorEngine for VBA

A powerful and lightweight VBA module for working with colors in different formats — RGB, HEX, HSL, CMYK, and OLE — with a native **Windows color picker** using the `ChooseColorA` API.

> 💻 Created by [@rokybeast](https://github.com/rokybeast) — because even VBA deserves cool tooling.

---

## ⚙️ Features

* 🔢 `toRGB(r, g, b)` → Convert RGB components to OLE color
* 🌈 `OLECol(ole)` → Convert OLE to RGB array
* 🧾 `HexToOLE("#rrggbb")` → HEX string to OLE
* 🧪 `oleToHex(ole)` → OLE to HEX string
* 💦 `CMYK(c, m, y, k)` → CMYK values to OLE
* 🌡️ `HSL(h, s, l)` → HSL to OLE conversion
* 🎨 `PickColor()` → Native Windows color picker using WinAPI (`ChooseColorA`)

---

## 🧱 Requirements

* **VBA6 or VBA7 (32/64-bit)** supported
* Works in **Excel, Word, PowerPoint, Access**, etc.
* No OCX or ActiveX dependencies
* No external files needed

---

## 💠 Installation

1. Open the VBA Editor (`ALT + F11`)
2. Import `ColorEngine.bas` or paste the contents into a new module
3. Done! You can now call the color functions in your VBA code.

---

## 🚀 Usage Examples

### Convert HEX to OLE

```vba
Dim oleColor As Long
oleColor = HexToOLE("#3498db")
```

### Convert OLE to RGB Array

```vba
Dim rgbArr As Variant
rgbArr = OLECol(oleColor) ' rgbArr(0)=R, rgbArr(1)=G, rgbArr(2)=B
```

### Use the Native Color Picker

```vba
Dim pickedColor As Long
pickedColor = PickColor()

If pickedColor <> -1 Then
    MsgBox "You picked: " & oleToHex(pickedColor)
End If
```

---

## 🧠 API Safety

The color picker uses conditional `Declare` statements and works with both:

* **VBA7/64-bit**: via `PtrSafe` and `LongPtr`
* **VBA6/32-bit**: via legacy `Long`

No crashing. No weird compatibility issues.

---

## 🦪 License

MIT — because sharing is caring 🫡

---

## 🥘 Built With Love by RokyBeast

Because every dev deserves access to clean color functions — even if you're coding in **VBA in 2025**.
No judgement. Only power. 🚀