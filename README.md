# NextHikes_Project_2 – Advanced Calculator

A desktop calculator built with **Python** and **Tkinter** as **Project 2** of **NextHikes IT Solutions**.

The project requirement was a basic calculator (display screen, clickable buttons, core operations, equals and clear). This version goes much further: scientific functions, history, a unit and currency converter, a maths graph, a data graph, a dark and light theme and a resizable window. Every key and screen is drawn in code on a Tkinter `Canvas`, so the look is fully custom.

**Author:** MD SHAHNSHAH  
**Company:** NextHikes IT Solutions

---

## Table of Contents

- [Features](#features)
- [Screenshots](#screenshots)
- [Requirements](#requirements)
- [How to Run](#how-to-run)
- [Keyboard Shortcuts](#keyboard-shortcuts)
- [How It Works](#how-it-works)
- [Challenges and Fixes](#challenges-and-fixes)
- [Project Structure](#project-structure)
- [Future Improvements](#future-improvements)
- [License](#license)

---

## Features

### Basic calculator (Home)
- Add, subtract, multiply, divide and equals, with a **live result preview** while typing
- **Phone-style percent:** typing `200+10%` gives `220`
- Memory keys: `MC`, `MR`, `M+`, `M−`, `MS` and `Ans`
- Delete one key (`⌫`) or clear all (`C`), plus a sign change key (`±`)

### Scientific tab
- Square, cube, `x^y`, square root, cube root, `y`-th root, `e^x`, `10^x`, `2^x`
- `ln`, `log`, log of any base (`logᵧx`) and factorial `x!`
- Constants `π` and `e`
- All bracket types work: round `( )`, square `[ ]` and curly `{ }`

### Trigo tab
- `sin`, `cos`, `tan`, inverse (`sin⁻¹`, `cos⁻¹`, `tan⁻¹`) and hyperbolic (`sinh`, `cosh`, `tanh`)
- `sec`, `csc`, `cot` and radian ↔ degree converters
- **RAD / DEG** mode, shown on the LCD

### Func tab
- Money tools: **GST add** (`gst`), **GST remove** (`gstr`), **discount** (`disc`), **EMI** (`emi`), **% change** (`chg`)
- Maths helpers: `nPr`, `nCr`, `mod`, `gcd`, `lcm`, `abs`, `round`, `floor`, `ceil`, `max`, `min`, `log`, `log₂`
- **Argument hints:** a hint line shows which value to type next, for example `gst(amount, GST %)`

### History
- Every result is saved with its expression
- Click an old result to put it back on the screen
- Scroll with the mouse wheel; clear from the header

### Converter
Seven categories:

| Category | Examples |
|---|---|
| Currency | USD to INR (15 currencies) |
| Length | m to ft |
| Area | acre to m² |
| Weight | kg to lb |
| Temperature | °C, °F, K |
| Speed | km/h to mph |
| Number System | BIN, OCT, DEC, HEX |

- **Live exchange rates** from an online API, with **offline backup rates** if there is no internet
- Swap the From and To units with one click or the `Tab` key

### Maths Graph
- Type any function of `x`: `sin(x)`, `x^2`, `ln(x)`; short forms like `sinx` and `sin90` also work
- Zoom with `Zoom +`, `Zoom −`, `Fit` or the mouse wheel
- Drag to pan; hover to read `x` and `y`
- RAD / DEG switch and auto-fit view

### Data Graph
- Load data from **CSV / TXT**, **Excel (.xlsx)** or **paste** a table from the clipboard
- **10 chart types:** Bar, Stacked Bar, Horizontal Bar, Line, Area, Scatter, Pie, Donut, Histogram, Radar
- **Automatic statistics:** Count, Sum, Mean, Median, Min, Max, Std Dev, Range
- Pick which column is used for the statistics, pie, donut and histogram

### Look and feel
- **Dark and Light theme**, changed with one switch on the title bar
- 3D keys drawn pixel by pixel with depth and gloss
- Custom title bar with minimize, maximize and close buttons
- **Window mode:** the window is resizable (the whole canvas scales), and in a wide window a **History panel** opens on the right
- Smooth sliding animation between tabs
- Keyboard and clipboard support

---

## Screenshots

Add your screenshots to a `screenshots` folder and the images below will show.

| Home | Scientific | Trigo |
|---|---|---|
| ![Home](screenshots/home.png) | ![Scientific](screenshots/scientific.png) | ![Trigo](screenshots/trigo.png) |

| Func | History | Converter |
|---|---|---|
| ![Func](screenshots/func.png) | ![History](screenshots/history.png) | ![Converter](screenshots/converter.png) |

| Maths Graph | Data Graph | Dark and Light |
|---|---|---|
| ![Maths Graph](screenshots/maths_graph.png) | ![Data Graph](screenshots/data_graph.png) | ![Themes](screenshots/themes.png) |

| Window mode |
|---|
| ![Window mode](screenshots/window_mode.png) |

---

## Requirements

- **Python 3.9 or newer** (`math.lcm` is used)
- **Tkinter** (comes with the standard Python installer on Windows and macOS; on Linux install `python3-tk`)
- No third-party packages. Only the standard library is used:
  `math`, `ast`, `operator`, `re`, `decimal`, `json`, `csv`, `zipfile`, `xml.etree`, `urllib`, `threading`, `os`, `tkinter`
- Internet is needed **only** for live currency rates. Everything else works offline.

> The app is designed for **Windows** (Segoe UI font, taskbar handling). On other systems it should still run, but the font and window behaviour may look different.

---

## How to Run

```bash
# 1. Clone the repository
git clone https://github.com/shahanshah786/NextHikes_Project_2.git

# 2. Open the project folder
cd NextHikes_Project_2

# 3. Run the calculator (use your script's file name)
python advanced_calculator.py
```

---

## Keyboard Shortcuts

| Key | Action |
|---|---|
| `0–9`, `+ - * / ^ ( ) . %` | Type numbers and operators |
| `Enter` or `=` | Calculate |
| `Backspace` | Delete last key |
| `Esc` | Clear all (AC) |
| `F9` | Change sign (±) |
| `p` / `e` | π / e |
| `Ctrl + C` / `Ctrl + V` | Copy result / paste |
| `H` | Open or close History |
| `Tab` | Swap From and To units (Converter) |
| Arrow keys | Change units (Converter) or move the view (Graph) |
| `Page Up` / `Page Down` | Zoom in / out (Graph) |
| `Home` | Fit the graph |
| `Ctrl + O` | Upload a file (Data Graph) |

---

## How It Works

```
Input  →  prepare()  →  safe_eval()  →  fmt()  →  Display
(keys)   (clean text)  (compute safely) (format)   (LCD + History)
```

1. **Input:** keys or the keyboard add text to the expression.
2. **`prepare()`:** fixes symbols (`×`, `÷`, `−`, `π`), percent, factorial, brackets and missing multiplication signs.
3. **`safe_eval()`:** the text is parsed with `ast.parse()` and **only allowed nodes** (numbers, `+ - * / **` and a whitelist of functions) are run.
4. **`fmt()`:** formats the number for the screen.
5. **Display:** the result shows on the LCD and is saved in History.

### Why `ast` and not `eval()`?
Python's `eval()` can run any code, which is unsafe. This calculator never calls `eval()`. Anything that is not allowed raises a `ValueError` and the screen shows **Error**. A power limit also stops huge numbers.

---

## Challenges and Fixes

| Challenge | Problem | Fix |
|---|---|---|
| **Currency API** | Live rates did not load; the API link failed | A free key-less API, fetched in a background thread with a 6 second timeout, plus offline backup rates |
| **CSV data in Data Graph** | Uploaded data showed wrong values | Auto-detect separator and header, clean currency and percent symbols, read `.xlsx` through ZIP + XML |
| **Maths Graph** | `tan(x)` drew false vertical lines; `sinx` was not understood | `clip_y()` splits broken lines and `gprep()` adds the missing brackets |

---

## Project Structure

```
NextHikes_Project_2/
├── NEXT HIKE IT SOLUTION PROJECT/
│   ├── advanced_calculator.py      # the calculator (single script)
│   └── Advanced_Calculator.pptx    # project presentation
├── screenshots/                    # images used in this README
├── LICENSE
└── README.md
```

Inside the script:

- **Math helpers:** `fact`, `npr`, `ncr`, `gcd`, `lcm`, `gst`, `emi` and more
- **`prepare()`, `safe_eval()`, `fmt()`:** the expression pipeline
- **File readers:** `read_table()`, `read_xlsx()`, `table_to_data()`
- **`ScaledCanvas`:** a Canvas that scales everything with the window size
- **`Calculator`:** the main class (keys, tabs, converter, graphs, history, themes)

---

## Future Improvements

- Save history to a file
- More chart types and the option to save a chart as an image
- More currencies
- Support for old `.xls` files

---

## License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

---

Made with Python and Tkinter by **MD SHAHNSHAH** for **NextHikes IT Solutions**.
