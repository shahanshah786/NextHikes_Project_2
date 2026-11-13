"""
Advanced Calculator, Using Python with Tkinter Library
Features: basic/scientific/trig calculator, unit converter (live currency),
function grapher, CSV/Excel data charts, history, dark/light theme.
Run: python calculator.py 
"""
import math
import ast
import operator
import re
import json
import threading
import urllib.request
import csv
import os
import zipfile
import xml.etree.ElementTree as ET
import tkinter as tk 
from tkinter import filedialog
import tkinter.font as tkfont 
from decimal import Decimal, ROUND_HALF_UP

#Themes
BG,BODY,BODY_EDGE = "#000000", "#000000", "#1c1c1e"
LCD, INK, INK2, ERR =  "#c6d3b2" , "#1d2a17", "#58654f", "#a3201f"
FONT = "Segoe UI"
TAB_ON, TAB_OFF = "#ff3b30", "#9aa1ad"
ACCENT = "#ff9f0a"
GLOW = INK


# Label , Style , And Value
BASIC = [
    [("⌫", "top", "DEL", 1.5), ("C", "top", "AC", 1.5), ("%", "top", "%", 1.5), ("÷", "op", "÷", 1.5)],
    [("7", "num", "7", 1.5), ("8", "num", "8", 1.5), ("9", "num", "9", 1.5), ("×", "op", "×", 1.5)],
    [("4", "num", "4", 1.5), ("5", "num", "5", 1.5), ("6", "num", "6", 1.5), ("−", "op", "−", 1.5)],
    [("1", "num", "1", 1.5), ("2", "num", "2", 1.5), ("3", "num", "3", 1.5), ("+", "op", "+", 1.5)],
    [("±", "num", "NEG", 1.5), ("0", "num", "0", 1.5), (".", "num", ".", 1.5), ("=", "op", "=", 1.5)],
]

SCIENCETIFIC = [
    [("x²", "fn", "^2", 1.5), ("x³", "fn", "^3", 1.5), ("xʸ", "fn", "^", 1.5),("eˣ", "fn", "e^(", 1.5), ("10ˣ", "fn", "10^(", 1.5)],
    [("²√x", "fn", "√(", 1.5), ("³√x", "fn", "cbrt(", 1.5), ("ʸ√x", "fn", "^(1÷", 1.5),("ln", "fn", "ln(", 1.5), ("2ˣ", "fn", "2^(", 1.5)],
    [("x!", "fn", "!", 1.2), ("(", "fn", "(", 1.2), (")", "fn", ")", 1.2), ("e", "fn", "e", 1.2), ("π", "fn", "π", 1.2)],
    [("logᵧx", "fn", "LOGY", 1.2), ("[", "fn", "[", 1.2), ("]", "fn", "]", 1.2), ("{", "fn", "{", 1.2), ("}", "fn", "}", 1.2)],
]

TRIG = [
    [("sin", "fn", "sin(", 1.2), ("cos", "fn", "cos(", 1.2), ("tan", "fn", "tan(", 1.2), ("RAD", "fn", "SETRAD", 1.2), ("DEG", "fn", "SETDEG", 1.2)],
    [("sin⁻¹", "fn", "asin(", 1.2), ("cos⁻¹", "fn", "acos(", 1.2), ("tan⁻¹", "fn", "atan(", 1.2), ("π", "fn", "π", 1.2), ("e", "fn", "e", 1.2)],
    [("sinh", "fn", "sinh(", 1.2), ("cosh", "fn", "cosh(", 1.2), ("tanh", "fn", "tanh(", 1.2), ("(", "fn", "(", 1.2), (")", "fn", ")", 1.2)],
    [("sec", "fn", "sec(", 1.2), ("csc", "fn", "csc(", 1.2), ("cot", "fn", "cot(", 1.2), ("r→d", "fn", "todeg(", 1.2), ("d→r", "fn", "torad(", 1.2)],
]

MEMORY = [[("MC", "top", "MC", 1.2), ("MR", "top", "MR", 1.2), ("M+", "top", "M+", 1.2), ("M−", "top", "M−", 1.2), ("MS", "top", "MS", 1.2), ("Ans", "top", "Ans", 1.2)]]

MORE = [
    [("nPr", "fn", "npr(", 1.2), ("nCr", "fn", "ncr(", 1.2), ("mod", "fn", "mod(", 1.2), ("gcd", "fn", "gcd(", 1.2), ("lcm", "fn", "lcm(", 1.2)],
    [("abs", "fn", "abs(", 1.2), ("round", "fn", "round(", 1.2), ("floor", "fn", "floor(", 1.2), ("ceil", "fn", "ceil(", 1.2), (",", "fn", ",", 1.2)],
    [("max", "fn", "max(", 1.2), ("min", "fn", "min(", 1.2), ("log", "fn", "log(", 1.2), ("log₂", "fn", "logb(", 1.2), ("x⁻¹", "fn", "^(−1)", 1.2)],
    [("GST+", "fn", "gst(", 1.2), ("GST−", "fn", "gstr(", 1.2), ("Disc", "fn", "disc(", 1.2), ("EMI", "fn", "emi(", 1.2), ("%chg", "fn", "chg(", 1.2)],
]

HINTS = {
    "gst":   (("amount", "GST %"), "amount + GST", "gst(1000,18)=1180"),
    "gstr":  (("total", "GST %"), "price without GST", "gstr(1180,18)=1000"),
    "disc":  (("price", "discount %"), "price after discount", "disc(500,10)=450"),
    "emi":   (("loan", "yearly %", "months"), "monthly EMI", "emi(500000,9,60)"),
    "chg":   (("old", "new"), "% change", "chg(80,100)=25"),
    "npr":   (("n", "r"), "permutations", "npr(5,2)=20"),
    "ncr":   (("n", "r"), "combinations", "ncr(5,2)=10"),
    "mod":   (("a", "b"), "remainder", "mod(10,3)=1"),
    "gcd":   (("a", "b"), "HCF", "gcd(12,18)=6"),
    "lcm":   (("a", "b"), "LCM", "lcm(4,6)=12"),
    "abs":   (("x",), "ignore minus sign", "abs(−5)=5"),
    "round": (("x", "digits"), "round off", "round(3.456,1)=3.5"),
    "floor": (("x",), "round down", "floor(3.7)=3"),
    "ceil":  (("x",), "round up", "ceil(3.2)=4"),
    "max":   (("a", "b"), "bigger number", "max(4,9)=9"),
    "min":   (("a", "b"), "smaller number", "min(4,9)=4"),
    "log":   (("x",), "log base 10", "log(100)=2"),
    "logb":  (("x",), "log base 2", "logb(8)=3"),
}

DIG = "0123456789ABCDEF"
ARROWS = {"Left": "FPREV", "Right": "FNEXT", "Up": "TPREV", "Down": "TNEXT"}
PANELS = {"sci": SCIENCETIFIC, "trig": TRIG, "more": MORE}
TABS = ("home", "sci", "trig", "more", "conv", "graph", "hist")
TAB_NAMES = ("sci", "trig", "more", "conv", "graph")
TAB_WTS = (82, 50, 46, 84, 52)          
TAB_LABELS = {"sci": "Scientific", "trig": "Trigo", "more": "Func", "conv": "Converter", "graph": "Graph"}
ICONS = {"Currency": "💱", "Length": "📏", "Area": "▦", "Weight": "⚖", "Temperature": "🌡", "Speed": "⚡", "Number System": "#", "Maths Graph": "📈", "Data Graph": "📊"}

STYLES = {
  "fn" :("#2b2b2b", "#ffffff"),
  "num":("#505050", "#ffffff"),
  "top":("#a5a5a5", "#000000"),
  "op":("#ff9f0a", "#fffffe"),
}

THEMES = {
    "dark": {"BG": "#000000", "BODY": "#000000", "BODY_EDGE": "#1c1c1e", "GRIP": "#5a5a5e", "CMAP": {},"STYLES": dict(STYLES)},"light": {"BG": "#d1d1d6", "BODY": "#f2f2f7", "BODY_EDGE": "#c7c7cc", "GRIP": "#8e8e93",
    "CMAP": {"#15171b": "#d8d8dd", "#151518": "#e0e0e5", "#151517": "#ececf0", "#3a3a3c": "#c7c7cc", "#2c2c2e": "#ffffff", "#1c1c1e": "#ffffff", "#ffffff": "#1d1d1f", "#9aa1ad": "#6e6e73", "#222226": "#d1d1d6"},
    "STYLES": {"fn": ("#e4e4ea", "#1d1d1f"), "num": ("#ffffff", "#1d1d1f"),
    "top": ("#c7c7cc", "#000000"), "op": ("#ff9f0a", "#fffffe")}},
}
THEME = "dark"

def set_theme(name):
    global THEME, BG, BODY, BODY_EDGE
    t = THEMES[name]
    THEME, BG, BODY, BODY_EDGE = name, t["BG"], t["BODY"], t["BODY_EDGE"]
    STYLES.update(t["STYLES"])

W,H = 480, 780                    
MIN_S, MAX_S = 0.7, 4.0            
MINW, MINH = int(W * MIN_S), int(H * MIN_S)   
CELL, DEPTH = 70, 4
X0 = 65
HROWS, HY0, HDY = 5, 90, 20      
TABY = 214
PX0, PY0, PY1 = 24, 200, 432
GCHARS = set("abcdefghijklmnopqrstuvwxyz0123456789+-*/^().,%!")
DCHARS = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789,.:;-_ ")
GITEMS = [("Maths Graph", "GRF0"), ("Data Graph", "GRF1")]
PIE = ["#1d2a17", "#a3201f", "#c77700", "#2b6cb0", "#6b46c1", "#2f855a", "#975a16", "#b83280", "#4a5568", "#0987a0"]
ANIM_OUT, ANIM_IN = 7, 10

# Window  mode
WIDE_RATIO = 1.0                   
HP_MIN = 300                       
CW_MAX = 780                       
CW_FRAC = 0.6                     
HPN, HPDY = 9, 64                   

CURR = [("US Dollar", "USD"), ("Indian Rupee", "INR"), ("Euro", "EUR"), ("British Pound", "GBP"),
        ("Japanese Yen", "JPY"), ("UAE Dirham", "AED"), ("Australian Dollar", "AUD"),
        ("Canadian Dollar", "CAD"), ("Chinese Yuan", "CNY"), ("Singapore Dollar", "SGD"),
        ("Saudi Riyal", "SAR"), ("Pakistani Rupee", "PKR"), ("Bangladeshi Taka", "BDT"),
        ("Nepalese Rupee", "NPR"), ("Swiss Franc", "CHF")]
RATES = {"USD": 1.0, "INR": 88.0, "EUR": 0.86, "GBP": 0.75, "JPY": 150.0, "AED": 3.6725, "AUD": 1.52,
         "CAD": 1.38, "CNY": 7.1, "SGD": 1.3, "SAR": 3.75, "PKR": 281.0, "BDT": 122.0, "NPR": 141.0, "CHF": 0.8}

CATS = [ ("Currency", None, 0, 1),
    ("Length", [("Millimeter", "mm", 0.001), ("Centimeter", "cm", 0.01), ("Meter", "m", 1.0),("Kilometer", "km", 1000.0), ("Inch", "in", 0.0254), ("Foot", "ft", 0.3048),("Yard", "yd", 0.9144), ("Mile", "mi", 1609.344)], 2, 5),
    ("Area", [("Sq Millimeter", "mm²", 1e-6), ("Sq Centimeter", "cm²", 1e-4), ("Sq Meter", "m²", 1.0),
              ("Sq Kilometer", "km²", 1e6), ("Hectare", "ha", 1e4), ("Acre", "ac", 4046.8564224),
              ("Sq Foot", "ft²", 0.09290304), ("Sq Inch", "in²", 0.00064516), ("Sq Yard", "yd²", 0.83612736)], 2, 6),
    ("Weight", [("Milligram", "mg", 1e-6), ("Gram", "g", 1e-3), ("Kilogram", "kg", 1.0),
                ("Quintal", "q", 100.0), ("Tonne", "t", 1000.0), ("Ounce", "oz", 0.028349523125),("Pound", "lb", 0.45359237)], 2, 6),
    ("Temperature", [("Celsius", "°C", None), ("Fahrenheit", "°F", None), ("Kelvin", "K", None)], 0, 1),
    ("Speed", [("Meter/second", "m/s", 1.0), ("Kilometer/hour", "km/h", 1 / 3.6), ("Mile/hour", "mph", 0.44704),
               ("Knot", "kn", 0.514444444), ("Foot/second", "ft/s", 0.3048)], 1, 2),
    ("Number System", [("Binary", "BIN", 2), ("Octal", "OCT", 8), ("Decimal", "DEC", 10),("Hexadecimal", "HEX", 16)], 2, 0),
]

# Mathematical Logic
OPS = {ast.Add: operator.add, ast.Sub: operator.sub,ast.Mult: operator.mul, ast.Div: operator.truediv, ast.Pow: operator.pow}

def fact(n):
    if n < 0 or n != int(n) or n > 170:
        raise ValueError
    return math.factorial(int(n))

def cbrt(v):
    return math.copysign(abs(v) ** (1 / 3), v)

def to_int(v):
    if v != int(v):
        raise ValueError
    return int(v)

def npr(n, r):
    n, r = to_int(n), to_int(r)
    if n < 0 or r < 0 or r > n or n > 1000:
        raise ValueError
    return math.perm(n, r)

def ncr(n, r):
    n, r = to_int(n), to_int(r)
    if n < 0 or r < 0 or r > n or n > 1000:
        raise ValueError
    return math.comb(n, r)

def gcd(a, b):
    return math.gcd(to_int(a), to_int(b))

def lcm(a, b):
    return math.lcm(to_int(a), to_int(b))

def mod(a, b):
    return a % b

def round_half(x, n=0):
    n = to_int(n)
    if abs(n) > 15:
        raise ValueError
    q = Decimal(1).scaleb(-n)
    return float(Decimal(repr(float(x))).quantize(q, rounding=ROUND_HALF_UP))

def gst(amount, rate):
    return amount * (1 + rate / 100)

def gstr(total, rate):
    if rate <= -100:
        raise ValueError
    return total / (1 + rate / 100)

def disc(amount, pct):
    return amount * (1 - pct / 100)

def emi(principal, rate, months):
    """Monthly EMI for a loan. rate is yearly %, months is tenure."""
    n = to_int(months)
    if principal < 0 or rate < 0 or n <= 0 or n > 1200:
        raise ValueError
    r = rate / 1200
    if r == 0:
        return principal / n
    f = (1 + r) ** n
    return principal * r * f / (f - 1)

def chg(old, new):
    if old == 0:
        raise ValueError
    return (new - old) / abs(old) * 100

"""Evaluate an AST node using only whitelisted operators/functions.
Safe replacement for eval(); raises ValueError on anything else."""

def safe_eval(node, consts , funcs):
    ev = lambda n: safe_eval(n, consts, funcs)
    if isinstance(node, ast.Expression):
        return ev(node.body)
    if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)):
        return node.value
    if isinstance(node, ast.BinOp) and type(node.op) in OPS:
        a, b = ev(node.left), ev(node.right)
        if isinstance(node.op, ast.Pow) and abs(b) > 1000:
            raise ValueError
        return OPS[type(node.op)] (a,b)
    if isinstance(node, ast.UnaryOp) and isinstance(node.op , (ast.UAdd,ast.USub)):
        v = ev(node.operand)
        return v if isinstance(node.op, ast.UAdd) else -v
    if isinstance(node, ast.Name) and node.id in consts:
        return consts[node.id]
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in funcs and 1 <= len(node.args) <= 3 and not node.keywords:
        return funcs[node.func.id](*[ev(a) for a in node.args])
    raise ValueError

def to_round(s):
    return s.translate(str.maketrans("[{]}", "(())"))

def prepare(s):
    s = to_round(s)
    s = re.sub(r"(?<![\d.])0+(?=\d)", "", s)
    s = (s.replace("×", "*").replace("÷", "/").replace("−", "-").replace("^", "**").replace("√", "sqrt").replace("π", "pi").replace("Ans", "ANS").replace("Mem", "MEM"))
    s = re.sub(r"(?<![.A-Za-z])e(?![A-Za-z])", "E", s)

    m = re.fullmatch(r"(.*?[\d)A-Za-z])([+-])(\d+\.?\d*)%(.*)", s)
    if m and m.group(1).count("(") == m.group(1).count(")"):
        L, op, n, T = m.groups()
        s = f"(({L}){op}({L})*({n}/100)){T}"

    s = re.sub(r"(\d+\.?\d*)%", r"(\1/100)", s)
    s = re.sub(r"(\d+\.?\d*|\([^()]*\))!", r"fact(\1)", s)
    s = re.sub(r"(\d|\)|pi|E|ANS|MEM|X)(?=\(|pi|E|ANS|MEM|X|[a-z]+\()", r"\1*", s)
    s = re.sub(r"(\)|pi|E|ANS|MEM|X)(?=[\d.])", r"\1*", s)
    s += ")" * max(0, s.count("(") - s.count(")"))
    return s

def fmt(v):
    if v == int(v) and abs(v) < 1e15:
        return str(int(v))
    v = float(f"{v:.12g}")
    if v == 0:
        return "0"
    if abs(v) >= 1e15 or abs(v) < 1e-9:
        m, ex = f"{v:.11e}".split("e")
        return f"{m.rstrip('0').rstrip('.')} x10^{int(ex)}"
    if v == int(v):
        return str(int(v))
    return format(Decimal(f"{v:.12g}"), "f")

def to_expr(txt):
    return txt.replace(" x10^", "×10^")

def to_base(n, b):
    if n == 0:
        return "0"
    sign = "-" if n < 0 else ""
    n = abs(n)
    out = ""
    while n:
        out = DIG[n % b] + out
        n //= b
    return sign + out

def nice_step(span):
    raw = span / 7
    mag = 10 ** math.floor(math.log10(raw))
    for m in (1, 2, 5, 10):
        if m * mag >= raw:
            return m * mag

def clip_y(x1, y1, x2, y2, lo, hi):
    if (y1 < lo and y2 < lo) or (y1 > hi and y2 > hi):
        return None
    ox1, oy1, ox2, oy2 = x1, y1, x2, y2
    at = lambda y: ox1 + (ox2 - ox1) * (y - oy1) / (oy2 - oy1)
    if y1 < lo:   x1, y1 = at(lo), lo
    elif y1 > hi: x1, y1 = at(hi), hi
    if y2 < lo:   x2, y2 = at(lo), lo
    elif y2 > hi: x2, y2 = at(hi), hi
    return x1, y1, x2, y2

#Graph Function Logic
GFUNCS = "asin|acos|atan|sinh|cosh|tanh|sin|cos|tan|sec|csc|cot|logb|log|ln|sqrt|abs"

def gprep(s):
    s = re.sub(r"(?<![a-z])(" + GFUNCS + r")(\d+\.?\d*[xπ]?|[xπ]|pi)", r"\1(\2)", s)
    s = re.sub(r"(?<![A-Za-z])x(?=[a-z]{2,}\()", "X", s)
    s = re.sub(r"(?<![A-Za-z])x(?![A-Za-z])", "X", s)
    return prepare(s)

DCHARTS = [("bar", "Bar"), ("stack", "Stacked Bar"), ("hbar", "Horizontal Bar"), ("line", "Line"),
           ("area", "Area"), ("scatter", "Scatter"),("pie", "Pie"), ("donut", "Donut"), ("hist", "Histogram"), ("radar", "Radar")]

#File Reading (csv and xlsx)
def to_num(s):
    s = str(s).strip()
    if re.fullmatch(r"[-+]?\d+,\d{1,2}", s):
        s = s.replace(",", ".")
    s = re.sub(r"[₹$€£%\s,]", "", s)
    try:
        v = float(s)
    except ValueError:
        return None
    return v if math.isfinite(v) else None

def text_to_rows(text):
    lines = [l for l in text.splitlines() if l.strip()]
    if not lines:
        raise ValueError
    head = lines[0]
    delim = max(",\t;|", key=head.count)
    if head.count(delim) == 0:
        if len(head.split()) > 1:
            return [l.split() for l in lines]
        delim = ","
    return list(csv.reader(lines, delimiter=delim))

def read_xlsx(path):
    M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(M + "si"):
                shared.append("".join(t.text or "" for t in si.iter(M + "t")))
        sheet = "xl/worksheets/sheet1.xml"
        if sheet not in names:
            cand = sorted(n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))
            if not cand:
                raise ValueError
            sheet = cand[0]
        root = ET.fromstring(z.read(sheet))
    rows = []
    for r in root.iter(M + "row"):
        row = {}
        for pos, c in enumerate(r.findall(M + "c")):
            m = re.match(r"[A-Z]+", c.get("r", ""))
            col = pos
            if m:
                col = 0
                for ch in m.group():
                    col = col * 26 + ord(ch) - 64
                col -= 1
            t, v = c.get("t"), c.find(M + "v")
            if t == "s" and v is not None:
                val = shared[int(v.text)]
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in c.iter(M + "t"))
            else:
                val = v.text if v is not None and v.text else ""
            row[col] = val
        if row:
            rows.append([row.get(i, "") for i in range(max(row) + 1)])
    return rows

def read_table(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return read_xlsx(path)
    if ext == ".xls":
        raise ValueError("xls")
    with open(path, "rb") as f:
        raw = f.read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("latin-1")
    return text_to_rows(text)

def table_to_data(rows):
    rows = [[str(c).strip() for c in r] for r in rows]
    rows = [r for r in rows if any(r)]
    if not rows:
        raise ValueError
    ncol = max(len(r) for r in rows)
    rows = [r + [""] * (ncol - len(r)) for r in rows]
    nums = [[to_num(c) if c != "" else None for c in r] for r in rows]
    header = None
    if len(rows) > 1 and all(v is None for v in nums[0]):
        header, rows, nums = rows[0], rows[1:], nums[1:]
    numcols = []
    for j in range(ncol):
        filled = sum(1 for r in rows if r[j] != "")
        good = sum(1 for r in nums if r[j] is not None)
        if filled and good >= 0.6 * filled:
            numcols.append(j)
    if not numcols:
        raise ValueError
    labcol = next((j for j in range(ncol) if j not in numcols), None)
    if labcol is None and len(numcols) > 1:
        labcol = numcols.pop(0)
    numcols = numcols[:6]
    rows, nums = rows[:200], nums[:200]
    labs = [(r[labcol] if labcol is not None and r[labcol] else str(i + 1)) for i, r in enumerate(rows)]
    series = []
    for j in numcols:
        nm = header[j] if header and header[j] else f"Col {j + 1}"
        series.append((nm, [r[j] if r[j] is not None else 0.0 for r in nums]))
    return labs, series

class ScaledCanvas(tk.Canvas):

    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.S = 1.0
        self._fonts = {}
        self.cmap = {}

    def _n(self, args):
        out = []
        for a in args:
            if isinstance(a, (list, tuple)):
                out.extend(self._n(a))
            else:
                out.append(a * self.S)
        return out

    def _font(self, f):
        S = self.S
        if isinstance(f, tkfont.Font):
            a = f.actual()
            key = (a["family"], a["size"], a["weight"], S)
            if key not in self._fonts:
                self._fonts[key] = tkfont.Font(family=a["family"], size=max(5, round(abs(a["size"]) * S)),
                                               weight=a["weight"])
            return self._fonts[key]
        if isinstance(f, (tuple, list)) and len(f) >= 2:
            return (f[0], max(5, round(f[1] * S)), *f[2:])
        return f

    def _kw(self, kw):
        for c in ("fill", "outline"):
            if c in kw and kw[c] in self.cmap:
                kw[c] = self.cmap[kw[c]]
        if "font" in kw:
            kw["font"] = self._font(kw["font"])
        if "width" in kw:
            kw["width"] = max(1, round(kw["width"] * self.S))
        if "dash" in kw:
            kw["dash"] = tuple(max(1, round(d * self.S)) for d in kw["dash"])
        return kw

    def create_text(self, *a, **kw):      return super().create_text(*self._n(a), **self._kw(kw))
    def create_polygon(self, *a, **kw):   return super().create_polygon(*self._n(a), **self._kw(kw))
    def create_line(self, *a, **kw):      return super().create_line(*self._n(a), **self._kw(kw))
    def create_oval(self, *a, **kw):      return super().create_oval(*self._n(a), **self._kw(kw))
    def create_arc(self, *a, **kw):       return super().create_arc(*self._n(a), **self._kw(kw))
    def create_rectangle(self, *a, **kw): return super().create_rectangle(*self._n(a), **self._kw(kw))
    def create_image(self, *a, **kw):     return super().create_image(*self._n(a), **kw)

    def coords(self, tag, *a):
        if a:
            return super().coords(tag, *self._n(a))
        return [c / self.S for c in super().coords(tag)]

    def move(self, tag, dx, dy):
        super().move(tag, dx * self.S, dy * self.S)

    def bbox(self, *tags):
        b = super().bbox(*tags)
        return tuple(v / self.S for v in b) if b else b

    def itemconfigure(self, tag, cnf=None, **kw):
        if cnf is None and not kw:
            return super().itemconfigure(tag)
        return super().itemconfigure(tag, cnf, **self._kw(kw))
    itemconfig = itemconfigure

    def bind(self, seq=None, func=None, add=None):
        def wrapped(e):
            e.x = e.x / self.S      
            e.y = e.y / self.S
            return func(e)
        return super().bind(seq, wrapped, add)

class Calculator:
    def __init__(self, root):
        self.root = root
        root.title("Advanced  Calculator")
        root.resizable(True, True)
        root.configure(bg=BG)
        root.overrideredirect(True)
        root.geometry(f"{W}x{H}+200+30")
        root.attributes("-topmost", True)
        self.drag = None
        self.S = 1.0                                  
        self.maxed, self.minimized, self.rs_job = False, False, None
        self.norm_geo = self.mini_geo = f"{W}x{H}+200+30"
        self.gr = None
        self.wide = False       
        self.TW = W                
        self.CW = W                
        self.hpw = 0              
        self.PX1 = W - 24          
        self.tabx = {}             
        self.tab_span = (70, 380)
        self.hpoff, self.hp_rows = 0, []
        self.theme_x = (0, 0)
        self.p3d = None
        self.cv = ScaledCanvas(root, width=W, height=H, bg=BG, highlightthickness=0)
        self.cv.place(relx=0.5, rely=0.5, anchor="center")  

        try:
            self.grip = tk.Label(root, text="◢", bg=BG, fg="#5a5a5e", font=(FONT, 11), bd=0, cursor="size_nw_se")
        except tk.TclError:   
            self.grip = tk.Label(root, text="◢", bg=BG, fg="#5a5a5e", font=(FONT, 11), bd=0, cursor="bottom_right_corner")
        self.grip.place(relx=1.0, rely=1.0, anchor="se")
        self.grip.bind("<ButtonPress-1>", self.grip_down)
        self.grip.bind("<B1-Motion>", self.grip_drag)
        self.expr, self.ans, self.deg, self.done = "", 0.0, True, False
        self.mem = 0.0
        self.lg, self.lg_base, self.lg_val = None, "", ""
        self.keys = []
        self.held = None
        self.imgs = {}
        self.tab = "home"
        self.anim = False
        self.bar_dy = 0
        self.last_done = None
        self.mode_set = False
        self.history, self.hview, self.hoff, self.hrows = [], False, 0, []

        #Graph Expre..
        self.gexpr, self.gdeg, self.gstate, self.gfn = "sin(x)", False, "empty", None
        self.gcx, self.gcy = 0.0, 0.0
        self.gsx = self.gsy = (self.PX1 - PX0) / 16
        self.gfresh = True
        self.gxdeg, self.gconst_txt = False, ""
        self.gpan, self.gbtns = None, []
        self.gscreen = "menu"
        self.dlabs, self.dseries, self.dsel, self.dfile, self.dmsg = [], [], 0, "", ""
        self.dtype, self.ddopen, self.dbtns = "bar", False, []
        self.cat = 0
        self.cscreen = "menu"
        self.chips, self.unit_groups = [], []
        self.cfrom, self.cto = CATS[0][2], CATS[0][3]
        self.cin = ""
        self.conv_out = ""
        self.rates = dict(RATES)
        self.live, self.rates_done, self.rate_src = None, False, "offline"
        self.f_th = tkfont.Font(family=FONT, size=12)
        self.f_hr = tkfont.Font(family=FONT, size=13, weight="bold")

        self.draw_body()
        self.build_keys()
        self.update_tabs()
        self.cv.itemconfig(self.t_mode, text=self.mode_text())
        self.cv.bind("<ButtonPress-1>", self.on_down)
        self.cv.bind("<Double-Button-1>", self.on_double)
        self.cv.bind("<ButtonRelease-1>", self.on_up)
        self.cv.bind("<B1-Motion>", self.on_drag)
        self.cv.bind("<Motion>", self.on_move)
        root.bind("<Key>", self.on_key)
        root.bind("<Configure>", self.on_resize)
        root.bind("<Map>", self.on_map)
        self.cv.focus_set()
        root.after(200, root.focus_force)
        root.after(100, self.show_in_taskbar)
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.cv.bind(ev, self.on_wheel)
        threading.Thread(target=self.fetch_rates, daemon=True).start()
        root.after(500, self.check_rates)
        self.refresh()

    def show_in_taskbar(self):
      try:
         import ctypes
         u = ctypes.windll.user32
         hwnd = u.GetParent(self.root.winfo_id())
         style = u.GetWindowLongW(hwnd, -20)
         style = (style & ~0x00000080) | 0x00040000
         u.SetWindowLongW(hwnd, -20, style)
         self.root.withdraw()
         self.root.after(10, self.root.deiconify)
      except Exception:
        pass

    def work_area(self):
        try:
            import ctypes
            from ctypes import wintypes
            r = wintypes.RECT()
            ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(r), 0)
            return r.left, r.top, r.right - r.left, r.bottom - r.top
        except Exception:
            return 0, 0, self.root.winfo_screenwidth(), self.root.winfo_screenheight()

    def toggle_max(self):
        self.drag = None
        if self.maxed:
            self.root.geometry(self.norm_geo)
            self.cv.itemconfig(self.t_max, text="□")
            self.grip.place(relx=1.0, rely=1.0, anchor="se")
        else:
            self.norm_geo = (f"{self.root.winfo_width()}x{self.root.winfo_height()}"
                             f"+{self.root.winfo_x()}+{self.root.winfo_y()}")
            x, y, w, h = self.work_area()
            self.root.geometry(f"{w}x{h}+{x}+{y}")
            self.cv.itemconfig(self.t_max, text="❐")
            self.grip.place_forget()
        self.maxed = not self.maxed

    def on_double(self, e):
        if e.y < 50 and e.x < self.TW - 128 and not (self.theme_x[0] <= e.x <= self.theme_x[1]):
            self.toggle_max()

    def minimize(self):
        self.mini_geo = self.root.geometry()
        self.root.update_idletasks()
        self.root.overrideredirect(False)     
        self.minimized = True
        self.root.iconify()

    def on_map(self, e):
        if e.widget is self.root and self.minimized and self.root.state() == "normal":
            self.minimized = False
            self.root.after(10, self.restore_frameless)

    def restore_frameless(self):
        self.root.overrideredirect(True)
        self.root.geometry(self.mini_geo)
        self.root.attributes("-topmost", True)
        self.root.after(50, self.show_in_taskbar)
        self.root.after(100, self.root.focus_force)

    def grip_down(self, e):
        self.gr = (e.x_root, e.y_root, self.root.winfo_width(), self.root.winfo_height())

    def grip_drag(self, e):
        if not self.gr:
            return
        x0, y0, w0, h0 = self.gr
        w = max(MINW, w0 + e.x_root - x0)
        h = max(MINH, h0 + e.y_root - y0)
        self.root.geometry(f"{w}x{h}")

    def on_resize(self, e):
        if e.widget is not self.root:
            return
        if self.rs_job:
            self.root.after_cancel(self.rs_job)
        self.rs_job = self.root.after(120, self.apply_size)      

    def has_panel(self):
        return self.wide and self.tab not in ("conv", "graph")

    def update_cw(self):
        if not self.wide:
            self.CW = W
        elif self.tab in ("conv", "graph"):
            self.CW = self.TW
        else:
            self.CW = int(self.TW - self.hpw)
        self.PX1 = self.CW - 24

    def apply_size(self):
        self.rs_job = None
        if self.anim:
            self.rs_job = self.root.after(100, self.apply_size)
            return
        w, h = self.root.winfo_width(), self.root.winfo_height()
        if w < 10 or h < 10:
            return
        wide = (w / h) >= WIDE_RATIO
        s = min(w / (W + HP_MIN), h / H) if wide else min(w / W, h / H)
        s = round(min(max(s, MIN_S), MAX_S), 2)
        tw = max(int(round(w / s)), W + HP_MIN) if wide else W
        if wide == self.wide and abs(s - self.S) < 0.02 and abs(tw - self.TW) < 8:
            return
        old_cw = self.CW
        if abs(s - self.S) >= 0.02:
            self.imgs.clear()
        self.wide, self.TW = wide, tw
        self.hpw = tw - min(max(tw * CW_FRAC, W), CW_MAX) if wide else 0
        self.S = self.cv.S = s
        self.cv.config(width=round(tw * s), height=round(H * s))
        self.update_cw()
        self.relayout(keep_view=(old_cw == self.CW))

    def relayout(self, keep_view=True):
        cv = self.cv
        if self.wide and self.hview:
            self.hview = False
        view = (self.gcx, self.gcy, self.gsx, self.gsy, self.gxdeg)
        cv.delete("all")
        self.keys, self.held, self.bar_dy = [], None, 0
        self.draw_body()
        self.cv.itemconfig(self.t_max, text="❐" if self.maxed else "□")
        self.apply_screen()
        self.update_tabs()
        if self.tab == "graph" and keep_view and self.gscreen == "trig":
            self.gcx, self.gcy, self.gsx, self.gsy, self.gxdeg = view
            self.graph_redraw()
        if self.lg:
            self.log_show()
        if self.hview:
            for t in (self.t_hist, self.t_expr, self.t_res):
                cv.itemconfig(t, state="hidden")
            cv.itemconfig(self.t_mode, text="✕ CLEAR HISTORY", fill=ERR)
            self.render_history()

    #  layout
    def sections(self):
        if self.tab == "more":
            return [(MORE, 264, 52, 23, True), (BASIC, 480, 60, 25, False)]
        if self.tab in PANELS:
            return [(PANELS[self.tab], 262, 48, 21, True), (BASIC, 480, 60, 25, False)]
        return [(MEMORY, 268, 0, 19, True), (BASIC, 352, 86, 34, False)]

    def find(self, val):
        return next((i for i, k in enumerate(self.keys) if k["base"][1] == val), None)

    def build_keys(self):
        self.cv.delete("key")
        self.cv.delete("cui")
        self.keys = []
        self.chips, self.unit_groups = [], []
        self.held = None
        CW = self.CW
        if self.tab == "conv":
            if self.cscreen == "menu":
                self.add_menu_keys()
            else:
                self.draw_unit_screen()
        elif self.tab == "graph":
            if self.gscreen == "menu":
                self.add_graph_menu_keys()
            elif self.gscreen == "trig":
                self.draw_graph_screen()
            else:
                self.draw_data_screen()
        else:
            for rows, y0, dy, r, panel in self.sections():
                for ri, row in enumerate(rows):
                    c = 0
                    for j, (label, style, val, *span) in enumerate(row):
                        n = span[0] if span else 1
                        if panel:
                            m = len(row)
                            cols = max(5, m)
                            slot = j * 2 if m == 3 else (cols - m) / 2 + j
                            pitch = (CW - 56) / cols
                            cx = 28 + pitch * (slot + 0.5)
                            half = max(0, (pitch - 8) / 2 - r) if self.wide else 0
                        else:
                            unit = (CW - 68) / 6          
                            cx = 34 + unit * (c + n / 2)
                            half = max(0, (unit * n - 8) / 2 - r)
                        self.keys.append({"label": label, "style": style, "val": val, "base": (label, val),"cx": cx, "cy": y0 + dy * ri, "half": half, "r": r, "panel": panel})
                        c += n
        self.panel_fs = min((self.fit_size(k["label"], 2 * (k["r"] + k["half"]) - 6, 12, 8) for k in self.keys if k["panel"]), default=12)
        self.rad_idx = self.find("SETRAD")
        self.deg_idx = self.find("SETDEG")
        for i in range(len(self.keys)):
            self.draw_key(i)
        self.update_mode_keys()

    def add_menu_keys(self):
        n = len(CATS)
        CW = self.CW
        bw = (CW - 48 - 16) / 2
        for i, c in enumerate(CATS):
            row, col = divmod(i, 2)
            last_full = (i == n - 1 and n % 2 == 1)
            cx = CW / 2 if last_full else (24 + bw / 2 if col == 0 else CW - 24 - bw / 2)
            w = (CW - 48) if last_full else bw
            self.keys.append({"label": c[0], "style": "num", "val": f"CAT{i}", "base": (c[0], f"CAT{i}"),"cx": cx, "cy": 160 + 100 * row, "half": 0, "r": 40, "panel": False,"flat": True, "w": w, "h": 84})

    def add_graph_menu_keys(self):
        CW = self.CW
        bw = (CW - 48 - 16) / 2
        for i, (name, val) in enumerate(GITEMS):
            cx = 24 + bw / 2 if i == 0 else CW - 24 - bw / 2
            self.keys.append({"label": name, "style": "num", "val": val, "base": (name, val), "cx": cx, "cy": 160, "half": 0, "r": 40, "panel": False, "flat": True, "w": bw, "h": 84})

    def update_tabs(self, name=None):
        name = "hist" if self.hview else (name or self.tab)
        for n, t in self.t_tabs.items():
            off = "#111111" if n in ("home", "hist") else TAB_OFF     
            self.cv.itemconfig(t, fill=TAB_ON if n == name else off)
        for n, p in self.t_pills.items():
            self.cv.itemconfig(p, state="normal" if n == name else "hidden")

    def groups(self, only_panel=False):
        g = [(f"k{i}", k["cx"], k["cy"]) for i, k in enumerate(self.keys) if (k["panel"] or not only_panel)]
        if not only_panel and self.tab == "conv" and self.cscreen == "unit":
            g += self.unit_groups
        if self.tab == "graph":
            g += self.unit_groups
        return g

    def slide(self, groups, dx, frames, out, done):
        if not groups:
            return done()
        cv = self.cv
        ys = sorted({round(cy) for _, _, cy in groups})
        rank = {y: i for i, y in enumerate(ys)}
        delay = {tag: min(rank[round(cy)], 4) for tag, _, cy in groups}
        last = frames + max(delay.values())
        prev = {tag: 0.0 for tag, _, _ in groups}
        if not out:
            for tag in prev:
                cv.move(tag, dx, 0)
                prev[tag] = dx

        def step(f):
            for tag in prev:
                t = min(max((f - delay[tag]) / frames, 0.0), 1.0)
                pos = dx * t * t if out else dx * (1 - t) ** 3
                if pos != prev[tag]:
                    cv.move(tag, pos - prev[tag], 0)
                    prev[tag] = pos
            if f < last:
                self.root.after(16, lambda: step(f + 1))
            else:
                done()
        step(1)

    def transition(self, direction, change, only_panel=False):
        self.anim = True

        def finish():
            self.anim = False

        def swap():
            change()
            self.raise_top()
            self.slide(self.groups(only_panel), direction * (self.TW + 60), ANIM_IN, False, finish)
        self.raise_top()
        self.slide(self.groups(only_panel), -direction * (self.TW + 60), ANIM_OUT, True, swap)

    def set_bar(self, on):
        target = -143 if on else 0
        if target != self.bar_dy:
            self.cv.move("tabbar", 0, target - self.bar_dy)
            self.bar_dy = target
        self.cv.itemconfig("lcd", state="hidden" if on else "normal")

    def apply_screen(self):
        self.set_bar(self.tab in ("conv", "graph"))
        self.build_keys()
        if self.tab == "conv":
            self.conv_show()
        elif self.tab != "graph":
            self.show_main()

    def set_tab(self, name):
        if self.lg or self.anim:
            return
        if self.hview:
            self.toggle_history()
        if name == self.tab:
            if name == "conv" and self.cscreen == "unit":
                self.go_menu()
            elif name == "graph" and self.gscreen != "menu":
                self.graph_menu()
            return
        self.update_tabs(name)

        def change():
            self.tab = name
            if name == "conv":
                self.cscreen = "menu"
            if name == "graph":
                self.gscreen = "menu"
            old = self.CW
            self.update_cw()
            if self.CW != old:
                self.relayout(keep_view=False)
            else:
                self.apply_screen()
        direction = 1 if TABS.index(name) > TABS.index(self.tab) else -1
        self.transition(direction, change, only_panel=True)

    def show_main(self):
        self.cv.itemconfig(self.t_mode, text=self.mode_text(), fill=INK2)
        if self.done and self.last_done:
            self.show_done(*self.last_done)
        else:
            self.refresh()

    def tab_click(self, e):
        if self.anim:
            return
        name = None
        for n, (x1, x2) in self.tabx.items():
            if x1 <= e.x <= x2:
                name = n
                break
        if name is None:
            return
        if name in ("home", "hist"):
            self.press3d(name)
        if name == "hist":
            if self.tab in ("conv", "graph"):
                self.set_tab("home")
                self.after_anim(self.toggle_history)
            else:
                self.press("HIST")
        else:
            self.set_tab(name)

    def after_anim(self, fn):
        if self.anim:
            self.root.after(50, lambda: self.after_anim(fn))
        else:
            fn()

    def mode_text(self):
        if self.tab == "graph":
            return "GRAPH"
        if self.tab == "conv":
            return "CONVERTER" if self.cscreen == "menu" else "CONVERTER · " + CATS[self.cat][0].upper()
        t = ""
        if self.tab == "trig" and (self.mode_set or not self.deg):
            t = "DEG" if self.deg else "RAD"
        if self.mem:
            t = (t + "  M").strip()
        return t

    def update_mode_keys(self):
        for idx, active in ((self.rad_idx, not self.deg), (self.deg_idx, self.deg)):
            if idx is None:
                continue
            self.keys[idx]["style"] = "top" if active else "fn"
            self.draw_key(idx)

    def show_done(self, expr_text, txt):
        cv = self.cv
        self.last_done = (expr_text, txt)
        cv.itemconfig(self.t_hist, text=self.fit(expr_text + " =", self.f_th, 320), font=(FONT, 12))
        cv.itemconfig(self.t_expr, text="")
        n = len(txt)
        size = 38 if n <= 10 else 28 if n <= 15 else 20
        cv.coords(self.t_res, self.CW - 44, 143)
        cv.itemconfig(self.t_res, text=txt, fill=GLOW, font=(FONT, size, "bold"))
        self.render_hpanel()

    def round_rect(self, x1,y1,x2,y2,r,**kw):
        pts = [ x1+r , y1, x2-r, y1, x2, y1, x2, y1+r, x2, y2-r, x2, y2, x2-r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y1+r, x1, y1 ]
        return self.cv.create_polygon(pts, smooth = True, **kw)

    def layout_tabs(self):
        L, R = 24, self.CW - 24
        self.tabx = {"home": (L, L + 40)}
        right = R
        if not self.wide:
            self.tabx["hist"] = (R - 70, R)
            right = R - 76
        x = L + 46
        span = right - x
        tot = sum(TAB_WTS)
        for n, wt in zip(TAB_NAMES, TAB_WTS):
            w = span * wt / tot
            self.tabx[n] = (x, x + w)
            x += w
        self.tab_span = (L + 46, right)

    def draw_theme_switch(self, x):
        cv = self.cv
        dark = THEME == "dark"
        w = 40
        self.theme_x = (x, x + w)
        self.round_rect(x, 23, x + w, 43, 10, fill="#3a3a3c" if dark else "#c7c7cc", outline="", tags="theme")
        kx = x + w - 10 if dark else x + 10
        cv.create_oval(kx - 8, 25, kx + 8, 41, fill="#fffffe", outline="", tags="theme")
        cv.create_text(kx, 33, text="☾" if dark else "☀", fill="#111111" if dark else ACCENT,font=(FONT, 9, "bold"), tags="theme")

    def toggle_theme(self):
        if self.anim:
            return
        set_theme("light" if THEME == "dark" else "dark")
        self.cv.cmap = THEMES[THEME]["CMAP"]
        self.imgs.clear()
        self.root.configure(bg=BG)
        self.cv.configure(bg=BG)
        self.grip.configure(bg=BG, fg=THEMES[THEME]["GRIP"])
        self.relayout()

    def press3d(self, name):
        if self.p3d:
            return
        self.p3d = name
        self.cv.move("t3d_" + name, 0, 2)

        def back():
            if self.p3d == name:
                self.cv.move("t3d_" + name, 0, -2)
                self.p3d = None
        self.root.after(130, back)

    def raise_top(self):
        self.cv.tag_raise("hmask")
        self.cv.tag_raise("tabbar")
        self.cv.tag_raise("hp")

    def draw_body(self):
        cv = self.cv
        TW, CW = self.TW, self.CW
        self.p3d = None
        self.round_rect(6, 6, TW - 6, H - 6, 38, fill=BODY_EDGE, outline="")
        self.round_rect(10, 10, TW - 10, H - 10, 35, fill=BODY, outline="")
        if self.has_panel():
            cv.create_rectangle(CW - 8, 54, TW - 10, H - 50, fill=BODY, outline="", tags="hmask")

        t_title = cv.create_text(34, 33, text="ADVANCED CALCULATOR", anchor="w",fill="#9aa1ad", font=(FONT, 9, "bold"))
        self.draw_theme_switch(cv.bbox(t_title)[2] + 20)
        cv.create_text(TW - 36, 33, text="✕", fill="#9aa1ad", font=(FONT, 12, "bold"))
        self.t_max = cv.create_text(TW - 72, 33, text="□", fill="#9aa1ad", font=(FONT, 12, "bold"))
        cv.create_text(TW - 108, 33, text="─", fill="#9aa1ad", font=(FONT, 12, "bold"))
        self.round_rect(24, 54, CW - 24, 190, 16, fill="#15171b", outline="", tags="lcd")
        self.round_rect(30, 60, CW - 30, 184, 10, fill=LCD, outline="", tags="lcd")
        self.t_mode = cv.create_text(44, 72, text="", anchor="w",fill=INK2, font=(FONT, 9, "bold"), tags="lcd")
        self.t_hist = cv.create_text(CW - 44, 96, text="", anchor="e",fill=INK2, font=(FONT, 12), tags="lcd")
        self.t_expr = cv.create_text(CW - 44, 124, text="0", anchor="e",fill=INK, font=(FONT, 26), tags="lcd")
        self.t_res = cv.create_text(CW - 44, 158, text="", anchor="e",fill=INK, font=(FONT, 28, "bold"), tags="lcd")
        self.layout_tabs()
        c0, c1 = self.tab_span
        self.round_rect(c0, 197, c1, 231, 14, fill="#151518", outline="#3a3a3c", width=1, tags="tabbar")
        for nm in ("home", "hist"):                 
            if nm in self.tabx:
                x1, x2 = self.tabx[nm]
                tg = ("tabbar", "t3d_" + nm)
                self.round_rect(x1, 200, x2, 231, 12, fill="#8e8e93", outline="", tags=tg)
                self.round_rect(x1, 197, x2, 228, 12, fill="#fffffe", outline="#d1d1d6", width=1, tags=tg)
        self.t_pills = {}
        for name in TAB_NAMES:
            x1, x2 = self.tabx[name]
            self.t_pills[name] = self.round_rect(x1 + 3, 200, x2 - 3, 228, 11, fill="#2c2c2e", outline="", state="hidden", tags="tabbar")

        start = 13 if self.wide else 12
        tsize = min(self.fit_size(TAB_LABELS[n], self.tabx[n][1] - self.tabx[n][0] - 14, start, 7) for n in TAB_NAMES)
        
        if "hist" in self.tabx:
            tsize = min(tsize, self.fit_size("History", 70 - 14, start, 7))
        self.t_tabs = {}
        for name in TABS:
            if name not in self.tabx:
                continue
            x1, x2 = self.tabx[name]
            txt = "⌂" if name == "home" else "History" if name == "hist" else TAB_LABELS[name]
            b3d = name in ("home", "hist")
            self.t_tabs[name] = cv.create_text((x1 + x2) / 2, TABY - 2 if b3d else TABY, text=txt, fill=TAB_OFF,
                                               font=(FONT, 22 if name == "home" else tsize, "bold"),
                                               tags=("tabbar", "t3d_" + name) if b3d else "tabbar")
        self.render_hpanel()

    def key_img(self, style, half, R, pressed=False):
        S = self.S
        half = round(half, 1)
        ck = (style, half, R, pressed, S)
        if ck in self.imgs:
            return self.imgs[ck]

        rgb = lambda c: [int(c[i:i + 2], 16) for i in (1, 3, 5)]
        lerp = lambda a, b, t: [x + (y - x) * t for x, y in zip(a, b)]
        clamp = lambda v: min(max(v, 0.0), 1.0)
        hexc = lambda c: "#%02x%02x%02x" % tuple(max(0, min(255, round(v))) for v in c)
        WHITE, BLACK = [255, 255, 255], [0, 0, 0]
        face = rgb(STYLES[style][0])
        body = rgb(BODY)
        light = THEME == "light"
        if light:
            top_c = lerp(face, WHITE, 0.55)
            bot_c = lerp(face, BLACK, 0.07)
            side_c = lerp(face, BLACK, 0.20)
        else:
            top_c = lerp(face, WHITE, 0.22)
            bot_c = lerp(face, BLACK, 0.30)
            side_c = lerp(face, BLACK, 0.55)
        gloss = 0.35 if light else 0.12

        Rs, hs, D, pad = R * S, half * S, DEPTH * S, 2 * S     
        w, h = int(2 * (Rs + hs) + 2 * pad) + 1, int(2 * Rs + 2 * pad + D) + 1
        cx = w / 2
        sy = Rs + pad + D
        fy = Rs + pad + (D - S if pressed else 0)
        img = tk.PhotoImage(width=w, height=h)
        for py in range(h):
            row = []
            for px in range(w):
                x, y = px + 0.5, py + 0.5
                nx = min(max(x, cx - hs), cx + hs)
                c = body
                cov_s = clamp(Rs - math.hypot(x - nx, y - sy) + 0.5)
                if cov_s > 0:
                    c = lerp(c, side_c, cov_s)
                    dist = math.hypot(x - nx, y - fy)
                    cov_f = clamp(Rs - dist + 0.5)
                    if cov_f > 0:
                        t = clamp((y - (fy - Rs)) / (2 * Rs))
                        col = lerp(top_c, bot_c, t)
                        gl = clamp(1 - math.hypot((x - cx) / ((Rs + hs) * 0.8),(y - (fy - Rs * 0.45)) / (Rs * 0.5)))
                        col = lerp(col, WHITE, gl * gloss)
                        if pressed:
                            col = lerp(col, BLACK, 0.08 if light else 0.15)
                        c = lerp(c, col, cov_f)
                row.append(hexc(c))
            img.put("{" + " ".join(row) + "}", to=(0, py))
        self.imgs[ck] = img
        return img

    def fit_size(self, text, avail, start=13, low=7):
        cache = self.__dict__.setdefault("_fs", {})
        size = start
        while size > low:
            f = cache.get(size)
            if f is None:
                f = cache[size] = tkfont.Font(family=FONT, size=size)
            if f.measure(text) <= avail:
                break
            size -= 1
        return size

    def draw_key(self, i, pressed=False):
        k = self.keys[i]
        tag = f"k{i}"
        r, half, cx, cy = k["r"], k["half"], k["cx"], k["cy"]
        self.cv.delete(tag)
        ink = STYLES[k["style"]][1]

        if k.get("flat"):
            w, h = k["w"], k["h"]
            self.round_rect(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2, 16,fill="#3a3a3c" if pressed else "#1c1c1e", outline="#3a3a3c", width=1, tags=(tag, "key"))
            lx = cx - w / 2 + 34
            self.cv.create_text(lx, cy, text=ICONS.get(k["label"], "•"), fill=ACCENT,font=("Segoe UI Emoji", 20), tags=(tag, "key"))
            avail = w - 62 - 10
            self.cv.create_text(cx - w / 2 + 62, cy, text=k["label"], anchor="w", fill="#ffffff",font=(FONT, self.fit_size(k["label"], avail, 12, 8), "bold"), tags=(tag, "key"))
            return

        img = self.key_img(k["style"], half, r, pressed)
        self.cv.create_image(cx - img.width() / (2 * self.S), cy - (r + 2), image=img, anchor="nw", tags=(tag, "key"))
        ty = cy + (DEPTH - 1 if pressed else 0)

        if k["panel"]:
            size = self.panel_fs
        elif k["style"] == "num":
            size = int(r * 0.72)
        elif k["style"] == "op":
            size = int(r * 0.95) + (2 if k["val"] == "=" else 0)
        else:
            size = int(r * 0.6)
        self.cv.create_text(cx, ty, text=k["label"], fill=ink, font=(FONT, size), tags=(tag, "key"))

    def hit(self, x, y):
        for i, k in enumerate(self.keys):
            if k.get("flat"):
                if abs(x - k["cx"]) <= k["w"] / 2 and abs(y - k["cy"]) <= k["h"] / 2:
                    return i
                continue
            nx = min(max(x, k["cx"] - k["half"]), k["cx"] + k["half"])
            if (x - nx)**2 + (y - k["cy"]) ** 2 <= k["r"] ** 2:
                return i
        return None

    def on_drag(self, e):
        if self.gpan and self.tab == "graph":
            x0, y0, cx0, cy0 = self.gpan
            self.gcx = cx0 - (e.x - x0) / self.gsx
            self.gcy = cy0 + (e.y - y0) / self.gsy
            return self.graph_redraw()
        if self.drag and not self.maxed:
            self.root.geometry(f"+{e.x_root - self.drag[0]}+{e.y_root - self.drag[1]}")

    def on_down(self, e):
        self.root.focus_force()
        if e.y < 50:
            if e.x > self.TW - 56:
                return self.root.destroy()
            if e.x > self.TW - 92:
                return self.toggle_max()
            if e.x > self.TW - 128:
                return self.minimize()
            if self.theme_x[0] <= e.x <= self.theme_x[1] and 21 <= e.y <= 45:
                return self.toggle_theme()
            if self.maxed:
                return
            self.drag = (e.x_root - self.root.winfo_x(), e.y_root - self.root.winfo_y())
            return
        self.drag = None
        if self.anim:
            return
        if self.has_panel() and e.x > self.CW - 8:
            return self.hpanel_click(e)
        if 196 + self.bar_dy <= e.y <= 234 + self.bar_dy:
            return self.tab_click(e)
        if self.tab == "conv" and self.cscreen == "unit" and e.y > 100:
            if self.conv_click(e):
                return
        if self.tab == "graph" and e.y > 100:
            if self.graph_click(e):
                return
        i = self.hit(e.x, e.y)
        if self.hview and i is None and e.y < 190:
            return self.history_click(e)
        if i is not None:
            self.held = i
            self.draw_key(i, True)

    def on_up(self, e):
        self.drag = None
        self.gpan = None
        if self.held is not None:
            i = self.held
            self.held = None
            if self.hit(e.x, e.y) == i:
                self.press(self.keys[i]["val"])
                if not self.anim:
                    self.draw_key(i, True)
                    self.root.after(140, lambda i=i: self.draw_key(i) if (not self.anim and i < len(self.keys)) else None)
            else:
                self.draw_key(i)

    def on_key(self, e):
        ch, ks = e.char, e.keysym
        if self.tab == "graph":
            return None if self.anim else self.graph_key(e)
        if e.state & 0x4 and ks in ("c", "C"):
            return self.copy_result()
        if e.state & 0x4 and ks in ("v", "V"):
            return self.paste()
        if ch in ("h", "H") and not re.search(r"[A-Za-z]$", self.expr):
            return self.toggle_history()
        if self.hview and ks in ("Up", "Down"):
            return self.scroll_history(1 if ks == "Up" else -1)
        if self.hview and ks == "Escape":
            return self.toggle_history()
        if self.tab != "conv":
            if ch == "i" and self.expr.endswith("π"):
                return
            run = re.search(r"[A-Za-z]+$", self.expr)
            if ch == "p" and run and run.group() != "e":
                return self.press("p")
        m = {"*": "×", "/": "÷", "-": "−", "+": "+", "^": "^", "(": "(", ")": ")","[": "[", "]": "]", "{": "{", "}": "}",".": ".", 
                 "%": "%", "!": "!", "p": "π", "e": "e", ",": ","}
        if ch.isdigit():
            v = ch
        elif ch in m:
            v = m[ch]
        elif ks in ("Return", "KP_Enter") or ch == "=":
            v = "="
        elif ks == "BackSpace":
            v = "DEL"
        elif ks == "F9":
            v = "NEG"
        elif ks == "Escape":
            v = "AC"
        elif ks == "Tab" or ks in ARROWS:
            if self.tab != "conv":
                return
            v = "SWAP" if ks == "Tab" else ARROWS[ks]
        elif ch and ch.isprintable() and not ch.isspace():
            v = ch
        else:
            return
        for i, k in enumerate(self.keys):
            if k["val"] == v:
                self.draw_key(i, True)
                self.root.after(110, lambda i=i: self.draw_key(i) if (not self.anim and i < len(self.keys)) else None)
                break
        self.press(v)

    def copy_result(self):
        if self.tab == "conv":
            txt = self.conv_out
        else:
            txt = f"{self.ans:.12g}" if self.done else self.expr
        self.root.clipboard_clear()
        self.root.clipboard_append(txt)

    def paste(self):
        try:
            txt = self.root.clipboard_get()
        except tk.TclError:
            return
        txt = re.sub(r"[₹$€£\s]", "", txt)
        if re.fullmatch(r"[-+]?[\d,]*\.?\d+", txt):
            txt = txt.replace(",", "")
        txt = txt.replace("*", "×").replace("/", "÷").replace("-", "−")[:60]
        ok = "0123456789.+−×÷()%^"
        for ch in txt:
            if ch in ok or (self.tab == "conv" and ch.upper() in DIG):
                self.press(ch)

    def is_num(self):
        return CATS[self.cat][0] == "Number System"

    def is_cur(self):
        return CATS[self.cat][0] == "Currency"

    def is_temp(self):
        return CATS[self.cat][0] == "Temperature"

    def cat_units(self):
        if self.is_cur():
            return [(n, c, 1 / self.rates[c]) for n, c in CURR]
        return CATS[self.cat][1]

    def valid_in_base(self, s, b):
        return all(ch in DIG[:b] for ch in s.lstrip("-").upper())

    def fix_input(self):
        if self.is_num() and not self.valid_in_base(self.cin, self.cat_units()[self.cfrom][2]):
            self.cin = ""

# Gets live exchange rates from The API (I used API from https://www.exchangerate-api.com/).

    def fetch_rates(self):
        try:
            with urllib.request.urlopen("https://open.er-api.com/v6/latest/USD", timeout=6) as r:
                data = json.loads(r.read().decode("utf-8"))
            rates = data.get("rates", {})
            got = {c: float(rates[c]) for _, c in CURR if c in rates}
            if "INR" in got and "USD" in got:
                self.live = got
        except Exception:
            pass
        self.rates_done = True

    def check_rates(self, tries=0):
        if self.rates_done:
            if self.live:
                self.rates.update(self.live)
                self.rate_src = "live"
                self.update_note()
                if self.tab == "conv":
                    self.conv_show()
        elif tries < 24:
            self.root.after(500, lambda: self.check_rates(tries + 1))

    def note_text(self):
        if self.is_cur():
            return "Live exchange rates" if self.rate_src == "live" else "Offline approx. rates (internet needed for live)"
        return ""

    def update_note(self):
        self.cv.itemconfig("cnote", text=self.note_text())

    def draw_unit_screen(self):
        cv = self.cv
        CW = self.CW
        cv.delete("cui")
        self.chips = []
        g = self.unit_groups = []
        units = self.cat_units()
        n = len(units)
        fu, tu = units[self.cfrom], units[self.cto]
        self.round_rect(24, 110, 112, 140, 12, fill="#151517", outline="#3a3a3c", width=1, tags=("cui", "cback"))
        cv.create_text(68, 125, text="‹  Back", fill="#ffffff", font=(FONT, 11, "bold"), tags=("cui", "cback"))
        cv.create_text(CW - 28, 125, text=CATS[self.cat][0], anchor="e", fill="#ffffff", font=(FONT, 13, "bold"), tags=("cui", "ctitle"))
        g += [("cback", 68, 125), ("ctitle", CW - 80, 125)]

        t1 = ("cui", "cbox1")
        self.round_rect(24, 156, CW - 24, 236, 16, fill="#15171b", outline=ACCENT, width=2, tags=t1)
        cv.create_text(40, 174, text=f"FROM  ·  {fu[0]} ({fu[1]})", anchor="w", fill=TAB_OFF, font=(FONT, 9, "bold"), tags=t1)
        self.t_cin = cv.create_text(CW - 40, 210, text="", anchor="e", fill="#ffffff", font=(FONT, 26), tags=t1)

        ts = ("cui", "cswap")
        cv.create_oval(CW / 2 - 18, 240, CW / 2 + 18, 276, fill=ACCENT, outline="", tags=ts)
        cv.create_text(CW / 2, 258, text="↕", fill="#fffffe", font=(FONT, 16, "bold"), tags=ts)

        t2 = ("cui", "cbox2")
        self.round_rect(24, 280, CW - 24, 360, 16, fill=LCD, outline="", tags=t2)
        cv.create_text(40, 298, text=f"TO  ·  {tu[0]} ({tu[1]})", anchor="w", fill=INK2, font=(FONT, 9, "bold"), tags=t2)
        self.t_cout = cv.create_text(CW - 40, 334, text="", anchor="e", fill=INK, font=(FONT, 26, "bold"), tags=t2)
        g += [("cbox1", CW / 2, 196), ("cswap", CW / 2, 258), ("cbox2", CW / 2, 320)]
        cv.create_text(CW / 2, 378, text=self.note_text(), fill=TAB_OFF, font=(FONT, 8), tags=("cui", "cnote"))
        g.append(("cnote", CW / 2, 378))
        cols = max(6, min(10, int((CW - 48) // 80))) if self.wide else 6
        pitch = (CW - 42) / cols

        def block(kind, label, y0):
            cv.create_text(28, y0, text=label, anchor="w", fill=TAB_OFF, font=(FONT, 8, "bold"), tags=("cui", "clab" + kind))
            g.append(("clab" + kind, 40, y0))
            cur = self.cfrom if kind == "f" else self.cto
            for idx, (nm, sym, _) in enumerate(units):
                r, c = divmod(idx, cols)
                x1 = 24 + c * pitch
                x2 = x1 + pitch - 6
                cy = y0 + 24 + r * 34
                sel = idx == cur
                tag = f"ch{kind}{idx}"
                tg = ("cui", tag)
                self.round_rect(x1, cy - 14, x2, cy + 14, 11, fill="#2c2c2e" if sel else "#151517",
                                outline="" if sel else "#3a3a3c", width=1, tags=tg)
                cv.create_text((x1 + x2) / 2, cy, text=sym, fill=TAB_ON if sel else TAB_OFF, font=(FONT, 10, "bold"), tags=tg)
                self.chips.append((x1, cy - 14, x2, cy + 14, kind, idx))
                g.append((tag, (x1 + x2) / 2, cy))
            return y0 + 24 + ((n - 1) // cols + 1) * 34

        y = block("f", "CHANGE FROM", 400)
        block("t", "CHANGE TO", y + 2)
        help_lines = ("Type numbers with your keyboard",
                      "Here All Convert the units,used by API",
                      "Here Only Type Integer and Float Number",
                      "Not Used String or Character")
        for i, t in enumerate(help_lines):
            cv.create_text(CW / 2, 656 + 17 * i, text=t, fill=TAB_OFF, font=(FONT, 9), tags=("cui", "chelp"))
        g.append(("chelp", CW / 2, 673))
        self.conv_show()

    def open_cat(self, i):
        def change():
            self.cat = i
            self.cfrom, self.cto = CATS[i][2], CATS[i][3]
            self.cin = ""
            self.cscreen = "unit"
            self.apply_screen()
        self.transition(1, change)

    def go_menu(self):
        def change():
            self.cscreen = "menu"
            self.apply_screen()
        self.transition(-1, change)

    def open_graph(self, i):
        def change():
            self.gscreen = "trig" if i == 0 else "data"
            self.apply_screen()
        self.transition(1, change)

    def graph_menu(self):
        def change():
            self.gscreen = "menu"
            self.apply_screen()
        self.transition(-1, change)

    def swap_units(self):
        self.cfrom, self.cto = self.cto, self.cfrom
        self.fix_input()
        self.draw_unit_screen()

    def conv_click(self, e):
        if 24 <= e.x <= 112 and 110 <= e.y <= 140:
            self.go_menu()
            return True
        if math.hypot(e.x - self.CW / 2, e.y - 258) <= 20:
            self.swap_units()
            return True
        for x1, y1, x2, y2, kind, idx in self.chips:
            if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                if kind == "f":
                    self.cfrom = idx
                else:
                    self.cto = idx
                self.fix_input()
                self.draw_unit_screen()
                return True
        return False

    def conv_press(self, v):
        if self.cscreen == "menu":
            if v.startswith("CAT"):
                self.open_cat(int(v[3:]))
            elif len(v) == 1 and v in "123456789" and int(v) <= len(CATS):
                self.open_cat(int(v) - 1)
            return
        n = len(self.cat_units())
        if v == "AC":
            if self.cin:
                self.cin = ""
            else:
                return self.go_menu()
        elif v == "DEL":
            self.cin = self.cin[:-1]
        elif v in ("NEG", "−"):
            self.cin = self.cin[1:] if self.cin.startswith("-") else "-" + self.cin
        elif v == "SWAP":
            return self.swap_units()
        elif v in ("FPREV", "FNEXT", "TPREV", "TNEXT"):
            step = -1 if v.endswith("PREV") else 1
            if v[0] == "F":
                self.cfrom = (self.cfrom + step) % n
            else:
                self.cto = (self.cto + step) % n
            self.fix_input()
            return self.draw_unit_screen()
        elif v == ".":
            if self.is_num() or "." in self.cin:
                return
            self.cin += "0." if self.cin in ("", "-") else "."
        elif len(v) == 1:
            d = v.upper()
            if self.is_num():
                if d not in DIG[:self.cat_units()[self.cfrom][2]]:
                    return
            elif d not in "0123456789":
                return
            if len(self.cin) >= 20:
                return
            if self.cin in ("0", "-0"):
                self.cin = self.cin[:-1]
            self.cin += d
        else:
            return
        self.conv_show()

    def conv_result(self):
        units = self.cat_units()
        fu, tu = units[self.cfrom], units[self.cto]
        s = self.cin
        if s in ("", "-", ".", "-."):
            s = "0"
        try:
            if self.is_num():
                return to_base(int(s, fu[2]), tu[2])
            v = float(s)
            if self.is_temp():
                c = v if fu[1] == "°C" else (v - 32) * 5 / 9 if fu[1] == "°F" else v - 273.15
                out = c if tu[1] == "°C" else c * 9 / 5 + 32 if tu[1] == "°F" else c + 273.15
            else:
                out = v * fu[2] / tu[2]
            if self.is_cur():
                t = f"{out:.4f}".rstrip("0").rstrip(".")
                return "0" if t in ("", "-0") else t
            return fmt(out)
        except Exception:
            return "Error"

    def conv_show(self):
        if self.tab != "conv" or self.cscreen != "unit" or not self.chips:
            return
        cv = self.cv
        res = self.conv_result()
        self.conv_out = res
        inp = self.cin.replace("-", "−") or "0"
        if len(inp) > 30:
            inp = "…" + inp[-29:]
        n = len(inp)
        cv.itemconfig(self.t_cin, text=inp, font=(FONT, 26 if n <= 14 else 20 if n <= 19 else 15))
        cv.delete("ccur")
        b = cv.bbox(self.t_cin)
        if b:
            h = b[3] - b[1]
            x = b[2] + 4
            cv.create_line(x, b[1] + h * 0.18, x, b[3] - h * 0.22, fill="#ffffff", width=2,
                           tags=("cui", "cbox1", "ccur"))
        out = res if len(res) <= 34 else res[:33] + "…"
        n = len(out)
        cv.itemconfig(self.t_cout, text=out, fill=ERR if res == "Error" else INK,font=(FONT, 26 if n <= 14 else 20 if n <= 19 else 15 if n <= 26 else 11, "bold"))

    def make_env(self, deg=None):
        d = math.pi / 180 if (self.deg if deg is None else deg) else 1
        clean = lambda x: 0.0 if abs(x) < 1e-12 else x
        def tan(v):
            if abs(math.cos(v * d)) < 1e-12:
                raise ValueError("tan undefined")
            return clean(math.tan(v * d))
        def sec(v):
            c = math.cos(v * d)
            if abs(c) < 1e-12:
                raise ValueError
            return 1 / c
        def csc(v):
            s_ = math.sin(v * d)
            if abs(s_) < 1e-12:
                raise ValueError
            return 1 / s_
        def cot(v):
            s_ = math.sin(v * d)
            if abs(s_) < 1e-12:
                raise ValueError
            return clean(math.cos(v * d) / s_)

        funcs = {"sin": lambda v: clean(math.sin(v * d)), "cos": lambda v: clean(math.cos(v * d)),"tan": tan, "ln": math.log, "log": math.log10, "logb": math.log2,
                 "asin": lambda v: clean(math.asin(v) / d), "acos": lambda v: clean(math.acos(v) / d),
                 "atan": lambda v: clean(math.atan(v) / d),"sinh": math.sinh, "cosh": math.cosh, "tanh": math.tanh, "sqrt": math.sqrt, "cbrt": cbrt, "fact": fact,
                 "sec": sec, "csc": csc, "cot": cot, "todeg": math.degrees, "torad": math.radians,
                 "abs": abs, "floor": math.floor, "ceil": math.ceil, "round": round_half,
                 "mod": mod, "npr": npr, "ncr": ncr, "gcd": gcd, "lcm": lcm,
                 "gst": gst, "gstr": gstr, "disc": disc, "emi": emi, "chg": chg,
                 "max": lambda a, b: max(a, b), "min": lambda a, b: min(a, b)}
        consts = {"pi": math.pi, "E": math.e, "ANS": self.ans, "MEM": self.mem}
        return funcs, consts

    def evaluate(self, s):
        funcs, consts = self.make_env()
        tree = ast.parse(prepare(s), mode="eval")
        r = float(safe_eval(tree, consts, funcs))
        if not math.isfinite(r):
            raise ValueError
        return r

    def negate(self):
        s = self.expr
        self.done = False
        if re.fullmatch(r"[−-]?\d+\.?\d*(?:×10\^[−-]?\d+)?", s):
            self.expr = s[1:] if s[0] in "−-" else "−" + s
        else:
            m = re.search(r"\([−-](\d+\.?\d*)\)$", s)
            if m:
                self.expr = s[:m.start()] + m.group(1)
            else:
                m = re.search(r"(\d+\.?\d*)$", s)
                if m:
                    self.expr = s[:m.start()] + "(−" + m.group(0) + ")"
                else:
                    self.expr += "(−"
        self.refresh()

    def memory(self, v):
        if v == "MC":
            self.mem = 0.0
        else:
            try:
                val = self.evaluate(self.expr) if self.expr else self.ans
            except Exception:
                self.cv.itemconfig(self.t_res, text="Error", fill=ERR)
                return
            if v == "MS":
                self.mem = val
            elif v == "M+":
                self.mem += val
            else:
                self.mem -= val
            if abs(self.mem) < 1e-12:
                self.mem = 0.0
        self.cv.itemconfig(self.t_mode, text=self.mode_text())
    SUB = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")

    def log_start(self):
        self.lg, self.lg_base, self.lg_val = "base", "", ""
        self.log_show()

    def log_calc(self):
        y = float(self.lg_base)
        x = float(self.lg_val.replace("−", "-"))
        if y <= 1 or x <= 0:
            raise ValueError
        return math.log(x) / math.log(y)

    def log_show(self):
        cv = self.cv
        cv.coords(self.t_res, self.CW - 44, 158)
        if self.lg == "base":
            hint = f"y = {self.lg_base or 'base'}     press = ("
        else:
            hint = f"y = {self.lg_base}     x = {self.lg_val or 'value'}"
        cv.itemconfig(self.t_mode, fill=INK2, text=hint)
        cv.itemconfig(self.t_hist, text="")
        sub = self.lg_base.translate(self.SUB)
        if self.lg == "base":
            line = "log" + sub + "(" if sub else "logᵧ("
        else:
            line = "log" + sub + "(" + self.lg_val + ")"
        cv.itemconfig(self.t_expr, text=line, font=(FONT, 16 if len(line) <= 14 else 13))
        text, col = "", INK2
        if self.lg == "value" and self.lg_val not in ("", "−"):
            try:
                text = "= " + fmt(self.log_calc())
            except Exception:
                text, col = "Error", ERR
            cv.itemconfig(self.t_res, text=text, fill=col, font=(FONT, 16, "bold"))
    def log_input(self, v):
        digit = len(v) == 1 and v in "0123456789"
        if self.lg == "base":
            if v == "AC" or (v == "DEL" and not self.lg_base):
                return self.log_end()
            if v == "DEL":
                self.lg_base = self.lg_base[:-1]
            elif digit:
                if (not self.lg_base and v == "0") or len(self.lg_base) >= 8:
                    return
                self.lg_base += v
            elif v == ".":
                if not self.lg_base or "." in self.lg_base:
                    return
                self.lg_base += "."
            elif v in ("=", "("):
                try:
                    if float(self.lg_base) <= 1:
                        raise ValueError
                except ValueError:
                    self.cv.itemconfig(self.t_res, text="Error", fill=ERR)
                    return
                self.lg = "value"
            else:
                return
            return self.log_show()

        cur = self.lg_val
        if v == "AC":
            return self.log_end()
        if digit:
            cur += v
        elif v == ".":
            if "." in cur:
                return
            cur += "0." if cur in ("", "−") else "."
        elif v == "DEL":
            if cur == "":
                self.lg = "base"
                return self.log_show()
            cur = cur[:-1]
        elif v == "=":
            self.log_finish()
            return
        elif v in ("+", "−", "×", "÷", "^", "%"):
            if self.log_finish():
                self.press(v)
            return
        else:
            return
        self.lg_val = cur
        self.log_show()

    def log_finish(self):
        try:
            r = self.log_calc()
        except Exception:
            if self.lg_val not in ("", "−"):
                self.cv.itemconfig(self.t_res, text="Error", fill=ERR)
            return False
        txt = fmt(r)
        title = "log" + self.lg_base.translate(self.SUB) + "(" + self.lg_val + ")"
        self.history.append((title, txt))
        self.lg = None
        self.ans, self.expr, self.done = r, to_expr(txt), True
        self.cv.itemconfig(self.t_mode, text=self.mode_text(), fill=INK2)
        self.show_done(title, txt)
        return True

    def log_end(self):
        self.lg = None
        self.done = False
        self.cv.itemconfig(self.t_mode, text=self.mode_text())
        self.refresh()

    def press(self, v):
        if self.anim:
            return
        if self.lg and v in ("SETRAD", "SETDEG", "HIST"):
            return
        if v == "HIST":
            return self.toggle_history()
        if self.hview:
            self.toggle_history()
        if self.tab == "graph":
            if self.gscreen == "menu" and v.startswith("GRF"):
                self.open_graph(int(v[3:]))
            return
        if self.tab == "conv":
            return self.conv_press(v)
        if v in ("SETRAD", "SETDEG"):
            self.deg = (v == "SETDEG")
            self.mode_set = True
            self.update_mode_keys()
            self.cv.itemconfig(self.t_mode, text=self.mode_text())
            return self.refresh()
        if v == "LOGY":
            return self.log_start()
        if self.lg:
            return self.log_input(v)
        if v in ("MC", "MS", "M+", "M−"):
            return self.memory(v)
        if v == "NEG":
            return self.negate()
        if v == "AC":
            self.expr, self.done = "", False
            self.cv.itemconfig(self.t_hist, text="")
            return self.refresh()
        if v == "DEL":
            self.done = False
            self.expr = re.sub(r"([a-z]+\(|√\(|\^\(1÷|\^\(−1\)|\^[23]|(?:10|2|e)\^\(|Ans|Mem|.)$", "", self.expr)
            return self.refresh()
        if v == "=":
            if not self.expr:
              return
            try:
                r = self.evaluate(self.expr)
                txt = fmt(r)
            except Exception:
                self.cv.itemconfig(self.t_res, text="Error", fill=ERR)
                return
            shown = to_round(self.expr)
            self.history.append((shown, txt))
            self.ans, self.expr, self.done = r, to_expr(txt), True
            self.show_done(shown, txt)
            return
        if v == "MR":
            v = "Mem"
        if self.done:
            self.done = False
            if not re.match(r"[+−×÷^%!]", v):
                self.expr = ""
        last = self.expr[-1:]
        if not self.expr and v[:1] in ("×", "÷", "^", "%", "!", ")", "]", "}", ","):
            return
        if v == "." and "." in re.split(r"[^0-9.]", self.expr)[-1]:
            return
        if v == "." and not last.isdigit():
            v = "0."
        if len(v) == 1 and v.isdigit() and re.search(r"(?<![\d.])0$", self.expr):
            self.expr = self.expr[:-1]
        if v in ("+", "−", "×", "÷") and last in ("+", "−", "×", "÷"):
            if v == "−" and last in ("×", "÷"):
                pass
            else:
                base = self.expr.rstrip("+−×÷")
                if not base and v in ("×", "÷"):
                    return
                self.expr = base
        tail = last.isdigit() or last in (".", ")", "π", "e") or self.expr.endswith(("Ans", "Mem"))
        if v in ("10^(", "2^(") and tail:
            v = "×" + v
        self.expr += v
        self.refresh()
    def gx2s(self, x): return (PX0 + self.PX1) / 2 + (x - self.gcx) * self.gsx
    def gy2s(self, y): return (PY0 + PY1) / 2 - (y - self.gcy) * self.gsy
    def gs2x(self, sx): return self.gcx + (sx - (PX0 + self.PX1) / 2) / self.gsx
    def gs2y(self, sy): return self.gcy - (sy - (PY0 + PY1) / 2) / self.gsy

    # Graph
    def graph_reset_view(self):
        self.gcx = self.gcy = 0.0
        if self.gdeg:
            self.gsx = (self.PX1 - PX0) / 760
            self.gsy = (PY1 - PY0) / 3.2
        else:
            self.gsx = self.gsy = (self.PX1 - PX0) / 16

    def gstep_x(self, span):
        if self.gxdeg:
            for s in (1, 2, 5, 10, 15, 30, 45, 90, 180, 360, 720, 1440, 2880):
                if span / s <= 12:
                    return s
        return nice_step(span)

    def gchip(self, tag, x1, cy, w, text, act, size=10, color="#ffffff", sel=False):
        tg = ("cui", tag)
        self.round_rect(x1, cy - 14, x1 + w, cy + 14, 11, fill="#2c2c2e" if sel else "#151517",
                        outline="" if sel else "#3a3a3c", width=1, tags=tg)
        t = self.cv.create_text(x1 + w / 2, cy, text=text, fill=color,
                                font=(FONT, self.fit_size(text, w - 8, size, 7), "bold"), tags=tg)
        self.gbtns.append((x1, cy - 14, x1 + w, cy + 14, act))
        self.unit_groups.append((tag, x1 + w / 2, cy))
        return t

    def draw_back_btn(self):
        cv = self.cv
        t = ("cui", "gback")
        self.round_rect(24, 104, 112, 132, 12, fill="#151517", outline="#3a3a3c", width=1, tags=t)
        cv.create_text(68, 118, text="‹  Back", fill="#ffffff", font=(FONT, 11, "bold"), tags=t)
        self.gbtns.append((24, 104, 112, 132, "BACK"))
        self.unit_groups.append(("gback", 68, 118))

    def draw_graph_screen(self):
        cv = self.cv
        CW = self.CW
        cv.delete("cui")
        g = self.unit_groups = []
        self.gbtns = []
        self.draw_back_btn()
        t = ("cui", "gin")
        self.round_rect(24, 138, CW - 24, 178, 14, fill="#15171b", outline=ACCENT, width=2, tags=t)
        cv.create_text(40, 158, text="y =", anchor="w", fill=ACCENT, font=(FONT, 16, "bold"), tags=t)
        self.t_gin = cv.create_text(80, 158, text="", anchor="w", fill="#ffffff", font=(FONT, 18), tags=t)
        g.append(("gin", CW / 2, 158))
        self.t_gread = cv.create_text(28, 188, text="", anchor="w", fill=TAB_OFF, font=(FONT, 9), tags=("cui", "gread"))
        g.append(("gread", CW / 2, 188))
        cv.create_rectangle(PX0, PY0, self.PX1, PY1, fill=LCD, outline="#15171b", width=2, tags=("cui", "gplot"))
        g.append(("gplot", CW / 2, (PY0 + PY1) / 2))

        cy = 452
        p6 = (CW - 42) / 6
        for i, (txt, act) in enumerate((("Zoom +", "ZIN"), ("Zoom −", "ZOUT"), ("Fit", "FIT"),
                                        ("RAD", "RAD"), ("DEG", "DEG"), ("Clear", "CLR"))):
            on = (act == "DEG" and self.gdeg) or (act == "RAD" and not self.gdeg)
            self.gchip(f"gk{i}", 24 + i * p6, cy, p6 - 6, txt, act, 10, TAB_ON if on else "#ffffff", on)
        F, O = "#ffffff", ACCENT


        KEYS = [
            [("sin", "I:sin(", F), ("cos", "I:cos(", F), ("tan", "I:tan(", F), ("ln", "I:ln(", F), ("log", "I:log(", F), ("√x", "I:sqrt(", F)],
            [("sin⁻¹", "I:asin(", F), ("cos⁻¹", "I:acos(", F), ("tan⁻¹", "I:atan(", F), ("sinh", "I:sinh(", F), ("cosh", "I:cosh(", F), ("tanh", "I:tanh(", F)],
            [("7", "I:7", F), ("8", "I:8", F), ("9", "I:9", F), ("(", "I:(", O), (")", "I:)", O), ("⌫", "BS", TAB_ON)],
            [("4", "I:4", F), ("5", "I:5", F), ("6", "I:6", F), ("×", "I:×", O), ("÷", "I:÷", O), ("x", "I:x", O)],
            [("1", "I:1", F), ("2", "I:2", F), ("3", "I:3", F), ("+", "I:+", O), ("−", "I:−", O), ("π", "I:π", O)],
            [("0", "I:0", F), (".", "I:.", F), ("x²", "I:^2", O), ("xʸ", "I:^", O), ("|x|", "I:abs(", F), ("e", "I:e", O)],
        ]

        p6 = (CW - 42) / 6
        for r, row in enumerate(KEYS):
            for c, (lab, act, col) in enumerate(row):
                self.gchip(f"gq{r}{c}", 24 + c * p6, 490 + r * 36, p6 - 6, lab, act, 11, col)
        cv.create_text(CW / 2, 720, text="Used This Type Keys:  sin(x),  sinx,  sin90,  asin(x)",fill=TAB_OFF, font=(FONT, 9), tags=("cui", "ghelp"))
        g.append(("ghelp", CW / 2, 720))
        self.graph_update()

    def graph_compile(self):
        self.gfn, self.gstate = None, "empty"
        s = self.gexpr.strip()
        if not s:
            return
        try:
            tree = ast.parse(gprep(s), mode="eval")
            funcs, consts = self.make_env(self.gdeg)
            consts["X"] = 0.0
            ok = set(funcs) | set(consts)
            for n in ast.walk(tree):
                if isinstance(n, ast.Name) and n.id not in ok:
                    raise ValueError
            has_x = any(isinstance(n, ast.Name) and n.id == "X" for n in ast.walk(tree))
        except Exception:
            self.gstate = "bad"
            return

        def f(x):
            consts["X"] = x
            try:
                v = float(safe_eval(tree, consts, funcs))
            except Exception:
                return None
            return v if math.isfinite(v) else None

        self.gfn, self.gstate = f, "ok"
        if not has_x:
            v = f(0.0)
            self.gstate = "const"
            self.gconst_txt = (f"{s} = " + ("undefined" if v is None else f"{v:.10g}") + f"[{'DEG' if self.gdeg else 'RAD'}]")

    def graph_autoview(self):
        e = self.gexpr
        if re.search(r"a(sin|cos)", e):
            xr, self.gxdeg = 1.5, False
        elif self.gdeg and re.search(r"(?<!a)(sin|cos|tan|sec|csc|cot)(?!h)", e) and "x" in e:
            xr, self.gxdeg = 360.0, True
        else:
            xr, self.gxdeg = 8.0, False
        self.gcx = 0.0
        self.gsx = (self.PX1 - PX0) / (2 * xr)
        ys = []
        if self.gfn:
            n = 240
            for i in range(n + 1):
                y = self.gfn(-xr + 2 * xr * i / n)
                if y is not None and abs(y) < 1e12:
                    ys.append(y)
        if len(ys) >= 3:
            ys.sort()
            lo = ys[int(0.05 * (len(ys) - 1))]
            hi = ys[int(0.95 * (len(ys) - 1))]
            span = hi - lo
            if span < 1e-9:
                lo, hi = lo - 1, hi + 1
                span = 2.0
            lo, hi = lo - 0.25 * span, hi + 0.25 * span
            self.gcy = (lo + hi) / 2
            self.gsy = (PY1 - PY0) / (hi - lo)
        else:
            self.gcy = 0.0
            self.gsy = self.gsx

    def graph_update(self):
        self.graph_compile()
        self.graph_autoview()
        s = self.gexpr + "|"
        if len(s) > 28:
            s = "…" + s[-27:]
        self.cv.itemconfig(self.t_gin, text=s, fill=TAB_ON if self.gstate == "bad" else "#ffffff",font=(FONT, 18 if len(s) <= 20 else 14))
        self.graph_redraw()

    def graph_status(self):
        txt, col = {"empty": ("Type a function of x  e.g. sin(x), x^2, ln(x)", TAB_OFF),
                    "bad": ("Invalid function", TAB_ON),
                    "const": (self.gconst_txt, "#ffffff"),
                    "ok": ("Move the mouse over the graph to see x and y", TAB_OFF)}[self.gstate]
        self.cv.itemconfig(self.t_gread, text=txt, fill=col)

    def graph_redraw(self):
        cv = self.cv
        PX1 = self.PX1
        cv.delete("gdyn")
        cv.delete("gtrace")
        T = ("cui", "gplot", "gdyn")
        deg = "°" if self.gxdeg else ""
        x_lo, x_hi = self.gs2x(PX0), self.gs2x(PX1)
        y_hi, y_lo = self.gs2y(PY0), self.gs2y(PY1)
        stepx = self.gstep_x(x_hi - x_lo)
        stepy = nice_step(y_hi - y_lo)
        ly = min(max(self.gy2s(0) + 9, PY0 + 9), PY1 - 9)
        lx = min(max(self.gx2s(0) + 4, PX0 + 3), PX1 - 34)

        for k in range(math.ceil(x_lo / stepx), math.floor(x_hi / stepx) + 1):
            sx = self.gx2s(k * stepx)
            cv.create_line(sx, PY0, sx, PY1, fill=INK if k == 0 else "#a8b896", width=2 if k == 0 else 1, tags=T)
            if k:
                cv.create_text(sx, ly, text=f"{k * stepx:.6g}{deg}", fill=INK2, font=(FONT, 8), tags=T)
        for k in range(math.ceil(y_lo / stepy), math.floor(y_hi / stepy) + 1):
            sy = self.gy2s(k * stepy)
            cv.create_line(PX0, sy, PX1, sy, fill=INK if k == 0 else "#a8b896", width=2 if k == 0 else 1, tags=T)
            if k:
                cv.create_text(lx, sy - 8, text=f"{k * stepy:.6g}", anchor="w", fill=INK2, font=(FONT, 8), tags=T)

        if self.gfn:
            yspan = (PY1 - PY0) / self.gsy
            run = []

            def flush():
                if len(run) >= 4:
                    cv.create_line(*run, fill=INK, width=2, joinstyle="round", tags=T)
                run.clear()

            prev = None
            for i in range(int(PX1 - PX0) + 1):
                sx = PX0 + i
                y = self.gfn(self.gs2x(sx))
                if y is None:
                    flush()
                    prev = None
                    continue
                y = max(min(y, 1e9), -1e9)
                sy = self.gy2s(y)
                if prev is not None and y * prev[2] < 0 and abs(y - prev[2]) > yspan:
                    flush()
                    prev = None
                if prev is not None:
                    seg = clip_y(prev[0], prev[1], sx, sy, PY0, PY1)
                    if seg is None:
                        flush()
                    else:
                        a, b, c, d = seg
                        if run and run[-2] == a and run[-1] == b:
                            run += [c, d]
                        else:
                            flush()
                            run += [a, b, c, d]
                prev = (sx, sy, y)
            flush()
        self.graph_status()

    def graph_zoom(self, f, sx=None, sy=None):
        mx, my = (PX0 + self.PX1) / 2, (PY0 + PY1) / 2
        if sx is None or not (PX0 <= sx <= self.PX1 and PY0 <= sy <= PY1):
            sx, sy = mx, my
        nx, ny = self.gsx * f, self.gsy * f
        if not (1e-4 <= nx <= 1e6 and 1e-4 <= ny <= 1e6):
            return
        x, y = self.gs2x(sx), self.gs2y(sy)
        self.gsx, self.gsy = nx, ny
        self.gcx = x - (sx - mx) / nx
        self.gcy = y + (sy - my) / ny
        self.graph_redraw()

    def graph_insert(self, s):
        if self.gfresh:
            self.gexpr, self.gfresh = "", False
        if len(self.gexpr) + len(s) <= 60:
            self.gexpr += s
        self.graph_update()

    def graph_back(self):
        self.gfresh = False
        self.gexpr = re.sub(r"([a-z]{2,}\(|.)$", "", self.gexpr)
        self.graph_update()

    def graph_action(self, act):
        if act == "ZIN":
            self.graph_zoom(1.25)
        elif act == "ZOUT":
            self.graph_zoom(0.8)
        elif act == "FIT":
            self.graph_autoview()
            self.graph_redraw()
        elif act in ("RAD", "DEG"):
            want = (act == "DEG")
            if want != self.gdeg:
                self.gdeg = want
                self.draw_graph_screen()
        elif act == "CLR":
            self.gexpr, self.gfresh = "", False
            self.graph_update()
        elif act == "BS":
            self.graph_back()
        elif act == "BACK":
            self.graph_menu()
        elif act.startswith("D:"):
            self.data_action(act[2:])
        elif act.startswith("I:"):
            self.graph_insert(act[2:])

    def graph_click(self, e):
        if self.gscreen == "menu":
            return False
        if self.gscreen == "data":
            for x1, y1, x2, y2, act in self.dbtns:
                if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                    self.graph_action(act)
                    return True
            if self.ddopen and PX0 <= e.x <= self.PX1 and PY0 <= e.y <= PY1:
                self.ddopen = False
                self.data_redraw()
                return True
        for x1, y1, x2, y2, act in self.gbtns:
            if x1 <= e.x <= x2 and y1 <= e.y <= y2:
                self.graph_action(act)
                return True
        if self.gscreen == "trig" and PX0 <= e.x <= self.PX1 and PY0 <= e.y <= PY1:
            self.gpan = (e.x, e.y, self.gcx, self.gcy)
            return True
        return False

    def graph_key(self, e):
        ch, ks = e.char, e.keysym
        if self.gscreen == "menu":
            if ch in ("1", "2"):
                self.open_graph(int(ch) - 1)
            return
        if self.gscreen == "data":
            return self.data_key(e)
        if e.state & 0x4:
            if ks in ("v", "V"):
                try:
                    txt = self.root.clipboard_get()
                except tk.TclError:
                    return
                self.gexpr = "".join(c for c in txt.lower() if c in GCHARS)[:60]
                self.gfresh = False
                return self.graph_update()
            if ks in ("c", "C"):
                self.root.clipboard_clear()
                self.root.clipboard_append(self.gexpr)
            return
        if ks == "BackSpace":
            return self.graph_back()
        if ks == "Escape":
            return self.graph_action("CLR")
        if ks in ("Left", "Right", "Up", "Down"):
            dx, dy = 40 / self.gsx, 40 / self.gsy
            self.gcx += {"Left": -dx, "Right": dx}.get(ks, 0)
            self.gcy += {"Up": dy, "Down": -dy}.get(ks, 0)
            return self.graph_redraw()
        if ks == "Prior":
            return self.graph_zoom(1.25)
        if ks == "Next":
            return self.graph_zoom(0.8)
        if ks == "Home":
            return self.graph_action("FIT")
        if ch and ch.lower() in GCHARS:
            self.graph_insert(ch.lower())

    def draw_data_screen(self):
        cv = self.cv
        CW = self.CW
        cv.delete("cui")
        g = self.unit_groups = []
        self.gbtns, self.dbtns, self.ddopen = [], [], False
        self.draw_back_btn()

        self.round_rect(24, 138, CW - 116, 178, 14, fill="#15171b", outline=ACCENT, width=2, tags=("cui", "gin"))
        cv.create_text(44, 158, text="📂", anchor="w", fill=ACCENT, font=("Segoe UI Emoji", 14), tags=("cui", "gin"))
        cv.create_text(76, 158, text="Upload CSV / Excel file", anchor="w", fill="#ffffff", font=(FONT, 12, "bold"), tags=("cui", "gin"))
        self.gbtns.append((24, 138, CW - 116, 178, "D:UP"))
        g.append(("gin", CW / 2, 158))
        if self.dseries or self.dmsg:
            self.gchip("dclr", CW - 104, 158, 80, "✕ Clear", "D:CLR", 11, TAB_ON)

        self.t_gread = cv.create_text(28, 188, text="", anchor="w", fill=TAB_OFF, font=(FONT, 9), tags=("cui", "gread"))
        g.append(("gread", CW / 2, 188))

        cv.create_rectangle(PX0, PY0, self.PX1, PY1, fill=LCD, outline="#15171b", width=2, tags=("cui", "gplot"))
        g.append(("gplot", CW / 2, (PY0 + PY1) / 2))

        self.round_rect(24, 452, CW - 24, 560, 14, fill="#15171b", outline="#3a3a3c", width=1, tags=("cui", "dstat"))
        g.append(("dstat", CW / 2, 506))

        if len(self.dseries) > 1:
            cv.create_text(28, 584, text="SELECT COLUMN  (used for stats, pie, donut, histogram)", anchor="w",
                           fill=TAB_OFF, font=(FONT, 8, "bold"), tags=("cui", "dlab"))
            g.append(("dlab", 40, 584))
            p6 = (CW - 42) / 6
            for i, (nm, _) in enumerate(self.dseries[:6]):
                on = i == self.dsel
                self.gchip(f"ds{i}", 24 + i * p6, 610, p6 - 6, nm[:9], f"D:S:{i}", 10, TAB_ON if on else "#ffffff", on)
        g.append(("dhelp", CW / 2, 721))
        self.data_redraw()

    def data_use(self, labs, series, name):
        self.dlabs, self.dseries, self.dsel, self.dfile, self.dmsg = labs, series, 0, name, ""
        self.draw_data_screen()

    def data_upload(self):
        self.root.attributes("-topmost", False)
        try:
            path = filedialog.askopenfilename(parent=self.root, title="Select data file",filetypes=[("Data files", "*.csv *.xlsx *.xlsm *.tsv *.txt"), ("All files", "*.*")])
        finally:
            self.root.attributes("-topmost", True)
            self.root.focus_force()
        if not path:
            return
        try:
            labs, series = table_to_data(read_table(path))
        except Exception:
            self.dmsg = ("Old .xls not supported - save it as .xlsx or .csv"
                if path.lower().endswith(".xls")
                else "Could not read numbers from this file")
            return self.data_redraw()
        self.data_use(labs, series, os.path.basename(path))

    def data_action(self, act):
        if act == "DD":
            self.ddopen = not self.ddopen
            return self.data_redraw()
        if act.startswith("C:"):
            self.dtype, self.ddopen = act[2:], False
            return self.data_redraw()
        if act.startswith("S:"):
            self.dsel = int(act[2:])
            return self.draw_data_screen()
        if act == "UP":
            return self.data_upload()
        if act == "CLR":
            self.dlabs, self.dseries, self.dsel, self.dfile, self.dmsg = [], [], 0, "", ""
            self.draw_data_screen()

    def data_key(self, e):
        ks = e.keysym
        if e.state & 0x4:
            if ks in ("o", "O"):
                return self.data_upload()
            if ks in ("v", "V"):
                try:
                    txt = self.root.clipboard_get()
                except tk.TclError:
                    return
                try:
                    labs, series = table_to_data(text_to_rows(txt))
                except Exception:
                    self.dmsg = "Clipboard does not have a table of numbers"
                    return self.data_redraw()
                return self.data_use(labs, series, "Pasted data")
            return
        if ks == "Escape":
            if self.ddopen:
                self.ddopen = False
                return self.data_redraw()
            return self.data_action("CLR")

    def data_stats(self, vals):
        cv = self.cv
        n = len(vals)
        cells = [("Count", "—"), ("Sum", "—"), ("Mean", "—"), ("Median", "—"), ("Min", "—"), ("Max", "—"), ("Std Dev", "—"), ("Range", "—")]
        if n:
            s = sorted(vals)
            tot = sum(vals)
            mean = tot / n
            med = s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2
            sd = math.sqrt(sum((v - mean) ** 2 for v in vals) / n)
            nums = [n, tot, mean, med, s[0], s[-1], sd, s[-1] - s[0]]
            cells = [(c[0], str(n) if i == 0 else f"{nums[i]:.6g}") for i, c in enumerate(cells)]
        cw = (self.CW - 48) / 4
        for i, (lab, val) in enumerate(cells):
            r, c = divmod(i, 4)
            x = 24 + c * cw + cw / 2
            y0 = 460 + r * 48
            cv.create_text(x, y0 + 12, text=lab, fill=TAB_OFF, font=(FONT, 8, "bold"), tags=("cui", "dstat", "ddyn"))
            cv.create_text(x, y0 + 32, text=val, fill="#ffffff", font=(FONT, 12 if len(val) <= 9 else 10, "bold"),tags=("cui", "dstat", "ddyn"))

    def data_redraw(self):
        cv = self.cv
        cv.delete("ddyn")
        self.dbtns = []
        T = ("cui", "gplot", "ddyn")
        labs, series = self.dlabs, self.dseries
        vals = series[min(self.dsel, len(series) - 1)][1] if series else []
        cname = dict(DCHARTS)[self.dtype]
        if self.dmsg:
            cv.itemconfig(self.t_gread, fill=ERR, text=self.dmsg)
        elif series:
            cv.itemconfig(self.t_gread, fill=TAB_OFF,text=f"{self.dfile}  ·  {len(labs)} rows × {len(series)} column{'s' if len(series) != 1 else ''}  ·  {cname}")
        else:
            cv.itemconfig(self.t_gread, fill=TAB_OFF, text=" ")
        self.data_stats(vals)
        if series:
            getattr(self, "dc_" + self.dtype)(T, labs, series)
        else:
            cv.create_text((PX0 + self.PX1) / 2, (PY0 + PY1) / 2 + 10, text="No data yet", fill=INK2,font=(FONT, 12, "bold"), tags=T)
        self.data_dropdown(T)

    def data_dropdown(self, T):
        cv = self.cv
        bx1, by1, bx2, by2 = self.PX1 - 140, PY0 + 6, self.PX1 - 6, PY0 + 30
        my = (by1 + by2) / 2
        self.round_rect(bx1, by1, bx2, by2, 8, fill="#15171b", outline=ACCENT, width=1, tags=T)
        cv.create_text(bx1 + 10, my, text=dict(DCHARTS)[self.dtype], anchor="w", fill="#ffffff",font=(FONT, 10, "bold"), tags=T)
        cv.create_text(bx2 - 10, my, text="▴" if self.ddopen else "▾", anchor="e", fill=ACCENT,font=(FONT, 10, "bold"), tags=T)
        self.dbtns.append((bx1, by1, bx2, by2, "D:DD"))
        if not self.ddopen:
            return
        cw_, rh = 130, 29
        px2 = self.PX1 - 6
        px1 = px2 - 2 * cw_ - 8
        py1 = by2 + 4
        self.round_rect(px1, py1, px2, py1 + 6 * rh + 8, 10, fill="#15171b", outline="#3a3a3c", width=1, tags=T)
        for i, (key, nm) in enumerate(DCHARTS):
            c, r = divmod(i, 6)
            x1 = px1 + 4 + c * cw_
            x2 = x1 + cw_ - 2
            y1 = py1 + 4 + r * rh
            y2 = y1 + rh - 3
            on = key == self.dtype
            if on:
                self.round_rect(x1, y1, x2, y2, 8, fill="#2c2c2e", outline="", tags=T)
            cv.create_text(x1 + 10, (y1 + y2) / 2, text=nm, anchor="w", fill=TAB_ON if on else "#ffffff",font=(FONT, 10, "bold"), tags=T)
            self.dbtns.append((x1, y1, x2, y2, "D:C:" + key))

    def dgeo(self):
        return PX0 + 40, self.PX1 - 14, PY0 + 38, PY1 - 26

    def dcol(self, s, m):
        return INK if m == 1 else PIE[s % len(PIE)]

    def data_yaxis(self, T, lo, hi, x0, x1, y0, y1, zero=True):
        cv = self.cv
        if zero:
            lo, hi = min(0, lo), max(0, hi)
        if hi == lo:
            lo, hi = lo - 1, hi + 1
        elif not zero:
            pad = (hi - lo) * 0.08
            lo, hi = lo - pad, hi + pad
        step = nice_step(hi - lo)
        lo = math.floor(lo / step) * step
        hi = math.ceil(hi / step) * step
        ys = lambda v: y1 - (v - lo) / (hi - lo) * (y1 - y0)
        k = math.ceil(lo / step - 1e-9)
        while k * step <= hi + 1e-9:
            v = k * step
            sy = ys(v)
            zr = abs(v) < 1e-12 * max(1, abs(step))
            cv.create_line(x0, sy, x1, sy, fill=INK if zr else "#a8b896", width=2 if zr else 1, tags=T)
            cv.create_text(x0 - 4, sy, text=f"{0 if zr else v:.6g}", anchor="e", fill=INK2, font=(FONT, 8), tags=T)
            k += 1
        return ys

    def data_xlabels(self, T, cxs, labs):
        every = max(1, math.ceil(len(labs) / 12))
        for i, (cx, lab) in enumerate(zip(cxs, labs)):
            if i % every == 0:
                self.cv.create_text(cx, PY1 - 12, text=lab[:7], fill=INK2, font=(FONT, 8), tags=T)

    def data_legend(self, T, series):
        if len(series) < 2:
            return
        x = PX0 + 10
        for i, (nm, _) in enumerate(series):
            t = nm[:10]
            if x + 14 + 6 * len(t) > self.PX1 - 150:
                break
            self.cv.create_rectangle(x, PY0 + 13, x + 10, PY0 + 23, fill=PIE[i % len(PIE)], outline="", tags=T)
            self.cv.create_text(x + 14, PY0 + 18, text=t, anchor="w", fill=INK, font=(FONT, 9), tags=T)
            x += 26 + 6 * len(t)

    def dcat(self, n):
        x0, x1, y0, y1 = self.dgeo()
        slot = (x1 - x0) / n
        return [x0 + slot * (i + 0.5) for i in range(n)], slot

    def dc_bar(self, T, labs, series):
        cv = self.cv
        x0, x1, y0, y1 = self.dgeo()
        m, n = len(series), len(labs)
        allv = [v for _, vs in series for v in vs]
        ys = self.data_yaxis(T, min(allv), max(allv), x0, x1, y0, y1)
        cxs, slot = self.dcat(n)
        base = ys(0)
        gw = min(slot * 0.8, 60 * m)
        bw = gw / m
        for s, (nm, vs) in enumerate(series):
            c = self.dcol(s, m)
            for i, v in enumerate(vs):
                bx = cxs[i] - gw / 2 + bw * s
                sy = ys(v)
                cv.create_rectangle(bx, min(sy, base), bx + bw, max(sy, base), fill=c, outline=LCD if m > 1 else "", tags=T)
                if n * m <= 12:
                    cv.create_text(bx + bw / 2, sy - 9 if v >= 0 else sy + 9, text=f"{v:.4g}", fill=INK,font=(FONT, 8, "bold"), tags=T)
        self.data_xlabels(T, cxs, labs)
        self.data_legend(T, series)

    def dc_stack(self, T, labs, series):
        cv = self.cv
        x0, x1, y0, y1 = self.dgeo()
        m, n = len(series), len(labs)
        pos = [sum(max(vs[i], 0) for _, vs in series) for i in range(n)]
        neg = [sum(min(vs[i], 0) for _, vs in series) for i in range(n)]
        ys = self.data_yaxis(T, min(neg), max(pos), x0, x1, y0, y1)
        cxs, slot = self.dcat(n)
        bw = min(slot * 0.62, 60)
        up, dn = [0.0] * n, [0.0] * n
        for s, (nm, vs) in enumerate(series):
            c = self.dcol(s, m)
            for i, v in enumerate(vs):
                if v >= 0:
                    a, b = up[i], up[i] + v
                    up[i] = b
                else:
                    a, b = dn[i], dn[i] + v
                    dn[i] = b
                cv.create_rectangle(cxs[i] - bw / 2, ys(max(a, b)), cxs[i] + bw / 2, ys(min(a, b)), fill=c,outline=LCD if m > 1 else "", tags=T)
        self.data_xlabels(T, cxs, labs)
        self.data_legend(T, series)

    def dc_hbar(self, T, labs, series):
        cv = self.cv
        x0, x1, y0, y1 = PX0 + 54, self.PX1 - 18, PY0 + 38, PY1 - 22
        m, n = len(series), len(labs)
        allv = [v for _, vs in series for v in vs]
        lo, hi = min(0, min(allv)), max(0, max(allv))
        if hi == lo:
            hi = lo + 1
        step = nice_step(hi - lo)
        lo = math.floor(lo / step) * step
        hi = math.ceil(hi / step) * step
        xs = lambda v: x0 + (v - lo) / (hi - lo) * (x1 - x0)
        k = math.ceil(lo / step - 1e-9)
        while k * step <= hi + 1e-9:
            v = k * step
            sx = xs(v)
            zr = abs(v) < 1e-12 * max(1, abs(step))
            cv.create_line(sx, y0, sx, y1, fill=INK if zr else "#a8b896", width=2 if zr else 1, tags=T)
            cv.create_text(sx, y1 + 10, text=f"{0 if zr else v:.6g}", fill=INK2, font=(FONT, 8), tags=T)
            k += 1
        slot = (y1 - y0) / n
        base = xs(0)
        gh = min(slot * 0.8, 40 * m)
        bh = gh / m
        every = max(1, math.ceil(n / 12))
        for i in range(n):
            cy = y0 + slot * (i + 0.5)
            if i % every == 0:
                cv.create_text(x0 - 4, cy, text=labs[i][:7], anchor="e", fill=INK2, font=(FONT, 8), tags=T)
            for s, (nm, vs) in enumerate(series):
                v = vs[i]
                sx = xs(v)
                by = cy - gh / 2 + bh * s
                cv.create_rectangle(min(sx, base), by, max(sx, base), by + bh, fill=self.dcol(s, m),outline=LCD if m > 1 else "", tags=T)
                if n * m <= 12:
                    cv.create_text(sx + 4 if v >= 0 else sx - 4, by + bh / 2, text=f"{v:.4g}", anchor="w" if v >= 0 else "e",fill=INK, font=(FONT, 8, "bold"), tags=T)
        self.data_legend(T, series)

    def dc_line(self, T, labs, series):
        cv = self.cv
        x0, x1, y0, y1 = self.dgeo()
        m, n = len(series), len(labs)
        allv = [v for _, vs in series for v in vs]
        ys = self.data_yaxis(T, min(allv), max(allv), x0, x1, y0, y1)
        cxs, slot = self.dcat(n)
        for s, (nm, vs) in enumerate(series):
            c = self.dcol(s, m)
            pts = []
            for cx, v in zip(cxs, vs):
                pts += [cx, ys(v)]
            if n >= 2:
                cv.create_line(*pts, fill=c, width=2, joinstyle="round", tags=T)
            for cx, v in zip(cxs, vs):
                sy = ys(v)
                cv.create_oval(cx - 4, sy - 4, cx + 4, sy + 4, fill=ERR if m == 1 else c, outline="", tags=T)
                if m == 1 and n <= 12:
                    cv.create_text(cx, sy - 9 if v >= 0 else sy + 9, text=f"{v:.4g}", fill=INK, font=(FONT, 8, "bold"), tags=T)
        self.data_xlabels(T, cxs, labs)
        self.data_legend(T, series)

    def dc_area(self, T, labs, series):
        cv = self.cv
        x0, x1, y0, y1 = self.dgeo()
        m, n = len(series), len(labs)
        allv = [v for _, vs in series for v in vs]
        ys = self.data_yaxis(T, min(allv), max(allv), x0, x1, y0, y1)
        cxs, slot = self.dcat(n)
        base = ys(0)
        for s, (nm, vs) in enumerate(series):
            c = self.dcol(s, m)
            pts = []
            for cx, v in zip(cxs, vs):
                pts += [cx, ys(v)]
            poly = [cxs[0], base] + pts + [cxs[-1], base]
            kw = {"stipple": "gray50"} if m > 1 else {}
            cv.create_polygon(*poly, fill=c, outline="", tags=T, **kw)
            if n >= 2:
                cv.create_line(*pts, fill=c, width=2, joinstyle="round", tags=T)
        self.data_xlabels(T, cxs, labs)
        self.data_legend(T, series)

    def dc_scatter(self, T, labs, series):
        cv = self.cv
        x0, x1, y0, y1 = self.dgeo()
        m, n = len(series), len(labs)
        allv = [v for _, vs in series for v in vs]
        ys = self.data_yaxis(T, min(allv), max(allv), x0, x1, y0, y1, zero=False)
        xv = [to_num(l) for l in labs]
        if all(v is not None for v in xv):
            lo, hi = min(xv), max(xv)
            if hi == lo:
                lo, hi = lo - 1, hi + 1
            px = [x0 + 12 + (v - lo) / (hi - lo) * (x1 - x0 - 24) for v in xv]
            for k in range(6):
                tv = lo + (hi - lo) * k / 5
                cv.create_text(x0 + 12 + (x1 - x0 - 24) * k / 5, PY1 - 12, text=f"{tv:.4g}", fill=INK2, font=(FONT, 8), tags=T)
        else:
            px, _ = self.dcat(n)
            self.data_xlabels(T, px, labs)
        for s, (nm, vs) in enumerate(series):
            c = ERR if m == 1 else self.dcol(s, m)
            for cx, v in zip(px, vs):
                sy = ys(v)
                cv.create_oval(cx - 4, sy - 4, cx + 4, sy + 4, fill=c, outline="", tags=T)
        self.data_legend(T, series)

    def dc_pie(self, T, labs, series):
        self.data_pie(labs, series[min(self.dsel, len(series) - 1)][1], T, False)

    def dc_donut(self, T, labs, series):
        self.data_pie(labs, series[min(self.dsel, len(series) - 1)][1], T, True)

    def dc_hist(self, T, labs, series):
        cv = self.cv
        vals = series[min(self.dsel, len(series) - 1)][1]
        n = len(vals)
        k = min(15, max(3, int(math.log2(n)) + 2)) if n > 1 else 3
        lo, hi = min(vals), max(vals)
        if hi == lo:
            lo, hi = lo - 1, hi + 1
        w = (hi - lo) / k
        counts = [0] * k
        for v in vals:
            counts[min(k - 1, int((v - lo) / w))] += 1
        x0, x1, y0, y1 = self.dgeo()
        ys = self.data_yaxis(T, 0, max(counts), x0, x1, y0, y1)
        slot = (x1 - x0) / k
        base = ys(0)
        every = max(1, math.ceil(k / 8))
        for i, c in enumerate(counts):
            bx = x0 + slot * i
            if c:
                cv.create_rectangle(bx + 1, ys(c), bx + slot - 1, base, fill=INK, outline="", tags=T)
                cv.create_text(bx + slot / 2, ys(c) - 9, text=str(c), fill=INK, font=(FONT, 8, "bold"), tags=T)
            if i % every == 0:
                cv.create_text(bx, PY1 - 12, text=f"{lo + w * i:.3g}", fill=INK2, font=(FONT, 8), tags=T)

    def dc_radar(self, T, labs, series):
        cv = self.cv
        n = min(len(labs), 12)
        cx, cy = (PX0 + self.PX1) / 2, (PY0 + 36 + PY1) / 2
        if n < 3:
            cv.create_text(cx, cy, text="Radar needs at least 3 rows", fill=ERR, font=(FONT, 11, "bold"), tags=T)
            return
        m = len(series)
        mx = max((abs(v) for _, vs in series for v in vs[:n]), default=1) or 1
        R = min((self.PX1 - PX0) / 2 - 50, (PY1 - PY0 - 36) / 2 - 18)
        ang = lambda i: -math.pi / 2 + 2 * math.pi * i / n
        for ring in (0.25, 0.5, 0.75, 1.0):
            pts = []
            for i in range(n):
                pts += [cx + R * ring * math.cos(ang(i)), cy + R * ring * math.sin(ang(i))]
            cv.create_polygon(*pts, fill="", outline="#a8b896", width=1, tags=T)
        for i in range(n):
            a = ang(i)
            cv.create_line(cx, cy, cx + R * math.cos(a), cy + R * math.sin(a), fill="#a8b896", tags=T)
            cv.create_text(cx + (R + 16) * math.cos(a), cy + (R + 10) * math.sin(a), text=labs[i][:6], fill=INK2,font=(FONT, 8), tags=T)
        for s, (nm, vs) in enumerate(series):
            c = self.dcol(s, m)
            pts = []
            for i in range(n):
                r_ = R * max(vs[i], 0) / mx
                pts += [cx + r_ * math.cos(ang(i)), cy + r_ * math.sin(ang(i))]
            cv.create_polygon(*pts, fill="", outline=c, width=2, tags=T)
            for j in range(0, len(pts), 2):
                cv.create_oval(pts[j] - 3, pts[j + 1] - 3, pts[j] + 3, pts[j + 1] + 3, fill=c, outline="", tags=T)
        self.data_legend(T, series)

    def data_pie(self, labs, vals, T, hole=False):
        cv = self.cv
        tot = sum(vals)
        if any(v < 0 for v in vals) or tot <= 0:
            cv.create_text((PX0 + self.PX1) / 2, (PY0 + PY1) / 2 + 10, text="Pie chart needs positive numbers",fill=ERR, font=(FONT, 11, "bold"), tags=T)
            return
        r = (PY1 - PY0) / 2 - 16
        cx, cy = PX0 + r + 24, (PY0 + PY1) / 2
        ang = 90.0
        for i, v in enumerate(vals):
            if v <= 0:
                continue
            col = PIE[i % len(PIE)]
            ext = -360.0 * v / tot
            if v >= tot * 0.9999:
                cv.create_oval(cx - r, cy - r, cx + r, cy + r, fill=col, outline="", tags=T)
            else:
                cv.create_arc(cx - r, cy - r, cx + r, cy + r, start=ang, extent=ext, style="pieslice",
                              fill=col, outline=LCD, width=1, tags=T)
            if v / tot >= 0.06:
                mid = math.radians(ang + ext / 2)
                lr = 0.78 if hole else 0.62
                cv.create_text(cx + r * lr * math.cos(mid), cy - r * lr * math.sin(mid),text=f"{v / tot * 100:.0f}%", fill="#fffffe", font=(FONT, 8, "bold"), tags=T)
            ang += ext
        if hole:
            cv.create_oval(cx - r * 0.5, cy - r * 0.5, cx + r * 0.5, cy + r * 0.5, fill=LCD, outline="", tags=T)
            cv.create_text(cx, cy - 7, text="Total", fill=INK2, font=(FONT, 8, "bold"), tags=T)
            cv.create_text(cx, cy + 8, text=f"{tot:.6g}", fill=INK, font=(FONT, 11, "bold"), tags=T)
        lx = cx + r + 22
        rows = int((PY1 - PY0 - 54) / 18)
        for i, (lab, v) in enumerate(zip(labs, vals)):
            y = PY0 + 46 + i * 18
            if i >= rows:
                cv.create_text(lx, y, text="…", anchor="w", fill=INK2, font=(FONT, 9), tags=T)
                break
            cv.create_rectangle(lx, y - 5, lx + 10, y + 5, fill=PIE[i % len(PIE)], outline="", tags=T)
            cv.create_text(lx + 16, y, text=f"{lab[:8]}  {v / tot * 100:.1f}%", anchor="w", fill=INK, font=(FONT, 9), tags=T)

    def on_move(self, e):
        if self.tab != "graph" or self.gscreen != "trig" or self.anim or self.gpan:
            return
        cv = self.cv
        cv.delete("gtrace")
        if not self.gfn or not (PX0 <= e.x <= self.PX1 and PY0 <= e.y <= PY1):
            return self.graph_status()
        x = self.gs2x(e.x)
        y = self.gfn(x)
        deg = "°" if self.gxdeg else ""
        if y is None:
            return cv.itemconfig(self.t_gread, text=f"x = {x:.6g}{deg}    y = undefined", fill="#ffffff")
        cv.itemconfig(self.t_gread, text=f"x = {x:.6g}{deg}    y = {y:.6g}", fill="#ffffff")
        sy = self.gy2s(y)
        if PY0 <= sy <= PY1:
            tg = ("cui", "gplot", "gtrace")
            cv.create_line(e.x, PY0, e.x, PY1, fill=INK2, dash=(3, 3), tags=tg)
            cv.create_oval(e.x - 4, sy - 4, e.x + 4, sy + 4, fill=ERR, outline="", tags=tg)

    # Window mode History
    def render_hpanel(self):
        cv = self.cv
        cv.delete("hp")
        self.hp_rows = []
        if not self.has_panel():
            return
        x0, x1, y0, y1 = self.CW - 8, self.TW - 24, 54, 740
        self.round_rect(x0, y0, x1, y1, 16, fill="#15171b", outline="#3a3a3c", width=1, tags="hp")
        cv.create_text(x0 + 20, 82, text="History", anchor="w", fill="#ffffff", font=(FONT, 15, "bold"), tags="hp")
        cv.create_text(x1 - 20, 82, text="✕ Clear", anchor="e", fill=TAB_ON, font=(FONT, 11, "bold"), tags="hp")
        cv.create_line(x0 + 16, 104, x1 - 16, 104, fill="#3a3a3c", tags="hp")
        n = len(self.history)
        if n == 0:
            cv.create_text(x0 + 20, 128, text="There's no history yet.", anchor="w", fill=TAB_OFF, font=(FONT, 11), tags="hp")
            return
        self.hpoff = min(self.hpoff, max(0, n - HPN))
        avail = x1 - x0 - 44
        for j in range(HPN):
            k = n - 1 - self.hpoff - j              
            if k < 0:
                break
            expr, res = self.history[k]
            y = 112 + HPDY * j
            cv.create_text(x1 - 20, y + 16, text=self.fit(expr + " =", self.f_th, avail), anchor="e",fill="#9aa1ad", font=(FONT, 11), tags="hp")
            cv.create_text(x1 - 20, y + 42, text=self.fit(res, self.f_hr, avail), anchor="e",fill="#ffffff", font=(FONT, 16, "bold"), tags="hp")
            cv.create_line(x0 + 16, y + 62, x1 - 16, y + 62, fill="#222226", tags="hp")
            self.hp_rows.append((y, k))
        if self.hpoff > 0:
            cv.create_text(x0 + 20, 112, text="▲", anchor="w", fill=TAB_OFF, font=(FONT, 8), tags="hp")
        if n - self.hpoff - HPN > 0:
            cv.create_text(x0 + 20, y1 - 12, text="▼", anchor="w", fill=TAB_OFF, font=(FONT, 8), tags="hp")

    def scroll_hpanel(self, step):
        top = max(0, len(self.history) - HPN)
        self.hpoff = min(max(self.hpoff + step, 0), top)
        self.render_hpanel()

    def hpanel_click(self, e):
        if e.y < 100:
            if e.x > self.TW - 24 - 110 and e.y > 64:
                self.history.clear()
                self.hpoff = 0
                self.render_hpanel()
            return
        if self.lg:
            return
        for y, k in self.hp_rows:
            if y <= e.y < y + HPDY:
                res = to_expr(self.history[k][1])
                self.expr = res if self.done else self.expr + res
                self.done = False
                self.refresh()
                return

    # History 
    def toggle_history(self):
        if self.lg or self.tab == "conv" or self.wide:
            return
        self.hview = not self.hview
        self.hoff = 0
        for t in (self.t_hist, self.t_expr, self.t_res):
            self.cv.itemconfig(t, state = "hidden"  if self.hview else "normal")
        self.cv.itemconfig(self.t_mode, text = "✕ CLEAR HISTORY" if self.hview else self.mode_text(), fill = ERR if self.hview else INK2)
        self.update_tabs()
        self.render_history()
        if not self.hview and self.tab == "conv":
            self.conv_show()

    def scroll_history(self, step):
        if not self.hview:
            return
        top = max(0, len(self.history) - HROWS)
        self.hoff = min(max(self.hoff + step, 0), top)
        self.render_history()

    def on_wheel(self, e):
        if self.has_panel() and e.x > self.CW - 8:
            up = e.num == 4 or getattr(e, "delta", 0) > 0
            return self.scroll_hpanel(-1 if up else 1)
        if self.tab == "graph":
            if self.gscreen == "trig":
                up = e.num == 4 or getattr(e, "delta", 0) > 0
                return self.graph_zoom(1.2 if up else 1 / 1.2, e.x, e.y)
            return
        if e.num == 4 or getattr(e, "delta" , 0) > 0:
            self.scroll_history(1)
        elif e.num == 5 or getattr(e, "delta", 0) < 0:
            self.scroll_history(-1)

    @staticmethod
    def fit(text, font, maxw):
        while font.measure(text) > maxw and len(text) > 4:
            text = "\u2026" + text[2:]
        return text

    def render_history(self):
        cv = self.cv
        cv.delete("hrow")
        self.hrows = []
        if not self.hview:
            return
        n = len(self.history)
        if n == 0:
            cv.create_text(W/2, 118, text="Please Used Calculator Keys", fill=INK2, font=(FONT, 12) , tags="hrow")
            return
        end  = n - self.hoff
        start = max(0, end - HROWS)
        first = HROWS - (end - start)
        for j, idx in enumerate(range(start, end)):
            y = HY0 + HDY * (first + j)
            expr, res = self.history[idx]
            res = self.fit(res, self.f_hr, 150)
            rt = cv.create_text(W-48, y, text=res, anchor="e", font=self.f_hr, tags="hrow")
            x2 = cv.bbox(rt)[0]-8
            cv.create_text(x2, y, text=self.fit(expr + "=", self.f_hr, x2-54), anchor="e", fill=INK2, font=self.f_hr, tags="hrow")
            self.hrows.append((y,idx))
        if start > 0:
            cv.create_text(40, HY0, text="▲", anchor="w", fill=INK2, font=(FONT, 9), tags="hrow")
        if self.hoff > 0:
            cv.create_text(40, HY0+HDY *(HROWS-1), text="▼", anchor="w", fill=INK2, font=(FONT,9), tags="hrow")

    def history_click(self, e):
        if e.y < 80:
            if e.x < 280:
                self.history.clear()
                self.render_history()
            return
        if not self.hrows:
            return
        y, idx = min(self.hrows, key=lambda t: abs(t[0] - e.y))
        if abs(y - e.y) <= HDY / 2 + 2:
            res = to_expr(self.history[idx][1])
            self.toggle_history()
            self.expr = res if self.done else self.expr + res
            self.done = False
            self.refresh()

    def hint_text(self):
        s = to_round(self.expr)
        stack = []
        for i, ch in enumerate(s):
            if ch == "(":
                m = re.search(r"[a-z]+$", s[:i])
                stack.append([m.group() if m else "", 0])
            elif ch == ")":
                if stack:
                    stack.pop()
            elif ch == "," and stack:
                stack[-1][1] += 1
        for name, idx in reversed(stack):
            if name in HINTS:
                args, desc, ex = HINTS[name]
                parts = [f"‹{a}›" if j == idx else a for j, a in enumerate(args)]
                line = f"{name}({', '.join(parts)}) → {desc}"
                full = f"{line}   ex: {ex}"
                return full if len(full) <= 58 else line
        return ""


    def show_expr(self):
        s = to_round(self.expr) or "0"
        n = len(s)
        size = 26 if n <= 14 else 20 if n <= 20 else 16 if n <= 27 else 13
        if n > 36:
            s = "…" + s[-35:]
        self.cv.itemconfig(self.t_expr, text=s, font=(FONT, size))


    def refresh(self):
        if self.done:
            return
        self.cv.itemconfig(self.t_hist, text=self.hint_text(), font=(FONT, 10))
        self.show_expr()
        text = ""
        if self.expr and re.search(r"[+−×÷^%!(\[{]|sin|cos|tan|ln|log|√|π|e|Ans", self.expr):
            try:
                text = "= " + fmt(self.evaluate(self.expr))
            except Exception:
                pass
        self.cv.coords(self.t_res, self.CW - 44, 158)
        self.cv.itemconfig(self.t_res, text=text, fill=INK2,
                           font=(FONT, 16 if len(text) < 20 else 12, "bold"))

if __name__ == "__main__":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    Calculator(root)
    root.mainloop()
