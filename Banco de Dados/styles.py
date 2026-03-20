# styles.py
COLORS = {
    "primary": "#0E3A66",   # navy
    "secondary": "#4BB6E5", # sky
    "accent": "#1F5E8F",
    "bg": "#F5FAFF",
    "border": "#D3E3F1",
    "text": "#0D2740",
    "muted": "#3B5D7A"
}

APP_QSS = f"""
QMainWindow {{
    background: {COLORS['bg']};
    color: {COLORS['text']};
    font-family: "Segoe UI", Tahoma, sans-serif;
}}

#dropArea {{
    background: white;
    border: 2px dashed {COLORS['border']};
    border-radius: 8px;
    color: {COLORS['muted']};
    padding: 12px;
    margin: 6px;
}}

QPushButton {{
    background-color: {COLORS['primary']};
    color: white;
    border-radius: 6px;
    padding: 8px 14px;
    font-weight: 600;
}}

QPushButton:hover {{
    background-color: {COLORS['accent']};
}}

QPushButton:disabled {{
    background-color: #A7C3D9;
    color: #F5FAFF;
}}

QProgressBar {{
    border: 1px solid {COLORS['border']};
    border-radius: 6px;
    text-align: center;
    height: 18px;
}}

QProgressBar::chunk {{
    background-color: {COLORS['secondary']};
    border-radius: 6px;
}}

QLabel {{
    color: {COLORS['text']};
}}
"""

# === Ajuste visual para QMessageBox ===
# === Ajuste visual para QMessageBox ===
APP_QSS += f"""
QMessageBox {{
    background-color: #1E1E1E;
    color: #EAEAEA;
    font-family: "Segoe UI", Tahoma, sans-serif;
    border: 1px solid #333;
}}

QMessageBox QLabel {{
    color: #EAEAEA;
    font-size: 10.5pt;
}}

QMessageBox QPushButton {{
    background-color: #3A86FF;
    color: white;
    border-radius: 6px;
    padding: 6px 12px;
    font-weight: 600;
}}

QMessageBox QPushButton:hover {{
    background-color: #005BCE;
}}
"""
