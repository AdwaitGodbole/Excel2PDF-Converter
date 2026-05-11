
### Author: Danny Cedrone and Adwait Godbole

import os, getpass
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.lib.pagesizes import landscape, elevenSeventeen
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# =========================================================
# REGISTER TIMES NEW ROMAN FONT
# =========================================================
pdfmetrics.registerFont(TTFont("TimesNewRoman", "C:/Windows/Fonts/times.ttf"))
pdfmetrics.registerFont(TTFont("TimesNewRoman-Bold", "C:/Windows/Fonts/timesbd.ttf"))

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._codes = []
    def showPage(self):
        self._codes.append({'code': self._code, 'stack': self._codeStack})
        self._startPage()
    def save(self):
        """add page info to each page (page x of y)"""
        # reset page counter
        self._pageNumber = 0
        width, _ = landscape(elevenSeventeen)
        for code in self._codes:
            # recall saved page
            self._code = code['code']
            self._codeStack = code['stack']
            self.setFont("TimesNewRoman", 11)
            timestamp = datetime.now().strftime("%m-%d-%Y %H:%M:%S")
            self.drawString(20, 15, timestamp)
            self.drawCentredString(
                width / 2,
                15,
                "Visiopharm 2025"
            )
            self.drawRightString(width - 20, 15,
                "Page %(this)i of %(total)i" % {
                   'this': self._pageNumber+1,
                   'total': len(self._codes),
                }
            )
            canvas.Canvas.showPage(self)
        self._doc.SaveToFile(self._filename, self)

# =========================================================
# HEADER / FOOTER FUNCTION
# =========================================================

def draw_header(canvas, doc, study_number):
    width, height = landscape(elevenSeventeen)
    page_num = canvas.getPageNumber()

    canvas.saveState()
    canvas.setFont("TimesNewRoman", 11)

    # Header
    canvas.drawString(20, height - 20, "CBSET, Inc.")
    canvas.drawCentredString(width / 2, height - 20, f"{study_number} Morphometry Data")
    canvas.drawRightString(width - 20, height - 20, "Confidential")

    canvas.restoreState()

# =========================================================
# TABLE BUILDER
# =========================================================
def build_autoscaling_table(df, max_width):
    df = df.fillna("")

    styles = getSampleStyleSheet()

    body_style = ParagraphStyle(
        name="BodyStyle",
        parent=styles["BodyText"],
        fontName="TimesNewRoman",
        fontSize=11,
        leading=13,
        wordWrap="CJK"
    )

    header_style = ParagraphStyle(
        name="HeaderStyle",
        parent=styles["Heading5"],
        fontName="TimesNewRoman-Bold",
        fontSize=11,
        leading=13,
        alignment=1,  # center
        wordWrap="CJK"
    )

    data = []

    # Header row
    data.append([Paragraph(f"<b>{col}</b>", header_style) for col in df.columns])

    # Body rows
    for _, row in df.iterrows():
        data.append([Paragraph(str(cell), body_style) for cell in row])

    num_cols = len(df.columns)
    col_widths = [max_width / num_cols] * num_cols

    table = Table(data, colWidths=col_widths, repeatRows=1)

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))

    # Auto-scale columns if necessary
    w, _ = table.wrap(0, 0)
    if w > max_width:
        scale = max_width / w
        table._argW = [cw * scale for cw in table._argW]

    return table

# =========================================================
# FILE LOADING WITH ENCODING FALLBACK
# =========================================================
def load_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".tsv":
        try:
            return pd.read_csv(path, sep="\t", encoding="utf-8")
        except UnicodeDecodeError:
            return pd.read_csv(path, sep="\t", encoding="latin-1")
    elif ext in [".xls", ".xlsx"]:
        return pd.read_excel(path)
    else:
        raise ValueError("Unsupported file type")

# =========================================================
# MANDATORY INPUT POPUP
# =========================================================
def ask_mandatory(prompt, title):
    value = None
    while not value:
        win = tk.Toplevel()
        win.title(title)
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text=prompt).pack(padx=10, pady=5)
        entry = tk.Entry(win)
        entry.pack(padx=10, pady=5)
        entry.focus_set()

        def submit():
            nonlocal value
            val = entry.get().strip()
            if not val:
                messagebox.showwarning("Required", "This field is required.")
                return
            value = val
            win.destroy()

        tk.Button(win, text="OK", command=submit).pack(pady=5)
        win.wait_window()
    return value

# =========================================================
# CONVERT FILE TO PDF
# =========================================================
def convert_file_to_pdf(path, study_number):
    df = load_file(path)
    pdf_path = os.path.splitext(path)[0] + ".pdf"

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(elevenSeventeen),
        leftMargin=20,
        rightMargin=20,
        topMargin=30,
        bottomMargin=30
    )

    max_width = landscape(elevenSeventeen)[0] - 40
    table = build_autoscaling_table(df, max_width)

    doc.build(
        [table],
        onFirstPage=lambda c, d: draw_header(c, d, study_number),
        onLaterPages=lambda c, d: draw_header(c, d, study_number),
        canvasmaker=NumberedCanvas
    )

    return pdf_path

# =========================================================
# MAIN FUNCTION
# =========================================================
def main():
    root = tk.Tk()
    root.withdraw()

    study_number = ask_mandatory("Enter Study Number:", "Study Number Required")

    files = filedialog.askopenfilenames(
        title="Select TSV or Excel Files",
        filetypes=[("Supported Files", "*.tsv *.xls *.xlsx")]
    )

    if not files:
        messagebox.showinfo("Cancelled", "No files selected.")
        return

    for f in files:
        convert_file_to_pdf(f, study_number)

    messagebox.showinfo("Done", "PDF conversion completed successfully.")

# =========================================================
if __name__ == "__main__":
    main()
