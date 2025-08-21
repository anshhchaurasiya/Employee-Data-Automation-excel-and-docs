# 📘 python-docx Complete Guide

---

## 1. Creating / Opening a Word File

```python
from docx import Document

# Create a new Word document
doc = Document()

# Open an existing Word document
doc = Document("file.docx")
```

---

## 2. Document Structure in python-docx

Think of a Word file like this:

```
Document
 ├── Paragraphs
 │     └── Runs (formatted text)
 ├── Tables
 │     ├── Rows
 │     └── Cells
 └── Sections (page setup)
```

👉 **Hierarchy is very important**:

* **Paragraphs** = lines of text.
* **Runs** = styled portions inside a paragraph.
* **Tables** = contain rows → cells → paragraphs.
* **Sections** = page layout (margins, orientation, headers/footers).

---

## 3. Working with Paragraphs

### Read all paragraphs

```python
for p in doc.paragraphs:
    print(p.text)
```

### Add a new paragraph

```python
doc.add_paragraph("This is a new line.")
```

### Style a paragraph

```python
para = doc.add_paragraph("Heading Example")
para.style = "Heading 1"  # built-in Word style
```

📌 Common built-in styles: `"Normal"`, `"Heading 1"`, `"Heading 2"`, `"Title"`, `"Subtitle"`

---

## 4. Working with Runs (inline formatting)

```python
from docx.shared import Pt, RGBColor

para = doc.add_paragraph()
run = para.add_run("Hello World")

# Formatting
run.font.bold = True
run.font.italic = True
run.font.underline = True
run.font.size = Pt(14)
run.font.color.rgb = RGBColor(255, 0, 0)   # Red color
run.font.name = "Calibri"
```

⚡ Use runs when you want **different styles in the same line**.

---

## 5. Line Breaks and Page Breaks

```python
para.add_run("\nNew line inside same paragraph")   # Line break
doc.add_page_break()   # Forces next page
```

---

## 6. Adding Lists

```python
# Bullet List
doc.add_paragraph("Item 1", style="List Bullet")
doc.add_paragraph("Item 2", style="List Bullet")

# Numbered List
doc.add_paragraph("Step 1", style="List Number")
doc.add_paragraph("Step 2", style="List Number")
```

---

## 7. Adding Tables

```python
table = doc.add_table(rows=2, cols=2)
table.style = "Table Grid"

# Add data
cell = table.cell(0, 0)
cell.text = "Row 1, Col 1"

table.cell(0, 1).text = "Row 1, Col 2"
table.cell(1, 0).text = "Row 2, Col 1"
table.cell(1, 1).text = "Row 2, Col 2"
```

---

## 8. Adding Images

```python
from docx.shared import Inches

doc.add_picture("image.png", width=Inches(2), height=Inches(2))
```

⚡ Supports `.png`, `.jpg`, `.bmp`, `.gif`

---

## 9. Headers and Footers

```python
section = doc.sections[0]

# Header
header = section.header
header_para = header.paragraphs[0]
header_para.text = "My Document Header"

# Footer
footer = section.footer
footer_para = footer.paragraphs[0]
footer_para.text = "Page Footer Example"
```

---

## 10. Page Setup (Margins, Orientation)

```python
from docx.shared import Inches
from docx.enum.section import WD_ORIENT

section = doc.sections[0]

# Margins
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)
section.left_margin = Inches(1)
section.right_margin = Inches(1)

# Orientation (Portrait / Landscape)
section.orientation = WD_ORIENT.LANDSCAPE
```

---

## 11. Adding Hyperlinks (custom trick 🚀)

`python-docx` doesn’t directly support hyperlinks, but workaround:

```python
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

para = doc.add_paragraph()
part = doc.part
r_id = part.relate_to(
    "https://www.google.com",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
    is_external=True
)

hyperlink = OxmlElement("w:hyperlink")
hyperlink.set(qn("r:id"), r_id)

new_run = OxmlElement("w:r")
rPr = OxmlElement("w:rPr")

new_run.append(rPr)
text = OxmlElement("w:t")
text.text = "Google"
new_run.append(text)

hyperlink.append(new_run)
para._p.append(hyperlink)
```

---

## 12. Saving the File

```python
doc.save("final_document.docx")
```

---

# ⚡ Quick Recap

* **Paragraphs** = full lines of text.
* **Runs** = styled portions inside paragraphs.
* **Styles** = Word’s built-in formatting (headings, lists, etc.).
* **Tables/Images** = supported easily.
* **Headers/Footers & Page Setup** = done via `sections`.

---

👉 With this knowledge, you can now **generate reports, resumes, invoices, or formatted notes automatically** using Python.
