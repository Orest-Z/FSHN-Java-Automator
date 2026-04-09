"""
=============================================================
  Java Lab Report Builder  —  by Claude for Orest Zogju
=============================================================
Generates a perfectly formatted .docx lab report from:
  • Exercise number & description
  • Java source code
  • Screenshots — multiple per exercise (auto-resized to fit page)

Usage:
    python lab_report_builder.py

Requirements:
    pip install python-docx Pillow

Streamlit UI (optional):
    pip install streamlit
    streamlit run lab_report_builder.py
=============================================================
"""

import os
import io
import sys
import argparse

# ─── TEMPLATE  ── change these to match your professor's requirements ──────────
TEMPLATE = {
    # Page
    "page_width_inches":  8.27,   # A4
    "page_height_inches": 11.69,
    "margin_inches":      1.0,

    # Header (top of first page)
    "header_font":        "Times New Roman",
    "header_size_pt":     14,
    "header_bold":        True,

    # Exercise title  "Ushtrimi X:"
    "exercise_title_font": "Times New Roman",
    "exercise_title_size": 12,
    "exercise_title_bold": True,

    # Description text
    "description_font":   "Times New Roman",
    "description_size":   11,

    # Code block
    "code_font":          "Courier New",
    "code_size":          8,           # pt
    "code_bg_color":      "EEEEEE",    # light grey shading

    # Image
    "image_max_width_inches": 5.5,     # auto-scales height to keep ratio

    # Spacing
    "space_after_code_pt":  12,
    "space_after_image_pt": 12,

    # Page break between exercises
    "page_break_between": True,
}
# ──────────────────────────────────────────────────────────────────────────────


# ╔══════════════════════════════════════════════════════════════════╗
# ║                         CORE BUILDER                            ║
# ╚══════════════════════════════════════════════════════════════════╝

def _inches(val):
    from docx.shared import Inches
    return Inches(val)

def _pt(val):
    from docx.shared import Pt
    return Pt(val)

def _rgb_from_hex(hex_str):
    from docx.shared import RGBColor
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    return RGBColor(r, g, b)


def _set_shading(paragraph, hex_color):
    """Apply background shading to a paragraph (code block style)."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


def _add_page_break(doc):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    p = doc.add_paragraph()
    run = p.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)


def _resize_image(image_bytes_or_path, max_width_inches):
    """
    Return (width_inches, height_inches) so the image fits within max_width_inches
    while preserving aspect ratio.
    """
    from PIL import Image
    if isinstance(image_bytes_or_path, (bytes, bytearray)):
        img = Image.open(io.BytesIO(image_bytes_or_path))
    else:
        img = Image.open(image_bytes_or_path)

    w_px, h_px = img.size
    dpi = img.info.get("dpi", (96, 96))
    dpi_x = dpi[0] if isinstance(dpi, tuple) else 96

    # Convert pixels → inches using the image's own DPI
    w_in = w_px / dpi_x
    h_in = h_px / dpi_x

    if w_in > max_width_inches:
        scale = max_width_inches / w_in
        w_in = max_width_inches
        h_in = h_in * scale

    return w_in, h_in


def create_document(
    doc_title: str,
    author_name: str,
    exercises: list,          # list of dicts — see add_exercise()
    output_path: str = "lab_report.docx",
    template: dict = None,
):
    """
    Build and save the .docx report.

    exercises: list of dicts, each with keys:
        number        (str)        – "1", "2", …
        description   (str)        – plain text description
        code          (str)        – Java source code
        images        (list)       – list of dicts, each:
                                       {"path": str|None, "bytes": bytes|None}
                                     (empty list = no images)
    """
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    T = template or TEMPLATE

    doc = Document()

    # ── Page setup ──────────────────────────────────────────────────
    section = doc.sections[0]
    section.page_width  = _inches(T["page_width_inches"])
    section.page_height = _inches(T["page_height_inches"])
    m = _inches(T["margin_inches"])
    section.top_margin    = m
    section.bottom_margin = m
    section.left_margin   = m
    section.right_margin  = m

    # ── Remove default paragraph spacing ────────────────────────────
    style = doc.styles["Normal"]
    style.font.name = T["description_font"]
    style.font.size = _pt(T["description_size"])
    pf = style.paragraph_format
    pf.space_before = _pt(0)
    pf.space_after  = _pt(4)

    # ── Document header line ─────────────────────────────────────────
    header_para = doc.add_paragraph()
    header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_title = header_para.add_run(doc_title)
    run_title.font.name  = T["header_font"]
    run_title.font.size  = _pt(T["header_size_pt"])
    run_title.font.bold  = T["header_bold"]

    # Tab stop to push author name to the right
    run_tab = header_para.add_run("\t")
    run_author = header_para.add_run(author_name)
    run_author.font.name  = T["header_font"]
    run_author.font.size  = _pt(T["header_size_pt"])
    run_author.font.bold  = T["header_bold"]

    # Add tab stop at right margin
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    pPr = header_para._p.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    content_width_twips = int(
        (T["page_width_inches"] - 2 * T["margin_inches"]) * 1440
    )
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(content_width_twips))
    tabs.append(tab)
    pPr.append(tabs)

    # Horizontal rule below header
    border_para = doc.add_paragraph()
    bPr = border_para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    bPr.append(pBdr)
    border_para.paragraph_format.space_after = _pt(10)

    # ── Exercises ────────────────────────────────────────────────────
    for idx, ex in enumerate(exercises):
        _add_exercise(doc, ex, T)
        if T["page_break_between"] and idx < len(exercises) - 1:
            _add_page_break(doc)

    doc.save(output_path)
    return output_path


def _add_exercise(doc, ex: dict, T: dict):
    """Append one exercise block to the document."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    number      = ex.get("number", "?")
    description = ex.get("description", "")
    code        = ex.get("code", "")
    images      = ex.get("images", [])   # list of {"path": …, "bytes": …}

    # ── Exercise title ───────────────────────────────────────────────
    title_para = doc.add_paragraph()
    run = title_para.add_run(f"Ushtrimi {number}:")
    run.font.name  = T["exercise_title_font"]
    run.font.size  = _pt(T["exercise_title_size"])
    run.font.bold  = T["exercise_title_bold"]
    title_para.paragraph_format.space_after  = _pt(6)
    title_para.paragraph_format.space_before = _pt(10)

    # ── Description ──────────────────────────────────────────────────
    if description.strip():
        for line in description.strip().split("\n"):
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.font.name = T["description_font"]
            run.font.size = _pt(T["description_size"])
            p.paragraph_format.space_after = _pt(4)

    doc.add_paragraph()  # small gap before code

    # ── Code block ───────────────────────────────────────────────────
    if code.strip():
        for line in code.split("\n"):
            p = doc.add_paragraph()
            _set_shading(p, T["code_bg_color"])
            p.paragraph_format.space_before = _pt(0)
            p.paragraph_format.space_after  = _pt(0)
            p.paragraph_format.left_indent  = _inches(0.1)
            run = p.add_run(line if line else " ")  # keep empty lines
            run.font.name = T["code_font"]
            run.font.size = _pt(T["code_size"])

        # Space after code block
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_after = _pt(T["space_after_code_pt"])

    # ── Screenshots (one or more) ─────────────────────────────────────
    for img in images:
        img_path  = img.get("path", None)
        img_bytes = img.get("bytes", None)
        if not img_path and not img_bytes:
            continue
        try:
            source = img_bytes if img_bytes else img_path
            w_in, h_in = _resize_image(source, T["image_max_width_inches"])

            p = doc.add_paragraph()
            p.alignment = 0  # left
            run = p.add_run()

            if img_bytes:
                run.add_picture(io.BytesIO(img_bytes), width=_inches(w_in))
            else:
                run.add_picture(img_path, width=_inches(w_in))

            p.paragraph_format.space_after = _pt(T["space_after_image_pt"])
        except Exception as e:
            err = doc.add_paragraph(f"[Image error: {e}]")
            err.runs[0].font.name = "Arial"



# ╔══════════════════════════════════════════════════════════════════╗
# ║                      STREAMLIT  UI                              ║
# ╚══════════════════════════════════════════════════════════════════╝

def run_streamlit():
    import streamlit as st

    st.set_page_config(page_title="Java Lab Report Builder", page_icon="📄", layout="wide")

    st.title("📄 Java Lab Report Builder")
    st.markdown("Fill in the fields below, click **Add Exercise**, then **Generate .docx**.")

    # ── Sidebar: document-level settings ────────────────────────────
    with st.sidebar:
        st.header("⚙️ Document Settings")
        doc_title   = st.text_input("Assignment Title",  value="Laborator Java 2")
        author_name = st.text_input("Name / Group",      value="Orest Zogju GR B1.2")
        output_name = st.text_input("Output filename",   value="lab_report.docx")

        st.divider()
        st.subheader("Template Overrides")
        img_width = st.slider("Max image width (inches)", 3.0, 8.0, float(TEMPLATE["image_max_width_inches"]), 0.5)
        code_size = st.slider("Code font size (pt)",      6,   14,  int(TEMPLATE["code_size"]))
        desc_size = st.slider("Description font size (pt)", 8, 14, int(TEMPLATE["description_size"]))

    # ── Session state: accumulated exercises ────────────────────────
    if "exercises" not in st.session_state:
        st.session_state.exercises = []

    # ── Input form ──────────────────────────────────────────────────
    st.subheader("➕ Add an Exercise")
    col1, col2 = st.columns([1, 2])

    with col1:
        ex_number = st.text_input("Exercise Number", value=str(len(st.session_state.exercises) + 1))
        screenshots = st.file_uploader(
            "Screenshots (PNG/JPG) — select multiple",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True,
        )

    with col2:
        description = st.text_area("Exercise Description", height=120,
                                   placeholder="Ndertoni nje dritare me titull…")

    code = st.text_area("Java Code", height=300,
                        placeholder="import javax.swing.*;\n\npublic class Ushtrimi1 {\n    …\n}")

    if st.button("➕ Add Exercise to Document"):
        if not ex_number.strip():
            st.error("Please enter an exercise number.")
        else:
            images = [{"path": None, "bytes": f.read()} for f in screenshots] if screenshots else []
            entry = {
                "number":      ex_number.strip(),
                "description": description,
                "code":        code,
                "images":      images,
            }
            st.session_state.exercises.append(entry)
            img_count = len(images)
            st.success(f"✅ Ushtrimi {ex_number} added! ({len(st.session_state.exercises)} total, {img_count} image(s))")

    # ── List of added exercises ─────────────────────────────────────
    if st.session_state.exercises:
        st.divider()
        st.subheader(f"📋 Exercises in queue ({len(st.session_state.exercises)})")
        for i, ex in enumerate(st.session_state.exercises):
            col_a, col_b = st.columns([5, 1])
            with col_a:
                img_count = len(ex.get("images", []))
                img_label = f"🖼️ ×{img_count}" if img_count else "—"
                st.write(f"**{i+1}.** Ushtrimi {ex['number']}  |  "
                         f"Code lines: {len(ex['code'].splitlines())}  |  Images: {img_label}")
            with col_b:
                if st.button("🗑️", key=f"del_{i}"):
                    st.session_state.exercises.pop(i)
                    st.rerun()

        if st.button("🗑️ Clear All Exercises"):
            st.session_state.exercises = []
            st.rerun()

        st.divider()

        # ── Generate ───────────────────────────────────────────────
        if st.button("📥 Generate .docx", type="primary"):
            custom_template = {**TEMPLATE,
                               "image_max_width_inches": img_width,
                               "code_size":  code_size,
                               "description_size": desc_size}
            buf = io.BytesIO()
            # Save to a temp file, then read back into buffer
            import tempfile
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                tmp_path = tmp.name

            create_document(
                doc_title   = doc_title,
                author_name = author_name,
                exercises   = st.session_state.exercises,
                output_path = tmp_path,
                template    = custom_template,
            )
            with open(tmp_path, "rb") as f:
                buf.write(f.read())
            os.unlink(tmp_path)

            buf.seek(0)
            st.download_button(
                label    = "⬇️  Download " + output_name,
                data     = buf,
                file_name= output_name,
                mime     = "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            st.success("Document ready! Click the button above to download.")


# ╔══════════════════════════════════════════════════════════════════╗
# ║                        CLI  INTERFACE                           ║
# ╚══════════════════════════════════════════════════════════════════╝

def run_cli():
    """Simple interactive CLI for building a report without Streamlit."""
    print("\n╔══════════════════════════════════════╗")
    print("║   Java Lab Report Builder  (CLI)    ║")
    print("╚══════════════════════════════════════╝\n")

    doc_title   = input("Assignment title  [Laborator Java 2]: ").strip() or "Laborator Java 2"
    author_name = input("Your name/group   [Orest Zogju GR B1.2]: ").strip() or "Orest Zogju GR B1.2"
    output_path = input("Output file name  [lab_report.docx]: ").strip() or "lab_report.docx"

    exercises = []
    print("\nFor each exercise, paste the details. Type END on its own line to finish multi-line input.")

    while True:
        print(f"\n─── Exercise {len(exercises)+1} ───────────────────────")
        num = input("  Exercise number (or DONE to finish): ").strip()
        if num.upper() == "DONE":
            break

        print("  Description (END to finish):")
        desc_lines = []
        while True:
            line = input()
            if line.strip() == "END":
                break
            desc_lines.append(line)

        print("  Java code (END to finish):")
        code_lines = []
        while True:
            line = input()
            if line.strip() == "END":
                break
            code_lines.append(line)

        print("  Screenshot paths — one per line (blank line to finish):")
        images = []
        while True:
            p = input("    path (or blank to stop): ").strip()
            if not p:
                break
            if os.path.isfile(p):
                images.append({"path": p, "bytes": None})
            else:
                print(f"    ⚠️  File not found, skipping: {p}")

        exercises.append({
            "number":      num,
            "description": "\n".join(desc_lines),
            "code":        "\n".join(code_lines),
            "images":      images,
        })
        print(f"  ✓ Ushtrimi {num} added ({len(images)} image(s)).")

    if not exercises:
        print("No exercises added. Exiting.")
        return

    path = create_document(doc_title, author_name, exercises, output_path)
    print(f"\n✅ Report saved to: {os.path.abspath(path)}\n")


# ╔══════════════════════════════════════════════════════════════════╗
# ║                  QUICK-DEMO  (no arguments)                     ║
# ╚══════════════════════════════════════════════════════════════════╝

def run_demo(output_path="lab_report_demo.docx"):
    """Generate a sample document with two exercises for testing."""
    exercises = [
        {
            "number":      "1",
            "description": (
                "Ndertoni nje dritare me titull Sistemi i Autorizimit. "
                "Dritarja duhet te kete permasa fikse, te shfaqet ne qender te ekranit "
                "dhe te kete nje label me tekstin \"Ju lutem identifikohuni\".\n"
                "Sigurohuni qe dritarja te mbyllet plotesisht kur shtypet butoni \"X\"."
            ),
            "code": (
                "import javax.swing.*;\n"
                "import java.awt.*;\n\n"
                "public class Ushtrimi1 extends JFrame {\n"
                "    Ushtrimi1() {\n"
                "        JFrame frame = new JFrame(\"Sistemi i Autorizimit\");\n"
                "        frame.setSize(400, 400);\n"
                "        frame.setLayout(new FlowLayout());\n\n"
                "        JLabel l1 = new JLabel(\"Ju lutem identifikohuni\");\n"
                "        JButton btn1 = new JButton(\"X\");\n"
                "        btn1.setBackground(Color.RED);\n"
                "        btn1.addActionListener(e -> System.exit(0));\n\n"
                "        frame.add(l1);\n"
                "        frame.add(btn1);\n"
                "        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);\n"
                "        frame.setLocationRelativeTo(null);\n"
                "        frame.setVisible(true);\n"
                "    }\n"
                "}"
            ),
            "images":      [],
        },
        {
            "number":      "2",
            "description": (
                "Ndertoni nje program qe perdor per te pyetur: "
                "\"A deshironi te mbyllni seancen?\". "
                "Nese klikohet Po, te shfaqet nje mesazh informues: \"Seanca u mbyll me sukses\"."
            ),
            "code": (
                "import javax.swing.*;\n"
                "import static javax.swing.JOptionPane.YES_NO_OPTION;\n\n"
                "public class Ushtrimi2 extends JFrame {\n"
                "    Ushtrimi2() {\n"
                "        int zgjedhja = JOptionPane.showConfirmDialog(\n"
                "            null, \"A deshironi te mbyllni seancen?\",\n"
                "            \"Zgjidh nje opsion\", YES_NO_OPTION);\n\n"
                "        if (zgjedhja == JOptionPane.YES_OPTION) {\n"
                "            JOptionPane.showMessageDialog(null, \"Seanca u mbyll me sukses\");\n"
                "            System.exit(0);\n"
                "        } else {\n"
                "            JOptionPane.showMessageDialog(null, \"Veprimi u refuzua\");\n"
                "        }\n"
                "    }\n"
                "}"
            ),
            "images":      [],
        },
    ]

    path = create_document(
        doc_title   = "Laborator Java 2",
        author_name = "Orest Zogju GR B1.2",
        exercises   = exercises,
        output_path = output_path,
    )
    print(f"Demo document created: {os.path.abspath(path)}")
    return path


# ╔══════════════════════════════════════════════════════════════════╗
# ║                          ENTRY POINT                            ║
# ╚══════════════════════════════════════════════════════════════════╝

if __name__ == "__main__":
    # Detect if launched by Streamlit
    if "streamlit" in sys.modules or any("streamlit" in a for a in sys.argv):
        run_streamlit()
    else:
        parser = argparse.ArgumentParser(description="Java Lab Report Builder")
        parser.add_argument("--demo", action="store_true", help="Generate a demo document")
        parser.add_argument("--cli",  action="store_true", help="Run interactive CLI")
        args = parser.parse_args()

        if args.demo:
            run_demo()
        elif args.cli:
            run_cli()
        else:
            # Default: try Streamlit, fall back to CLI
            try:
                import streamlit
                print("Tip: run with 'streamlit run lab_report_builder.py' for the visual UI.")
                print("Falling back to CLI mode...\n")
                run_cli()
            except ImportError:
                run_cli()

# If imported as a module by Streamlit (streamlit run lab_report_builder.py),
# this block runs automatically:
try:
    import streamlit as _st
    if _st.runtime.exists():
        run_streamlit()
except Exception:
    pass
