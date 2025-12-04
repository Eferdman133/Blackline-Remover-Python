#!/usr/bin/env python3
"""
Blackline Remover by Emil Ferdman - PyQt6 GUI

Features:
- Delete red text and green strikethrough text
- Remove formatting from remaining green and blue text
- Detect simulated strikethroughs (drawn green lines)
- Remove colored graphical shapes and empty text boxes
- Drag & drop .docx files, or browse and select multiple files
- Adjustable color tolerances and target RGB colors (via color pickers)
- Process to cleaned .docx OR to PDF (requires docx2pdf)
- Prompt to open processed file / output folder after completion
"""

import sys, os, zipfile, tempfile, re, subprocess
from pathlib import Path

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QLineEdit, QTextEdit, QProgressBar, QGroupBox, QSlider, QCheckBox,
    QColorDialog, QMessageBox, QFrame
)

# Try to use lxml for better namespace handling, fall back to ElementTree
try:
    from lxml import etree as ET
    USING_LXML = True
except ImportError:
    from xml.etree import ElementTree as ET
    USING_LXML = False

# Optional PDF support
try:
    from docx2pdf import convert as _docx2pdf_convert_orig
    DOCX2PDF_AVAILABLE = True
    
    def docx2pdf_convert(input_path, output_path):
        """Wrapper to handle stdout issues in GUI applications."""
        import io
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        try:
            # Redirect stdout/stderr to prevent tqdm errors in GUI
            if sys.stdout is None:
                sys.stdout = io.StringIO()
            if sys.stderr is None:
                sys.stderr = io.StringIO()
            _docx2pdf_convert_orig(input_path, output_path)
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            
except ImportError:
    DOCX2PDF_AVAILABLE = False

# XML namespaces used in OOXML
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'v': 'urn:schemas-microsoft-com:vml',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'o': 'urn:schemas-microsoft-com:office:office',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
    'cx1': 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex',
    'cx2': 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex',
    'aink': 'http://schemas.microsoft.com/office/drawing/2016/ink',
    'am3d': 'http://schemas.microsoft.com/office/drawing/2017/model3d',
}

# Register namespaces for ElementTree (standard library fallback)
if not USING_LXML:
    for prefix, uri in NAMESPACES.items():
        ET.register_namespace(prefix, uri)


# ---------- Color helpers ----------

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 6:
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return None


def color_distance(c1, c2):
    if c1 is None or c2 is None:
        return float('inf')
    return sum((a - b) ** 2 for a, b in zip(c1, c2)) ** 0.5


def is_color_match(hex_color, target_hex, tolerance):
    """Check if a color matches the target within tolerance (0-100)."""
    if not hex_color:
        return False
    hex_color = hex_color.lstrip('#').upper()
    target_hex = target_hex.lstrip('#').upper()
    c1 = hex_to_rgb(hex_color)
    c2 = hex_to_rgb(target_hex)
    if c1 is None or c2 is None:
        return False
    # Tolerance 0-100 scaled to RGB max distance ~441
    max_dist = tolerance * 4.41
    return color_distance(c1, c2) <= max_dist


# ---------- Core processor ----------

class WordColorProcessor:
    """Handles the actual document processing."""

    def __init__(self, red_tolerance=10, green_tolerance=10, blue_tolerance=10,
                 check_simulated_lines=True,
                 red_color='#FF0000', green_color='#008000', blue_color='#0000FF'):
        self.red_tolerance = red_tolerance
        self.green_tolerance = green_tolerance
        self.blue_tolerance = blue_tolerance
        self.check_simulated_lines = check_simulated_lines

        # Target colors (store as 'RRGGBB')
        self.RED = red_color.lstrip('#').upper()
        self.GREEN = green_color.lstrip('#').upper()
        self.BLUE = blue_color.lstrip('#').upper()

        self.stats = {
            'red_deleted': 0,
            'green_strike_deleted': 0,
            'green_cleaned': 0,
            'blue_cleaned': 0,
            'textboxes_removed': 0,
            'colored_shapes_removed': 0,
        }

        self.parent_map = {}
        self.paragraph_shapes = {}
        self.greenshape_striked_runs = set()

    def process_document(self, input_path, output_path,
                         progress_callback=None, progress_value_callback=None):
        """Process the Word document."""
        self.stats = {k: 0 for k in self.stats}

        with tempfile.TemporaryDirectory() as temp_dir:
            extract_dir = os.path.join(temp_dir, 'extracted')

            if progress_callback:
                progress_callback("Extracting document...")
            if progress_value_callback:
                progress_value_callback(10)
            with zipfile.ZipFile(input_path, 'r') as z:
                z.extractall(extract_dir)

            doc_path = os.path.join(extract_dir, 'word', 'document.xml')

            if progress_callback:
                progress_callback("Processing document content...")
            if progress_value_callback:
                progress_value_callback(20)

            if os.path.exists(doc_path):
                self.process_xml_file(doc_path, progress_callback, progress_value_callback)

            if progress_callback:
                progress_callback("Saving processed document...")
            if progress_value_callback:
                progress_value_callback(95)

            self.create_docx(extract_dir, output_path)

            if progress_value_callback:
                progress_value_callback(100)

        return self.stats

    def process_xml_file(self, xml_path, progress_callback=None, progress_value_callback=None):
        """Process an XML file to handle colored text."""
        # Parse with lxml or ElementTree
        if USING_LXML:
            parser = ET.XMLParser(remove_blank_text=False, strip_cdata=False)
            tree = ET.parse(xml_path, parser)
        else:
            tree = ET.parse(xml_path)
        
        root = tree.getroot()

        self.parent_map = {c: p for p in root.iter() for c in p}

        if progress_value_callback:
            progress_value_callback(30)

        if self.check_simulated_lines:
            if progress_callback:
                progress_callback("Scanning for drawn line shapes (possible simulated strikethroughs/underlines)...")
            self.paragraph_shapes = self.build_paragraph_shape_map(root)
            self.detect_green_shape_strikes(root)
        else:
            if progress_callback:
                progress_callback("Skipping simulated line detection...")
            self.paragraph_shapes = {}
            self.greenshape_striked_runs = set()

        if progress_value_callback:
            progress_value_callback(40)

        if progress_callback:
            progress_callback("Processing text runs...")

        self.process_element(root)

        if progress_value_callback:
            progress_value_callback(50)

        if progress_callback:
            progress_callback("Removing colored graphical shapes...")

        self.parent_map = {c: p for p in root.iter() for c in p}
        self.remove_colored_shapes(root)

        if progress_value_callback:
            progress_value_callback(70)

        if progress_callback:
            progress_callback("Removing empty text boxes...")

        self.parent_map = {c: p for p in root.iter() for c in p}
        self.remove_empty_textboxes(root)

        if progress_value_callback:
            progress_value_callback(85)

        # Write the XML file properly
        self.write_xml(tree, xml_path)

    def write_xml(self, tree, xml_path):
        """Write XML file with proper encoding and declarations."""
        if USING_LXML:
            # lxml preserves namespaces correctly
            tree.write(
                xml_path,
                encoding='UTF-8',
                xml_declaration=True,
                standalone=True
            )
        else:
            # For standard ElementTree, write with declaration
            tree.write(xml_path, encoding='UTF-8', xml_declaration=True)

    def process_element(self, element):
        """Recursively process elements, handling text runs."""
        w_ns = '{' + NAMESPACES['w'] + '}'
        runs_to_remove = []

        for child in list(element):
            self.process_element(child)
            if child.tag == f'{w_ns}r':
                action = self.analyze_run(child)
                if action == 'delete':
                    runs_to_remove.append(child)
                elif action == 'clean':
                    self.clean_run(child)

        for r in runs_to_remove:
            element.remove(r)

    def get_attr(self, elem, name):
        """Get attribute, checking namespaced and non-namespaced versions."""
        w_ns = '{' + NAMESPACES['w'] + '}'
        val = elem.get(f'{w_ns}{name}')
        if val is None:
            val = elem.get(name)
        return val

    def has_strikethrough(self, rPr):
        """Check explicit w:strike / w:dstrike on a run."""
        w_ns = '{' + NAMESPACES['w'] + '}'
        strike_elem = rPr.find(f'{w_ns}strike')
        if strike_elem is not None:
            v = self.get_attr(strike_elem, 'val')
            if v is None or v.lower() not in ('0', 'false', 'off', 'none'):
                return True

        dstrike_elem = rPr.find(f'{w_ns}dstrike')
        if dstrike_elem is not None:
            v = self.get_attr(dstrike_elem, 'val')
            if v is None or v.lower() not in ('0', 'false', 'off', 'none'):
                return True

        return False

    def is_underlined(self, rPr):
        """Return True if run properties indicate an underline (w:u)."""
        if rPr is None:
            return False
        w_ns = '{' + NAMESPACES['w'] + '}'
        u = rPr.find(f'{w_ns}u')
        if u is None:
            return False
        v = self.get_attr(u, 'val')
        if v is None:
            return True
        return str(v).lower() not in ('0', 'false', 'none', 'off')

    def run_has_visible_text(self, run):
        """Does this run have any non-whitespace text?"""
        w_ns = '{' + NAMESPACES['w'] + '}'
        for t in run.iter(f'{w_ns}t'):
            if t.text and t.text.strip():
                return True
        return False

    def get_ancestor(self, elem, tag):
        """Walk up parent_map to find ancestor with given tag."""
        cur = elem
        while cur in self.parent_map:
            cur = self.parent_map[cur]
            if cur.tag == tag:
                return cur
        return None

    def analyze_run(self, run):
        """Analyze a run and decide: keep / delete / clean."""
        w_ns = '{' + NAMESPACES['w'] + '}'

        rPr = run.find(f'{w_ns}rPr')
        if rPr is None:
            return 'keep'

        color_elem = rPr.find(f'{w_ns}color')
        if color_elem is None:
            return 'keep'

        color_val = self.get_attr(color_elem, 'val')
        if not color_val or color_val == 'auto':
            return 'keep'

        # structural runs: tabs, line breaks, etc.
        is_structural = False
        if run.find(f'{w_ns}tab') is not None or \
           run.find(f'{w_ns}br') is not None or \
           run.find(f'{w_ns}cr') is not None:
            is_structural = True

        has_strike = self.has_strikethrough(rPr)

        # Red -> delete (except structural: clean)
        if is_color_match(color_val, self.RED, self.red_tolerance):
            if is_structural:
                self.stats['red_deleted'] += 1
                return 'clean'
            self.stats['red_deleted'] += 1
            return 'delete'

        # Green
        if is_color_match(color_val, self.GREEN, self.green_tolerance):
            is_under = self.is_underlined(rPr)

            # Explicit strike always wins
            if has_strike:
                if is_structural:
                    self.stats['green_strike_deleted'] += 1
                    return 'clean'
                self.stats['green_strike_deleted'] += 1
                return 'delete'

            # Simulated strikethrough?
            has_simulated_strike = False
            if self.check_simulated_lines:
                if run in self.greenshape_striked_runs:
                    has_simulated_strike = True
                else:
                    # local neighborhood of runs in the paragraph
                    para = self.get_ancestor(run, f'{w_ns}p')
                    if para is not None and self.greenshape_striked_runs:
                        runs = list(para.iter(f'{w_ns}r'))
                        try:
                            idx = runs.index(run)
                        except ValueError:
                            idx = None
                        if idx is not None:
                            start = max(0, idx - 4)
                            end = min(len(runs), idx + 5)
                            for j in range(start, end):
                                if runs[j] in self.greenshape_striked_runs:
                                    has_simulated_strike = True
                                    break

            if has_simulated_strike:
                if is_structural:
                    self.stats['green_strike_deleted'] += 1
                    return 'clean'
                self.stats['green_strike_deleted'] += 1
                return 'delete'

            # Otherwise -> clean formatting only
            self.stats['green_cleaned'] += 1
            return 'clean'

        # Blue -> clean formatting
        if is_color_match(color_val, self.BLUE, self.blue_tolerance):
            self.stats['blue_cleaned'] += 1
            return 'clean'

        return 'keep'

    def clean_run(self, run):
        """Remove color, underline, bold, italic, strikethrough, shading, highlight."""
        w_ns = '{' + NAMESPACES['w'] + '}'
        rPr = run.find(f'{w_ns}rPr')
        if rPr is None:
            return

        tags = [
            f'{w_ns}color',
            f'{w_ns}u',
            f'{w_ns}b',
            f'{w_ns}bCs',
            f'{w_ns}i',
            f'{w_ns}iCs',
            f'{w_ns}strike',
            f'{w_ns}dstrike',
            f'{w_ns}shd',
            f'{w_ns}highlight',
        ]
        for child in list(rPr):
            if child.tag in tags:
                rPr.remove(child)

    def classify_shape_role(self, vshape):
        """
        Classify a VML shape as 'strike' or 'underline' based on style/top heuristic.
        """
        style = vshape.get('style') or ''
        top_pt = self.parse_top_from_style(style)
        role = 'strike'
        if top_pt is not None and top_pt >= 8.0:
            role = 'underline'
        return role

    def detect_green_shape_strikes(self, root):
        """
        Detect green drawn lines that act as strikethroughs and flag nearby green runs.
        """
        a_ns = '{' + NAMESPACES['a'] + '}'
        mc_ns = '{' + NAMESPACES['mc'] + '}'
        w_ns = '{' + NAMESPACES['w'] + '}'
        v_ns = '{' + NAMESPACES['v'] + '}'

        self.greenshape_striked_runs = set()
        candidates = []

        # DrawingML shapes
        for srgb in root.iter(f'{a_ns}srgbClr'):
            color_val = srgb.get('val')
            if color_val and is_color_match(color_val, self.GREEN, self.green_tolerance):
                cur = srgb
                while cur in self.parent_map and cur.tag != f'{mc_ns}AlternateContent':
                    cur = self.parent_map[cur]
                if cur is not None and cur.tag == f'{mc_ns}AlternateContent':
                    role = 'strike'
                    for vshape in cur.iter():
                        if vshape.tag in (f'{v_ns}line', f'{v_ns}shape', f'{v_ns}rect'):
                            role = self.classify_shape_role(vshape)
                            break
                    if role == 'strike':
                        candidates.append(cur)

        # VML shapes directly
        for vshape in root.iter():
            if vshape.tag in (f'{v_ns}line', f'{v_ns}shape', f'{v_ns}rect'):
                fillcolor = vshape.get('fillcolor') or vshape.get('strokecolor') or vshape.get('color')
                if not fillcolor:
                    fill_child = vshape.find('.//' + v_ns + 'fill')
                    if fill_child is not None:
                        fillcolor = fill_child.get('color') or fill_child.get('fillcolor')
                if fillcolor and is_color_match(fillcolor.lstrip('#').upper(), self.GREEN, self.green_tolerance):
                    role = self.classify_shape_role(vshape)
                    if role == 'strike':
                        cur = vshape
                        while cur in self.parent_map and cur.tag != f'{mc_ns}AlternateContent':
                            cur = self.parent_map[cur]
                        candidates.append(cur)

        # For each candidate, flag green text runs in that paragraph
        for cand in candidates:
            cur = cand
            while cur is not None and cur.tag != f'{w_ns}p':
                cur = self.parent_map.get(cur)
            if cur is None:
                continue
            para = cur
            for run in para.iter(f'{w_ns}r'):
                rPr = run.find(f'{w_ns}rPr')
                if rPr is None:
                    continue
                color_elem = rPr.find(f'{w_ns}color')
                if color_elem is None:
                    continue
                color_val = self.get_attr(color_elem, 'val')
                if color_val and is_color_match(color_val, self.GREEN, self.green_tolerance):
                    # Only flag runs with visible text and no explicit underline
                    if (not self.is_underlined(rPr)) and self.run_has_visible_text(run):
                        self.greenshape_striked_runs.add(run)

    def parse_top_from_style(self, style):
        """Parse 'top' value from style string (in pt/px/in → pt)."""
        if not style:
            return None
        m = re.search(r'top\s*:\s*([-\d.]+)\s*(pt|px|in)?', style, flags=re.IGNORECASE)
        if not m:
            return None
        val = float(m.group(1))
        unit = m.group(2)
        if not unit or unit.lower() == 'pt':
            return val
        if unit.lower() == 'px':
            return val * 0.75
        if unit.lower() == 'in':
            return val * 72.0
        return val

    def build_paragraph_shape_map(self, root):
        """Map paragraph -> list of shapes (currently not heavily used, but kept)."""
        w_ns = '{' + NAMESPACES['w'] + '}'
        mc_ns = '{' + NAMESPACES['mc'] + '}'
        a_ns = '{' + NAMESPACES['a'] + '}'
        v_ns = '{' + NAMESPACES['v'] + '}'

        para_map = {}

        for p in root.iter(f'{w_ns}p'):
            found = []

            for alt in p.iter():
                # DrawingML colors
                for srgb in alt.findall('.//' + a_ns + 'srgbClr'):
                    color_val = srgb.get('val')
                    if not color_val:
                        continue
                    color_val = color_val.lstrip('#').upper()
                    role = 'strike'
                    cur = srgb
                    while cur in self.parent_map and cur.tag != f'{mc_ns}AlternateContent':
                        cur = self.parent_map[cur]
                    if cur is not None and cur.tag == f'{mc_ns}AlternateContent':
                        for vshape in cur.iter():
                            if vshape.tag in (f'{v_ns}line', f'{v_ns}shape', f'{v_ns}rect'):
                                role = self.classify_shape_role(vshape)
                                break
                    found.append({'color': color_val, 'role': role})

                # VML shapes
                for tagname in (f'{v_ns}line', f'{v_ns}shape', f'{v_ns}rect'):
                    for vshape in alt.findall('.//' + tagname):
                        strokecolor = vshape.get('strokecolor') or vshape.get('stroke') or vshape.get('color')
                        fillcolor = vshape.get('fillcolor')
                        color_val = strokecolor or fillcolor
                        if not color_val:
                            fill_child = vshape.find('.//' + v_ns + 'fill')
                            if fill_child is not None:
                                color_val = fill_child.get('color') or fill_child.get('fillcolor')
                        if not color_val:
                            continue
                        color_val = color_val.lstrip('#').upper()
                        role = self.classify_shape_role(vshape)
                        found.append({'color': color_val, 'role': role})

            if found:
                para_map[p] = found

        return para_map

    def remove_colored_shapes(self, root):
        """Remove colored graphical shapes (non-textbox shapes)."""
        mc_ns = '{' + NAMESPACES['mc'] + '}'
        a_ns = '{' + NAMESPACES['a'] + '}'
        v_ns = '{' + NAMESPACES['v'] + '}'
        w_ns = '{' + NAMESPACES['w'] + '}'

        to_remove = []

        for alt in root.iter(f'{mc_ns}AlternateContent'):
            should_remove = False

            # DrawingML colors
            for srgb in alt.iter(f'{a_ns}srgbClr'):
                color_val = srgb.get('val')
                if color_val:
                    color_val = color_val.upper()
                    if (is_color_match(color_val, self.RED, self.red_tolerance) or
                        is_color_match(color_val, self.GREEN, self.green_tolerance) or
                        is_color_match(color_val, self.BLUE, self.blue_tolerance)):
                        if alt.find(f'.//{w_ns}txbxContent') is None:
                            should_remove = True
                            break

            if should_remove:
                to_remove.append(alt)
                continue

            # VML shapes
            for vml in alt.iter():
                if vml.tag in (f'{v_ns}rect', f'{v_ns}shape', f'{v_ns}line', f'{v_ns}oval'):
                    fillcolor = vml.get('fillcolor') or vml.get('strokecolor') or vml.get('color')
                    if fillcolor:
                        fillcolor = fillcolor.lstrip('#').upper()
                        if (is_color_match(fillcolor, self.RED, self.red_tolerance) or
                            is_color_match(fillcolor, self.GREEN, self.green_tolerance) or
                            is_color_match(fillcolor, self.BLUE, self.blue_tolerance)):
                            if vml.find(f'.//{v_ns}textbox') is None:
                                should_remove = True
                                break

            if should_remove:
                to_remove.append(alt)

        for elem in to_remove:
            parent = self.parent_map.get(elem)
            if parent is not None:
                try:
                    parent.remove(elem)
                    self.stats['colored_shapes_removed'] += 1
                except ValueError:
                    pass

    def remove_empty_textboxes(self, root):
        """Remove completely empty text boxes."""
        w_ns = '{' + NAMESPACES['w'] + '}'

        self.parent_map = {c: p for p in root.iter() for c in p}
        txbxs = list(root.iter(f'{w_ns}txbxContent'))
        to_remove = []

        for tx in txbxs:
            if self.textbox_is_empty(tx):
                target = self.find_removal_target(tx)
                if target is not None and target not in to_remove:
                    to_remove.append(target)

        for elem in to_remove:
            parent = self.parent_map.get(elem)
            if parent is not None:
                try:
                    parent.remove(elem)
                    self.stats['textboxes_removed'] += 1
                except ValueError:
                    pass

    def textbox_is_empty(self, txbx_content):
        """Heuristic: no text, no drawings, no objects."""
        w_ns = '{' + NAMESPACES['w'] + '}'
        for t in txbx_content.iter(f'{w_ns}t'):
            if t.text and t.text.strip():
                return False
        for _ in txbx_content.iter(f'{w_ns}drawing'):
            return False
        for _ in txbx_content.iter(f'{w_ns}object'):
            return False
        return True

    def find_removal_target(self, txbx_content):
        """Find the appropriate parent element to remove for an empty textbox."""
        mc_ns = '{' + NAMESPACES['mc'] + '}'
        w_ns = '{' + NAMESPACES['w'] + '}'

        current = txbx_content
        while current in self.parent_map:
            parent = self.parent_map[current]
            if parent.tag == f'{mc_ns}AlternateContent':
                return parent
            if parent.tag == f'{w_ns}r':
                return parent
            current = parent
        return None

    def create_docx(self, extract_dir, output_path):
        """Re-pack directory into a .docx file."""
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root_dir, dirs, files in os.walk(extract_dir):
                for f in files:
                    fp = os.path.join(root_dir, f)
                    arc = os.path.relpath(fp, extract_dir)
                    z.write(fp, arc)


# ---------- small helpers ----------

def make_output_docx(input_path):
    p = Path(input_path)
    return str(p.parent / f"(No Blacklines) {p.name}")


def make_output_pdf(docx_output):
    p = Path(docx_output)
    return str(p.with_suffix(".pdf"))


# ---------- Worker for threading ----------

class ProcessorWorker(QObject):
    progress = pyqtSignal(int)          # 0-100
    status = pyqtSignal(str)
    finished = pyqtSignal(list, dict)   # outputs, total_stats
    error = pyqtSignal(str)

    def __init__(self, files, mode,
                 red_tolerance, green_tolerance, blue_tolerance,
                 check_simulated,
                 red_color, green_color, blue_color,
                 parent=None):
        super().__init__(parent)
        self.files = files
        self.mode = mode  # 'docx' or 'pdf'
        self.red_tolerance = red_tolerance
        self.green_tolerance = green_tolerance
        self.blue_tolerance = blue_tolerance
        self.check_simulated = check_simulated
        self.red_color = red_color
        self.green_color = green_color
        self.blue_color = blue_color

    def run(self):
        try:
            total_files = len(self.files)
            processed_outputs = []
            total_stats = {
                'red_deleted': 0,
                'green_strike_deleted': 0,
                'green_cleaned': 0,
                'blue_cleaned': 0,
                'textboxes_removed': 0,
                'colored_shapes_removed': 0,
            }

            self.status.emit(f"Processing {total_files} file(s)...")
            self.status.emit(f"Using {'lxml' if USING_LXML else 'ElementTree'} for XML processing")
            self.status.emit(f"Red tolerance: {self.red_tolerance}")
            self.status.emit(f"Green tolerance: {self.green_tolerance}")
            self.status.emit(f"Blue tolerance: {self.blue_tolerance}")
            self.status.emit(f"Red target color: {self.red_color}")
            self.status.emit(f"Green target color: {self.green_color}")
            self.status.emit(f"Blue target color: {self.blue_color}")
            self.status.emit(f"Check simulated lines: {self.check_simulated}")
            self.status.emit("-" * 40)

            for idx, input_path in enumerate(self.files, start=1):
                base = os.path.basename(input_path)
                self.status.emit(f"[{idx}/{total_files}] Processing: {base}")

                output_docx = make_output_docx(input_path)

                processor = WordColorProcessor(
                    red_tolerance=self.red_tolerance,
                    green_tolerance=self.green_tolerance,
                    blue_tolerance=self.blue_tolerance,
                    check_simulated_lines=self.check_simulated,
                    red_color=self.red_color,
                    green_color=self.green_color,
                    blue_color=self.blue_color
                )

                def progress_callback(msg, idx=idx, total_files=total_files):
                    self.status.emit(f"[{idx}/{total_files}] {msg}")

                def progress_value_callback(val, idx=idx, total_files=total_files):
                    global_val = ((idx - 1) + (val / 100.0)) / total_files * 100.0
                    self.progress.emit(int(global_val))

                stats = processor.process_document(
                    input_path, output_docx,
                    progress_callback=progress_callback,
                    progress_value_callback=progress_value_callback
                )

                final_output = output_docx
                if self.mode == 'pdf' and DOCX2PDF_AVAILABLE:
                    output_pdf = make_output_pdf(output_docx)
                    self.status.emit(f"[{idx}/{total_files}] Converting to PDF: {os.path.basename(output_pdf)}")
                    docx2pdf_convert(output_docx, output_pdf)
                    # Optionally delete the intermediate docx
                    try:
                        os.remove(output_docx)
                    except OSError:
                        pass
                    final_output = output_pdf

                processed_outputs.append(final_output)

                for k in total_stats:
                    total_stats[k] += stats.get(k, 0)

                self.status.emit(
                    f"[{idx}/{total_files}] Done. "
                    f"Red deleted: {stats['red_deleted']}, "
                    f"Green strike deleted: {stats['green_strike_deleted']}, "
                    f"Green cleaned: {stats['green_cleaned']}, "
                    f"Blue cleaned: {stats['blue_cleaned']}, "
                    f"Colored shapes removed: {stats['colored_shapes_removed']}, "
                    f"Empty text boxes removed: {stats['textboxes_removed']}"
                )
                self.status.emit("-" * 20)

            self.progress.emit(100)
            self.finished.emit(processed_outputs, total_stats)

        except Exception as e:
            import traceback
            self.error.emit(f"{str(e)}\n{traceback.format_exc()}")


# ---------- GUI helpers ----------

class ColorSwatch(QFrame):
    """Clickable color box."""
    clicked = pyqtSignal()

    def __init__(self, color, parent=None):
        super().__init__(parent)
        self._color = color
        self.setFixedSize(24, 24)
        self.setFrameShape(QFrame.Shape.Box)
        self.setFrameShadow(QFrame.Shadow.Plain)
        self._apply_style()

    def _apply_style(self):
        self.setStyleSheet(f"background-color: {self._color}; border: 1px solid #444;")

    def set_color(self, color):
        self._color = color
        self._apply_style()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


# ---------- Main Window ----------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Blackline Remover by Emil Ferdman")
        self.resize(900, 700)
        self.setMinimumSize(900, 700)
        self.setAcceptDrops(True)

        self.input_files = []
        self.red_color = "#FF0000"
        self.green_color = "#008000"
        self.blue_color = "#0000FF"

        self.worker_thread = None
        self.worker = None

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        # Title
        title_label = QLabel("Blackline Remover by Emil Ferdman")
        font = title_label.font()
        font.setPointSize(16)
        font.setBold(True)
        title_label.setFont(font)
        main_layout.addWidget(title_label)

        # Description
        xml_lib = "lxml (recommended)" if USING_LXML else "ElementTree (install lxml for better compatibility)"
        desc = QLabel(
            f"This tool processes Word documents to:\n"
            "• Delete red text and green strikethrough text\n"
            "• Remove formatting from remaining green and blue text\n"
            "• Remove colored graphical shapes (lines/rectangles)\n"
            "• Remove empty text boxes\n"
            "• Output saves as '(No Blacklines) filename.docx' or PDF\n\n"
            f"Drag and drop .docx files onto this window or use Browse.\n"
            f"XML Library: {xml_lib}"
        )
        desc.setWordWrap(True)
        main_layout.addWidget(desc)

        # File selection group
        file_group = QGroupBox("File Selection")
        file_layout = QHBoxLayout(file_group)
        main_layout.addWidget(file_group)

        self.input_line = QLineEdit()
        self.input_line.setReadOnly(True)
        file_layout.addWidget(QLabel("Input Files:"))
        file_layout.addWidget(self.input_line)
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_files)
        file_layout.addWidget(browse_btn)

        # Color settings group
        color_group = QGroupBox("Color Settings & Tolerances (0-100)")
        color_layout = QVBoxLayout(color_group)
        main_layout.addWidget(color_group)

        # Red
        self.red_slider, self.red_swatch, self.red_hex_label = self.create_color_row(
            color_layout, "Red (delete all):", "red_color", 10
        )
        # Green
        self.green_slider, self.green_swatch, self.green_hex_label = self.create_color_row(
            color_layout, "Green (delete if strikethrough):", "green_color", 10
        )
        # Blue
        self.blue_slider, self.blue_swatch, self.blue_hex_label = self.create_color_row(
            color_layout, "Blue (remove formatting):", "blue_color", 10
        )

        # Options group
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout(options_group)
        main_layout.addWidget(options_group)

        self.simulated_checkbox = QCheckBox(
            "Detect simulated strikethroughs (drawn lines over green text)"
        )
        self.simulated_checkbox.setChecked(True)
        options_layout.addWidget(self.simulated_checkbox)

        help_label = QLabel(
            "When enabled, green text with drawn line shapes positioned as\n"
            "strikethroughs will be deleted. Underlines are preserved unless\n"
            "clearly associated with a strikethrough."
        )
        help_label.setWordWrap(True)
        options_layout.addWidget(help_label)

        # Progress group
        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout(progress_group)
        main_layout.addWidget(progress_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        progress_layout.addWidget(self.progress_bar)
        self.progress_label = QLabel("Ready")
        progress_layout.addWidget(self.progress_label)

        # Buttons (DOCX / PDF)
        buttons_layout = QHBoxLayout()
        main_layout.addLayout(buttons_layout)

        self.process_docx_btn = QPushButton("Process Document(s)")
        self.process_docx_btn.clicked.connect(lambda: self.start_processing("docx"))
        buttons_layout.addWidget(self.process_docx_btn)

        self.process_pdf_btn = QPushButton("Process PDF(s)")
        self.process_pdf_btn.clicked.connect(lambda: self.start_processing("pdf"))
        buttons_layout.addWidget(self.process_pdf_btn)

        # Status log
        log_group = QGroupBox("Status Log")
        log_layout = QVBoxLayout(log_group)
        main_layout.addWidget(log_group, stretch=1)

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        log_layout.addWidget(self.log_edit)

        # Hint if docx2pdf not installed
        if not DOCX2PDF_AVAILABLE:
            self.process_pdf_btn.setToolTip(
                "docx2pdf is not installed. PDF processing will not work until you install it:\n"
                "pip install docx2pdf"
            )

    # ----- Color row helper -----

    def create_color_row(self, parent_layout, label_text, color_attr_name, default_tolerance):
        row = QHBoxLayout()
        parent_layout.addLayout(row)

        color_value = getattr(self, color_attr_name)
        swatch = ColorSwatch(color_value)
        row.addWidget(swatch)

        label = QLabel(label_text)
        label.setMinimumWidth(220)
        row.addWidget(label)

        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(0, 100)
        slider.setValue(default_tolerance)
        row.addWidget(slider, stretch=1)

        value_label = QLabel(str(default_tolerance))
        value_label.setMinimumWidth(30)
        row.addWidget(value_label)

        hex_label = QLabel(color_value)
        hex_label.setMinimumWidth(80)
        row.addWidget(hex_label)

        slider.valueChanged.connect(lambda v, lbl=value_label: lbl.setText(str(v)))

        def pick_color():
            current = getattr(self, color_attr_name)
            initial = QColor(current)
            color = QColorDialog.getColor(initial, self, f"Select color for {label_text}")
            if color.isValid():
                hex_str = color.name().upper()  # '#RRGGBB'
                setattr(self, color_attr_name, hex_str)
                swatch.set_color(hex_str)
                hex_label.setText(hex_str)

        swatch.clicked.connect(pick_color)

        return slider, swatch, hex_label

    # ----- Drag & drop support -----

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.isLocalFile() and url.toLocalFile().lower().endswith(".docx"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        paths = []
        for url in event.mimeData().urls():
            if url.isLocalFile():
                p = url.toLocalFile()
                if p.lower().endswith(".docx"):
                    paths.append(p)
        if paths:
            self.add_input_files(paths, append=True)
        event.acceptProposedAction()

    # ----- File handling -----

    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Word Document(s)",
            "",
            "Word Documents (*.docx);;All Files (*.*)"
        )
        if files:
            self.add_input_files(files, append=False)

    def add_input_files(self, paths, append=True):
        valid = [p for p in paths if p.lower().endswith(".docx") and os.path.isfile(p)]
        if not valid:
            QMessageBox.warning(self, "No valid files", "Please select .docx files only.")
            return
        if not append:
            self.input_files = []
        for p in valid:
            if p not in self.input_files:
                self.input_files.append(p)
        self.update_input_display()

    def update_input_display(self):
        if not self.input_files:
            self.input_line.setText("")
        elif len(self.input_files) == 1:
            self.input_line.setText(self.input_files[0])
        else:
            self.input_line.setText(f"{len(self.input_files)} files selected")

    # ----- Logging & UI updates -----

    def log(self, message):
        self.log_edit.append(message)
        self.log_edit.verticalScrollBar().setValue(
            self.log_edit.verticalScrollBar().maximum()
        )

    # ----- Processing -----

    def start_processing(self, mode):
        if not self.input_files:
            QMessageBox.warning(self, "No files", "Please select at least one .docx file.")
            return

        if mode == 'pdf' and not DOCX2PDF_AVAILABLE:
            QMessageBox.critical(
                self,
                "PDF Conversion Not Available",
                "PDF processing requires the 'docx2pdf' package.\n\n"
                "Install it with:\n    pip install docx2pdf"
            )
            return

        self.log_edit.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText("Starting...")
        self.process_docx_btn.setEnabled(False)
        self.process_pdf_btn.setEnabled(False)

        self.worker_thread = QThread()
        self.worker = ProcessorWorker(
            files=list(self.input_files),
            mode=mode,
            red_tolerance=self.red_slider.value(),
            green_tolerance=self.green_slider.value(),
            blue_tolerance=self.blue_slider.value(),
            check_simulated=self.simulated_checkbox.isChecked(),
            red_color=self.red_color,
            green_color=self.green_color,
            blue_color=self.blue_color,
        )
        self.worker.moveToThread(self.worker_thread)

        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_worker_progress)
        self.worker.status.connect(self.on_worker_status)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.error.connect(self.on_worker_error)

        self.worker.finished.connect(self.worker_thread.quit)
        self.worker.error.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)

        self.worker_thread.start()

    def on_worker_progress(self, value):
        self.progress_bar.setValue(value)

    def on_worker_status(self, message):
        self.progress_label.setText(message)
        self.log(message)

    def on_worker_finished(self, outputs, total_stats):
        self.log("All processing complete!")
        self.log(
            f"TOTALS across {len(self.input_files)} file(s):\n"
            f"  Red text runs deleted: {total_stats['red_deleted']}\n"
            f"  Green strikethrough deleted: {total_stats['green_strike_deleted']}\n"
            f"  Green text cleaned: {total_stats['green_cleaned']}\n"
            f"  Blue text cleaned: {total_stats['blue_cleaned']}\n"
            f"  Colored shapes removed: {total_stats['colored_shapes_removed']}\n"
            f"  Empty text boxes removed: {total_stats['textboxes_removed']}\n"
        )

        if outputs:
            if len(outputs) == 1:
                out_path = outputs[0]
                self.log(f"Saved to: {out_path}")
                self.progress_label.setText("Complete!")
                reply = QMessageBox.question(
                    self,
                    "Open Processed File",
                    f"Processing complete.\n\nOpen processed file?\n\n{out_path}",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.Yes:
                    self.open_path(out_path)
            else:
                self.log("Processed files:")
                for p in outputs:
                    self.log(f"  {p}")
                self.progress_label.setText("Complete!")
                folder = os.path.dirname(outputs[0])
                reply = QMessageBox.question(
                    self,
                    "Open Output Folder",
                    f"Processing complete.\n\nOpen output folder?\n\n{folder}",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.Yes:
                    self.open_path(folder)

        self.process_docx_btn.setEnabled(True)
        self.process_pdf_btn.setEnabled(True)

    def on_worker_error(self, message):
        self.log(f"ERROR: {message}")
        self.progress_label.setText("Error!")
        self.progress_bar.setValue(0)
        self.process_docx_btn.setEnabled(True)
        self.process_pdf_btn.setEnabled(True)
        QMessageBox.critical(self, "Error", f"An error occurred:\n{message}")

    def open_path(self, path):
        """Open a file or folder with the OS default handler."""
        try:
            if sys.platform.startswith("darwin"):
                subprocess.Popen(["open", path])
            elif os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            QMessageBox.critical(
                self,
                "Open Failed",
                f"Could not open:\n{path}\n\nError:\n{e}"
            )


# ---------- main ----------

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()