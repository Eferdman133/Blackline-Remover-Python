#!/usr/bin/env python3
"""
Blackline Remover by Emil Ferdman - PyQt6 GUI (Optimized Version with Preview)

Features:
- Delete red text and green strikethrough text
- Remove formatting from remaining green and blue text
- Detect simulated strikethroughs (drawn green lines)
- Remove colored graphical shapes and empty text boxes
- Drag & drop .docx files, or browse and select multiple files
- Adjustable color tolerances and target RGB colors (via color pickers)
- Process to cleaned .docx OR to PDF (requires docx2pdf)
- PREVIEW mode: Shows what will be deleted (light red highlight) and cleaned (cyan highlight)
- Prompt to open processed file / output folder after completion

Optimizations applied:
- Pre-computed namespace strings
- Single parent_map build per document
- Cached RGB color conversions
- Batched element operations
- Early exits in hot paths
- Reduced redundant iterations
- Single-pass shape scanning with deferred paragraph lookup
"""

import sys
import os
import zipfile
import tempfile
import re
import subprocess
from pathlib import Path
from functools import lru_cache

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

# =============================================================================
# PRE-COMPUTED NAMESPACE STRINGS (Optimization #1)
# =============================================================================
# These are computed once at module load instead of repeatedly during processing

W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
W14_NS = '{http://schemas.microsoft.com/office/word/2010/wordml}'
A_NS = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
MC_NS = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
V_NS = '{urn:schemas-microsoft-com:vml}'
WP_NS = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'

# Pre-computed tag names for hot paths
W_R = f'{W_NS}r'
W_RPR = f'{W_NS}rPr'
W_COLOR = f'{W_NS}color'
W_T = f'{W_NS}t'
W_P = f'{W_NS}p'
W_STRIKE = f'{W_NS}strike'
W_DSTRIKE = f'{W_NS}dstrike'
W_U = f'{W_NS}u'
W_TAB = f'{W_NS}tab'
W_BR = f'{W_NS}br'
W_CR = f'{W_NS}cr'
W_B = f'{W_NS}b'
W_BCS = f'{W_NS}bCs'
W_I = f'{W_NS}i'
W_ICS = f'{W_NS}iCs'
W_SHD = f'{W_NS}shd'
W_HIGHLIGHT = f'{W_NS}highlight'
W_DRAWING = f'{W_NS}drawing'
W_OBJECT = f'{W_NS}object'
W_TXBXCONTENT = f'{W_NS}txbxContent'
W_VAL = f'{W_NS}val'

MC_ALTCONTENT = f'{MC_NS}AlternateContent'
A_SRGBCLR = f'{A_NS}srgbClr'
V_LINE = f'{V_NS}line'
V_SHAPE = f'{V_NS}shape'
V_RECT = f'{V_NS}rect'
V_OVAL = f'{V_NS}oval'
V_TEXTBOX = f'{V_NS}textbox'
V_FILL = f'{V_NS}fill'

# Tags to clean from run properties (pre-computed set for O(1) lookup)
CLEAN_TAGS = frozenset({
    W_COLOR, W_U, W_B, W_BCS, W_I, W_ICS,
    W_STRIKE, W_DSTRIKE, W_SHD, W_HIGHLIGHT
})

# VML shape tags (pre-computed tuple for iteration)
VML_SHAPE_TAGS = (V_LINE, V_SHAPE, V_RECT, V_OVAL)

# VML shape tags as a set for O(1) lookup
VML_SHAPE_TAGS_SET = frozenset(VML_SHAPE_TAGS[:3])  # line, shape, rect

# XML namespaces dict (kept for ElementTree registration and lxml XPath)
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


# =============================================================================
# COLOR HELPERS (Optimization #3 - Cached conversions)
# =============================================================================

@lru_cache(maxsize=256)
def hex_to_rgb(hex_color):
    """Convert hex color to RGB tuple. Cached for repeated calls."""
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 6:
        return (
            int(hex_color[0:2], 16),
            int(hex_color[2:4], 16),
            int(hex_color[4:6], 16)
        )
    return None


def color_distance_sq(c1, c2):
    """Squared color distance (avoids sqrt for comparison)."""
    if c1 is None or c2 is None:
        return float('inf')
    return (c1[0] - c2[0]) ** 2 + (c1[1] - c2[1]) ** 2 + (c1[2] - c2[2]) ** 2


@lru_cache(maxsize=1024)
def is_color_match(hex_color, target_hex, tolerance):
    """
    Check if a color matches the target within tolerance (0-100).
    Cached for repeated comparisons of the same color pairs.
    """
    if not hex_color:
        return False
    hex_color = hex_color.lstrip('#').upper()
    target_hex = target_hex.lstrip('#').upper()
    c1 = hex_to_rgb(hex_color)
    c2 = hex_to_rgb(target_hex)
    if c1 is None or c2 is None:
        return False
    # Tolerance 0-100 scaled to RGB max distance ~441
    # Using squared distance to avoid sqrt
    max_dist_sq = (tolerance * 4.41) ** 2
    return color_distance_sq(c1, c2) <= max_dist_sq


# =============================================================================
# CORE PROCESSOR (Optimized)
# =============================================================================

class WordColorProcessor:
    """Handles the actual document processing with optimizations."""

    __slots__ = (
        'red_tolerance', 'green_tolerance', 'blue_tolerance',
        'check_simulated_lines', 'RED', 'GREEN', 'BLUE',
        'stats', 'parent_map', 'paragraph_shapes', 'greenshape_striked_runs',
        'preview_mode'
    )

    def __init__(self, red_tolerance=10, green_tolerance=10, blue_tolerance=10,
                 check_simulated_lines=True,
                 red_color='#FF0000', green_color='#008000', blue_color='#0000FF',
                 preview_mode=False):
        self.red_tolerance = red_tolerance
        self.green_tolerance = green_tolerance
        self.blue_tolerance = blue_tolerance
        self.check_simulated_lines = check_simulated_lines
        self.preview_mode = preview_mode

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
        # Reset stats
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
                if self.preview_mode:
                    progress_callback("Generating preview (highlighting changes)...")
                else:
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

        # Optimization #2: Build parent map ONCE
        self.parent_map = {c: p for p in root.iter() for c in p}

        if progress_value_callback:
            progress_value_callback(30)

        if self.check_simulated_lines:
            if progress_callback:
                progress_callback("Scanning for drawn line shapes...")
            self.detect_green_shape_strikes(root)
        else:
            if progress_callback:
                progress_callback("Skipping simulated line detection...")
            self.greenshape_striked_runs = set()

        if progress_value_callback:
            progress_value_callback(40)

        if progress_callback:
            if self.preview_mode:
                progress_callback("Highlighting text runs for preview...")
            else:
                progress_callback("Processing text runs...")

        self.process_element(root)

        if progress_value_callback:
            progress_value_callback(50)

        # In preview mode, we highlight shapes instead of removing them
        if not self.preview_mode:
            if progress_callback:
                progress_callback("Removing colored graphical shapes...")

            # Rebuild parent map only if structure changed significantly
            self.parent_map = {c: p for p in root.iter() for c in p}
            self.remove_colored_shapes(root)

            if progress_value_callback:
                progress_value_callback(70)

            if progress_callback:
                progress_callback("Removing empty text boxes...")

            # Rebuild once more for textbox removal
            self.parent_map = {c: p for p in root.iter() for c in p}
            self.remove_empty_textboxes(root)
        else:
            if progress_callback:
                progress_callback("Skipping shape/textbox removal in preview mode...")
            if progress_value_callback:
                progress_value_callback(70)

        if progress_value_callback:
            progress_value_callback(85)

        self.write_xml(tree, xml_path)

    def write_xml(self, tree, xml_path):
        """Write XML file with proper encoding and declarations."""
        if USING_LXML:
            tree.write(
                xml_path,
                encoding='UTF-8',
                xml_declaration=True,
                standalone=True
            )
        else:
            tree.write(xml_path, encoding='UTF-8', xml_declaration=True)

    def process_element(self, element):
        """Recursively process elements, handling text runs."""
        runs_to_remove = []
        runs_to_clean = []
        runs_to_highlight_delete = []
        runs_to_highlight_clean = []

        for child in list(element):
            self.process_element(child)
            if child.tag == W_R:
                action = self.analyze_run(child)
                if self.preview_mode:
                    # In preview mode, highlight instead of delete/clean
                    if action == 'delete':
                        runs_to_highlight_delete.append(child)
                    elif action == 'clean':
                        runs_to_highlight_clean.append(child)
                else:
                    if action == 'delete':
                        runs_to_remove.append(child)
                    elif action == 'clean':
                        runs_to_clean.append(child)

        # Batch operations (Optimization #4)
        if self.preview_mode:
            for r in runs_to_highlight_delete:
                self.highlight_run(r, 'red')   # Light red shading for deletions
            for r in runs_to_highlight_clean:
                self.highlight_run(r, 'cyan')  # Cyan highlight for cleaning
        else:
            for r in runs_to_remove:
                element.remove(r)
            for r in runs_to_clean:
                self.clean_run(r)

    def get_attr(self, elem, name):
        """Get attribute, checking namespaced and non-namespaced versions."""
        val = elem.get(f'{W_NS}{name}')
        if val is None:
            val = elem.get(name)
        return val

    def set_attr(self, elem, name, value):
        """Set attribute with namespace."""
        elem.set(f'{W_NS}{name}', value)

    def has_strikethrough(self, rPr):
        """Check explicit w:strike / w:dstrike on a run."""
        strike_elem = rPr.find(W_STRIKE)
        if strike_elem is not None:
            v = self.get_attr(strike_elem, 'val')
            if v is None or v.lower() not in ('0', 'false', 'off', 'none'):
                return True

        dstrike_elem = rPr.find(W_DSTRIKE)
        if dstrike_elem is not None:
            v = self.get_attr(dstrike_elem, 'val')
            if v is None or v.lower() not in ('0', 'false', 'off', 'none'):
                return True

        return False

    def is_underlined(self, rPr):
        """Return True if run properties indicate an underline (w:u)."""
        if rPr is None:
            return False
        u = rPr.find(W_U)
        if u is None:
            return False
        v = self.get_attr(u, 'val')
        if v is None:
            return True
        return v.lower() not in ('0', 'false', 'none', 'off')

    def run_has_visible_text(self, run):
        """Does this run have any non-whitespace text?"""
        for t in run.iter(W_T):
            if t.text and t.text.strip():
                return True
        return False

    def get_ancestor(self, elem, tag):
        """Walk up parent_map to find ancestor with given tag."""
        cur = elem
        parent_map = self.parent_map  # Local reference for speed
        while cur in parent_map:
            cur = parent_map[cur]
            if cur.tag == tag:
                return cur
        return None

    def analyze_run(self, run):
        """Analyze a run and decide: keep / delete / clean."""
        rPr = run.find(W_RPR)
        if rPr is None:
            return 'keep'

        color_elem = rPr.find(W_COLOR)
        if color_elem is None:
            return 'keep'

        color_val = self.get_attr(color_elem, 'val')
        if not color_val or color_val == 'auto':
            return 'keep'

        # Optimization #5: Pre-compute color matches once
        is_red = is_color_match(color_val, self.RED, self.red_tolerance)
        is_green = is_color_match(color_val, self.GREEN, self.green_tolerance)
        is_blue = is_color_match(color_val, self.BLUE, self.blue_tolerance)

        # Early exit if no color match
        if not (is_red or is_green or is_blue):
            return 'keep'

        # Check structural elements once
        is_structural = (
            run.find(W_TAB) is not None or
            run.find(W_BR) is not None or
            run.find(W_CR) is not None
        )

        # Red -> delete (except structural: clean)
        if is_red:
            self.stats['red_deleted'] += 1
            return 'clean' if is_structural else 'delete'

        # Green
        if is_green:
            has_strike = self.has_strikethrough(rPr)

            # Explicit strike always wins
            if has_strike:
                self.stats['green_strike_deleted'] += 1
                return 'clean' if is_structural else 'delete'

            # Simulated strikethrough?
            if self.check_simulated_lines and self._has_simulated_strike(run, rPr):
                self.stats['green_strike_deleted'] += 1
                return 'clean' if is_structural else 'delete'

            # Otherwise -> clean formatting only
            self.stats['green_cleaned'] += 1
            return 'clean'

        # Blue -> clean formatting
        if is_blue:
            self.stats['blue_cleaned'] += 1
            return 'clean'

        return 'keep'

    def _has_simulated_strike(self, run, rPr):
        """Check if run has simulated strikethrough from green shapes."""
        if run in self.greenshape_striked_runs:
            return True

        if not self.greenshape_striked_runs:
            return False

        # Check neighborhood of runs in the paragraph
        para = self.get_ancestor(run, W_P)
        if para is None:
            return False

        runs = list(para.iter(W_R))
        try:
            idx = runs.index(run)
        except ValueError:
            return False

        # Check nearby runs
        start = max(0, idx - 4)
        end = min(len(runs), idx + 5)
        greenshape_runs = self.greenshape_striked_runs  # Local reference
        for j in range(start, end):
            if runs[j] in greenshape_runs:
                return True

        return False

    def highlight_run(self, run, highlight_color):
        """
        Apply a highlight color to a run for preview purposes.
        - For 'red': use a much lighter red via shading fill.
        - For others: use standard Word highlight colors.
        """
        rPr = run.find(W_RPR)
        if rPr is None:
            # Create rPr if it doesn't exist
            rPr = ET.Element(W_RPR)
            run.insert(0, rPr)

        # Remove existing highlight if present
        existing_highlight = rPr.find(W_HIGHLIGHT)
        if existing_highlight is not None:
            rPr.remove(existing_highlight)

        # Remove existing shading if present (we'll apply our own)
        existing_shd = rPr.find(W_SHD)
        if existing_shd is not None:
            rPr.remove(existing_shd)

        if highlight_color == 'red':
            # Use a much lighter red background via shading instead of the intense Word 'red' highlight
            shd_elem = ET.Element(W_SHD)
            shd_elem.set(f'{W_NS}val', 'clear')
            shd_elem.set(f'{W_NS}color', 'auto')
            # Light red fill (pastel)
            shd_elem.set(f'{W_NS}fill', 'FFC0C0')
            rPr.append(shd_elem)
        else:
            # Standard Word highlight for cyan/others
            highlight_elem = ET.Element(W_HIGHLIGHT)
            highlight_elem.set(f'{W_NS}val', highlight_color)
            rPr.append(highlight_elem)

    def clean_run(self, run):
        """Remove color, underline, bold, italic, strikethrough, shading, highlight."""
        rPr = run.find(W_RPR)
        if rPr is None:
            return

        # Use pre-computed set for O(1) lookup
        for child in list(rPr):
            if child.tag in CLEAN_TAGS:
                rPr.remove(child)

    def classify_shape_role(self, vshape):
        """Classify a VML shape as 'strike' or 'underline' based on style/top."""
        style = vshape.get('style') or ''
        top_pt = self.parse_top_from_style(style)
        if top_pt is not None and top_pt >= 8.0:
            return 'underline'
        return 'strike'

    def detect_green_shape_strikes(self, root):
        """
        Detect green drawn lines that act as strikethroughs.

        Optimized version using single-pass collection and deferred paragraph lookup.
        """
        self.greenshape_striked_runs = set()
        green_tolerance = self.green_tolerance
        GREEN = self.GREEN
        parent_map = self.parent_map

        # Collect paragraphs that contain green strike candidates
        strike_paragraphs = set()

        # --- Use XPath for lxml, fallback to iter for ElementTree ---
        if USING_LXML:
            alt_contents = root.xpath('.//mc:AlternateContent', namespaces=NAMESPACES)
        else:
            alt_contents = list(root.iter(MC_ALTCONTENT))

        # --- Single-pass collection of all colored elements in AlternateContent ---
        for alt in alt_contents:
            colors_found = []
            has_textbox = False
            vml_shapes = []

            for elem in alt.iter():
                tag = elem.tag

                # Check for textbox content (disqualifies as shape)
                if tag == W_TXBXCONTENT:
                    has_textbox = True
                    break

                # Collect DrawingML colors
                if tag == A_SRGBCLR:
                    color_val = elem.get('val')
                    if color_val:
                        colors_found.append(color_val.upper())

                # Collect VML shapes and their colors
                if tag in VML_SHAPE_TAGS_SET:
                    vml_shapes.append(elem)
                    fillcolor = (
                        elem.get('fillcolor') or
                        elem.get('strokecolor') or
                        elem.get('color')
                    )
                    if not fillcolor:
                        fill_child = elem.find('.//' + V_FILL)
                        if fill_child is not None:
                            fillcolor = fill_child.get('color') or fill_child.get('fillcolor')
                    if fillcolor:
                        colors_found.append(fillcolor.lstrip('#').upper())

            # Skip if this is a textbox
            if has_textbox:
                continue

            # Skip if no colors found
            if not colors_found:
                continue

            # Check if any color is green
            has_green = False
            for c in colors_found:
                if is_color_match(c, GREEN, green_tolerance):
                    has_green = True
                    break

            if not has_green:
                continue

            # Determine role from VML shapes
            role = 'strike'
            for vshape in vml_shapes:
                role = self.classify_shape_role(vshape)
                if role == 'underline':
                    break

            if role != 'strike':
                continue

            # Find parent paragraph (deferred until we know it's a valid candidate)
            cur = alt
            while cur is not None and cur.tag != W_P:
                cur = parent_map.get(cur)
            if cur is not None:
                strike_paragraphs.add(cur)

        # --- Check standalone VML shapes not wrapped in AlternateContent ---
        # Build a set of all AlternateContent elements for quick membership testing
        alt_set = set(alt_contents)

        for tag in VML_SHAPE_TAGS[:3]:  # line, shape, rect
            if USING_LXML:
                vml_shapes = root.xpath(f'.//v:{tag.split("}")[-1]}', namespaces=NAMESPACES)
            else:
                vml_shapes = list(root.iter(tag))

            for vshape in vml_shapes:
                # Check if inside an AlternateContent (already processed)
                cur = vshape
                inside_alt = False
                while cur in parent_map:
                    cur = parent_map[cur]
                    if cur in alt_set:
                        inside_alt = True
                        break

                if inside_alt:
                    continue

                # Check for textbox child
                if vshape.find('.//' + V_TEXTBOX) is not None:
                    continue

                # Get fill color
                fillcolor = (
                    vshape.get('fillcolor') or
                    vshape.get('strokecolor') or
                    vshape.get('color')
                )
                if not fillcolor:
                    fill_child = vshape.find('.//' + V_FILL)
                    if fill_child is not None:
                        fillcolor = fill_child.get('color') or fill_child.get('fillcolor')

                if not fillcolor:
                    continue

                if not is_color_match(fillcolor.lstrip('#').upper(), GREEN, green_tolerance):
                    continue

                role = self.classify_shape_role(vshape)
                if role != 'strike':
                    continue

                # Find parent paragraph
                cur = vshape
                while cur is not None and cur.tag != W_P:
                    cur = parent_map.get(cur)
                if cur is not None:
                    strike_paragraphs.add(cur)

        # --- Single pass through strike paragraphs to find green runs ---
        if not strike_paragraphs:
            return

        for para in strike_paragraphs:
            for run in para.iter(W_R):
                rPr = run.find(W_RPR)
                if rPr is None:
                    continue
                color_elem = rPr.find(W_COLOR)
                if color_elem is None:
                    continue
                color_val = self.get_attr(color_elem, 'val')
                if not color_val:
                    continue
                if not is_color_match(color_val, GREEN, green_tolerance):
                    continue
                if self.is_underlined(rPr):
                    continue
                if not self.run_has_visible_text(run):
                    continue
                self.greenshape_striked_runs.add(run)

    @staticmethod
    def parse_top_from_style(style):
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

    def remove_colored_shapes(self, root):
        """Remove colored graphical shapes (non-textbox shapes)."""
        to_remove = []
        parent_map = self.parent_map
        red_tol = self.red_tolerance
        green_tol = self.green_tolerance
        blue_tol = self.blue_tolerance
        RED, GREEN, BLUE = self.RED, self.GREEN, self.BLUE

        for alt in root.iter(MC_ALTCONTENT):
            should_remove = False

            # Check DrawingML colors
            for srgb in alt.iter(A_SRGBCLR):
                color_val = srgb.get('val')
                if color_val:
                    color_val = color_val.upper()
                    if (is_color_match(color_val, RED, red_tol) or
                        is_color_match(color_val, GREEN, green_tol) or
                        is_color_match(color_val, BLUE, blue_tol)):
                        if alt.find('.//' + W_TXBXCONTENT) is None:
                            should_remove = True
                            break

            if should_remove:
                to_remove.append(alt)
                continue

            # Check VML shapes
            for vml in alt.iter():
                if vml.tag in VML_SHAPE_TAGS:
                    fillcolor = vml.get('fillcolor') or vml.get('strokecolor') or vml.get('color')
                    if fillcolor:
                        fillcolor = fillcolor.lstrip('#').upper()
                        if (is_color_match(fillcolor, RED, red_tol) or
                            is_color_match(fillcolor, GREEN, green_tol) or
                            is_color_match(fillcolor, BLUE, blue_tol)):
                            if vml.find('.//' + V_TEXTBOX) is None:
                                should_remove = True
                                break

            if should_remove:
                to_remove.append(alt)

        # Batch removal
        for elem in to_remove:
            parent = parent_map.get(elem)
            if parent is not None:
                try:
                    parent.remove(elem)
                    self.stats['colored_shapes_removed'] += 1
                except ValueError:
                    pass

    def remove_empty_textboxes(self, root):
        """Remove completely empty text boxes."""
        parent_map = self.parent_map
        txbxs = list(root.iter(W_TXBXCONTENT))
        to_remove = []

        for tx in txbxs:
            if self.textbox_is_empty(tx):
                target = self.find_removal_target(tx)
                if target is not None and target not in to_remove:
                    to_remove.append(target)

        for elem in to_remove:
            parent = parent_map.get(elem)
            if parent is not None:
                try:
                    parent.remove(elem)
                    self.stats['textboxes_removed'] += 1
                except ValueError:
                    pass

    @staticmethod
    def textbox_is_empty(txbx_content):
        """Heuristic: no text, no drawings, no objects."""
        for t in txbx_content.iter(W_T):
            if t.text and t.text.strip():
                return False
        for _ in txbx_content.iter(W_DRAWING):
            return False
        for _ in txbx_content.iter(W_OBJECT):
            return False
        return True

    def find_removal_target(self, txbx_content):
        """Find the appropriate parent element to remove for an empty textbox."""
        current = txbx_content
        parent_map = self.parent_map
        while current in parent_map:
            parent = parent_map[current]
            if parent.tag == MC_ALTCONTENT:
                return parent
            if parent.tag == W_R:
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


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def make_output_docx(input_path):
    p = Path(input_path)
    return str(p.parent / f"(No Blacklines) {p.name}")


def make_output_pdf(docx_output):
    p = Path(docx_output)
    return str(p.with_suffix(".pdf"))


def make_preview_docx(input_path):
    """Create a path for preview docx in the same folder."""
    p = Path(input_path)
    return str(p.parent / f"(Preview) {p.name}")


# =============================================================================
# WORKER THREAD
# =============================================================================

class ProcessorWorker(QObject):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(list, dict)
    error = pyqtSignal(str)

    def __init__(self, files, mode,
                 red_tolerance, green_tolerance, blue_tolerance,
                 check_simulated,
                 red_color, green_color, blue_color,
                 parent=None):
        super().__init__(parent)
        self.files = files
        self.mode = mode
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


class PreviewWorker(QObject):
    """Worker for generating preview DOCX files."""
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(str)  # Path to preview DOCX
    error = pyqtSignal(str)

    def __init__(self, input_file,
                 red_tolerance, green_tolerance, blue_tolerance,
                 check_simulated,
                 red_color, green_color, blue_color,
                 parent=None):
        super().__init__(parent)
        self.input_file = input_file
        self.red_tolerance = red_tolerance
        self.green_tolerance = green_tolerance
        self.blue_tolerance = blue_tolerance
        self.check_simulated = check_simulated
        self.red_color = red_color
        self.green_color = green_color
        self.blue_color = blue_color

    def run(self):
        try:
            self.status.emit("Generating preview...")
            self.status.emit(f"Input file: {os.path.basename(self.input_file)}")
            self.progress.emit(5)

            # Create processor in preview mode
            processor = WordColorProcessor(
                red_tolerance=self.red_tolerance,
                green_tolerance=self.green_tolerance,
                blue_tolerance=self.blue_tolerance,
                check_simulated_lines=self.check_simulated,
                red_color=self.red_color,
                green_color=self.green_color,
                blue_color=self.blue_color,
                preview_mode=True
            )

            # Create preview DOCX with highlights (no PDF conversion)
            preview_docx = make_preview_docx(self.input_file)

            def progress_callback(msg):
                self.status.emit(msg)

            def progress_value_callback(val):
                # Scale 0-100 to 5-70 for the processing phase
                scaled = 5 + int(val * 0.65)
                self.progress.emit(scaled)

            self.status.emit("Creating highlighted document...")
            stats = processor.process_document(
                self.input_file, preview_docx,
                progress_callback=progress_callback,
                progress_value_callback=progress_value_callback
            )

            self.status.emit(
                f"Preview stats - Will delete: {stats['red_deleted'] + stats['green_strike_deleted']} runs, "
                f"Will clean: {stats['green_cleaned'] + stats['blue_cleaned']} runs"
            )

            # Finalize preview (DOCX only)
            self.progress.emit(95)
            self.progress.emit(100)
            self.status.emit("Preview ready!")
            self.finished.emit(preview_docx)

        except Exception as e:
            import traceback
            self.error.emit(f"{str(e)}\n{traceback.format_exc()}")


# =============================================================================
# GUI COMPONENTS
# =============================================================================

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


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Blackline Remover by Emil Ferdman (Optimized with Preview)")
        self.resize(900, 750)
        self.setMinimumSize(900, 750)
        self.setAcceptDrops(True)

        self.input_files = []
        self.red_color = "#FF0000"
        self.green_color = "#008000"
        self.blue_color = "#0000FF"

        self.worker_thread = None
        self.worker = None
        self.preview_thread = None
        self.preview_worker = None

        # Track preview files so we can clean them up on app close
        self.preview_files = []

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

        self.red_slider, self.red_swatch, self.red_hex_label = self.create_color_row(
            color_layout, "Red (delete all):", "red_color", 10
        )
        self.green_slider, self.green_swatch, self.green_hex_label = self.create_color_row(
            color_layout, "Green (delete if strikethrough):", "green_color", 10
        )
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

        # Preview legend
        legend_layout = QHBoxLayout()
        options_layout.addLayout(legend_layout)

        legend_label = QLabel("Preview Legend:")
        legend_label.setStyleSheet("font-weight: bold;")
        legend_layout.addWidget(legend_label)

        red_legend = QLabel("■ Light red background = Will be DELETED")
        red_legend.setStyleSheet("color: #CC0000; font-weight: bold;")
        legend_layout.addWidget(red_legend)

        blue_legend = QLabel("■ Cyan highlight = Will be CLEANED (formatting removed)")
        blue_legend.setStyleSheet("color: #0088AA; font-weight: bold;")
        legend_layout.addWidget(blue_legend)

        legend_layout.addStretch()

        # Progress group
        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout(progress_group)
        main_layout.addWidget(progress_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        progress_layout.addWidget(self.progress_bar)
        self.progress_label = QLabel("Ready")
        progress_layout.addWidget(self.progress_label)

        # Buttons
        buttons_layout = QHBoxLayout()
        main_layout.addLayout(buttons_layout)

        self.preview_btn = QPushButton("Preview Changes")
        self.preview_btn.setToolTip(
            "Generate a DOCX preview showing what will be deleted (light red background)\n"
            "and what will be cleaned (cyan highlight)"
        )
        self.preview_btn.clicked.connect(self.start_preview)
        buttons_layout.addWidget(self.preview_btn)

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

        if not DOCX2PDF_AVAILABLE:
            self.process_pdf_btn.setToolTip(
                "docx2pdf is not installed. PDF processing will not work until you install it:\n"
                "pip install docx2pdf"
            )

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
                hex_str = color.name().upper()
                setattr(self, color_attr_name, hex_str)
                swatch.set_color(hex_str)
                hex_label.setText(hex_str)

        swatch.clicked.connect(pick_color)

        return slider, swatch, hex_label

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

    def log(self, message):
        self.log_edit.append(message)
        self.log_edit.verticalScrollBar().setValue(
            self.log_edit.verticalScrollBar().maximum()
        )

    def set_buttons_enabled(self, enabled):
        """Enable or disable all action buttons."""
        self.preview_btn.setEnabled(enabled)
        self.process_docx_btn.setEnabled(enabled)
        self.process_pdf_btn.setEnabled(enabled)

    def start_preview(self):
        """Generate a preview DOCX showing what will be deleted/cleaned."""
        if not self.input_files:
            QMessageBox.warning(self, "No files", "Please select at least one .docx file.")
            return

        if len(self.input_files) > 1:
            QMessageBox.information(
                self,
                "Preview",
                "Preview will be generated for the first file only.\n"
                f"File: {os.path.basename(self.input_files[0])}"
            )

        self.log_edit.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText("Generating preview...")
        self.set_buttons_enabled(False)

        self.preview_thread = QThread()
        self.preview_worker = PreviewWorker(
            input_file=self.input_files[0],
            red_tolerance=self.red_slider.value(),
            green_tolerance=self.green_slider.value(),
            blue_tolerance=self.blue_slider.value(),
            check_simulated=self.simulated_checkbox.isChecked(),
            red_color=self.red_color,
            green_color=self.green_color,
            blue_color=self.blue_color,
        )
        self.preview_worker.moveToThread(self.preview_thread)

        self.preview_thread.started.connect(self.preview_worker.run)
        self.preview_worker.progress.connect(self.on_worker_progress)
        self.preview_worker.status.connect(self.on_worker_status)
        self.preview_worker.finished.connect(self.on_preview_finished)
        self.preview_worker.error.connect(self.on_worker_error)

        self.preview_worker.finished.connect(self.preview_thread.quit)
        self.preview_worker.error.connect(self.preview_thread.quit)
        self.preview_thread.finished.connect(self.preview_thread.deleteLater)

        self.preview_thread.start()

    def on_preview_finished(self, preview_path):
        """Handle preview completion."""
        self.log(f"Preview saved to: {preview_path}")
        self.progress_label.setText("Preview ready!")
        self.set_buttons_enabled(True)

        # Track preview for later cleanup and delete any older previews now
        self.cleanup_preview_files(exclude=preview_path)
        if preview_path not in self.preview_files:
            self.preview_files.append(preview_path)

        # Open the preview file
        self.log("Opening preview...")
        self.open_path(preview_path)

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
        self.set_buttons_enabled(False)

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

        self.set_buttons_enabled(True)

    def on_worker_error(self, message):
        self.log(f"ERROR: {message}")
        self.progress_label.setText("Error!")
        self.progress_bar.setValue(0)
        self.set_buttons_enabled(True)
        QMessageBox.critical(self, "Error", f"An error occurred:\n{message}")

    def open_path(self, path):
        """Open a file or folder with the OS default handler."""
        try:
            if sys.platform.startswith("darwin"):
                subprocess.Popen(["open", path])
            elif os.name == "nt":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            QMessageBox.critical(
                self,
                "Open Failed",
                f"Could not open:\n{path}\n\nError:\n{e}"
            )

    def cleanup_preview_files(self, exclude=None):
        """
        Attempt to delete any preview DOCX files we've created.
        - If exclude is provided, that path is kept (for the current active preview).
        - Files that can't be removed (e.g., still open/locked) are kept in the list
          so we can try again on app close.
        """
        remaining = []
        for p in self.preview_files:
            if exclude is not None and os.path.abspath(p) == os.path.abspath(exclude):
                remaining.append(p)
                continue
            try:
                if os.path.isfile(p):
                    os.remove(p)
            except OSError:
                # Likely still open/locked; keep it for another attempt later
                remaining.append(p)
        self.preview_files = remaining

    def closeEvent(self, event):
        """On app close, try to clean up any remaining preview DOCX files."""
        self.cleanup_preview_files()
        super().closeEvent(event)


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()