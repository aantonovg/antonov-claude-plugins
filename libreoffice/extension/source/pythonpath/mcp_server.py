"""
LibreOffice MCP Extension - MCP Server Module

This module implements an embedded MCP server that integrates with LibreOffice
via the UNO API, providing real-time document manipulation capabilities.
"""

import asyncio
import json
import logging
from typing import Dict, Any, Optional, List
from uno_bridge import UNOBridge

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class LibreOfficeMCPServer:
    """Embedded MCP server for LibreOffice plugin"""
    
    def __init__(self):
        """Initialize the MCP server"""
        self.uno_bridge = UNOBridge()
        self.tools = {}
        self._register_tools()
        logger.info("LibreOffice MCP Server initialized")
    
    def _register_tools(self):
        """Register all available MCP tools"""
        
        # Document creation tools
        self.tools["create_document_live"] = {
            "description": (
                "Create a new document in LibreOffice. "
                "visible=true (default) shows the window immediately. "
                "Set visible=false BEFORE running a large execute_batch of writes — "
                "on macOS, mutating a visible doc many times in a burst can deadlock "
                "the SolarMutex (worker thread mutates → main thread layout reflow → "
                "contention). Render hidden, then call show_window after the batch."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "doc_type": {
                        "type": "string",
                        "enum": ["writer", "calc", "impress", "draw"],
                        "default": "writer",
                    },
                    "visible": {
                        "type": "boolean",
                        "default": True,
                        "description": "False keeps the window hidden — use for batch writes; pair with show_window after.",
                    },
                },
            },
            "handler": self.create_document_live,
        }

        self.tools["show_window"] = {
            "description": (
                "Make the active document's window visible. Use after a batch of "
                "writes on a doc created with visible=false."
            ),
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.show_window(),
        }

        self.tools["hide_window"] = {
            "description": (
                "Hide the active document's window. Use before a burst of "
                "clone_page_style / clone_paragraph_style / set_page_style_props "
                "/ set_paragraph_style_props on macOS to avoid SolarMutex "
                "deadlock with AppKit paint cycles. Pair with show_window "
                "afterwards. Returns was_visible so caller can restore prior "
                "state. execute_batch with auto_hide='auto' (default) does "
                "this automatically when heavy ops are detected."
            ),
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.hide_window(),
        }

        self.tools["shutdown_application"] = {
            "description": (
                "Cleanly terminate LibreOffice via Desktop.terminate(). Use for "
                "hot-reload instead of `pkill -9` — clean terminate skips the "
                "Document-Recovery dialog on next launch. If unsaved docs exist, "
                "returns terminated=False; pass force=True to discard unsaved "
                "edits before terminating (DESTRUCTIVE)."
            ),
            "parameters": {"type": "object", "properties": {
                "force": {"type": "boolean", "default": False},
            }},
            "handler": lambda force=False: self.uno_bridge.shutdown_application(force),
        }
        
        # Text manipulation tools
        self.tools["insert_text_live"] = {
            "description": (
                "Insert text into the active Writer document. '\\n' becomes a real "
                "paragraph break. `position` controls WHERE to insert: "
                "'end' (default — append at end of body, safe for batch generation), "
                "'cursor' (use the live view-cursor — note: select_range moves it, "
                "so this can land at a previous selection), or an int char offset "
                "from doc start. After inserting at 'end', the new text is the LAST "
                "paragraph(s) — pair with apply_paragraph_style(target='last') to "
                "style the last inserted paragraph."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "text": {"type": "string", "description": "Text to insert (\\n = paragraph break)"},
                    "position": {
                        "description": "'end' (default), 'cursor', or int char offset from doc start",
                    },
                },
                "required": ["text"],
            },
            "handler": self.insert_text_live,
        }
        
        # Document info tools
        self.tools["get_document_info_live"] = {
            "description": "Get information about the currently active document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_document_info_live
        }
        
        # Text formatting tools
        self.tools["format_text_live"] = {
            "description": (
                "Apply character formatting to a range. Pass start/end (char offsets, "
                "end-exclusive) to format an explicit range — preferred for batch ops. "
                "Without start/end, formats the current selection."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "bold": {"type": "boolean"},
                    "italic": {"type": "boolean"},
                    "underline": {"type": "boolean"},
                    "font_size": {"type": "number", "description": "Font size in points"},
                    "font_name": {"type": "string", "description": "Font family name"},
                    "kerning": {"type": "integer", "description": "Per-char extra horizontal spacing in 1/100 mm (CharKerning)"},
                    "scale_width": {"type": "integer", "description": "Horizontal scale percent (CharScaleWidth, 100 = normal)"},
                    "start": {"type": "integer", "description": "Char offset (incl.)"},
                    "end": {"type": "integer", "description": "Char offset (excl.)"},
                },
            },
            "handler": self.format_text_live,
        }
        
        # NOTE: save_document_live / export_document_live were removed —
        # UNO save (doc.store / storeToURL / .uno:Save) blocks forever on
        # macOS Sequoia from the background HTTP-server thread (UI-thread
        # barrier). User must press Cmd+S in the LibreOffice window.

        # Content reading tools
        self.tools["get_text_content_live"] = {
            "description": "Get the text content of the currently active document",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.get_text_content_live
        }
        
        # Document list tools
        self.tools["list_open_documents"] = {
            "description": "List all currently open documents in LibreOffice",
            "parameters": {
                "type": "object",
                "properties": {}
            },
            "handler": self.list_open_documents
        }
        
        # ---- Extended editing tools (Writer) ---------------------------

        self.tools["set_text_color"] = {
            "description": "Set font color. Pass start/end for explicit char range (preferred for batch). Without them — uses current selection. Color: hex '#FF0000' or int.",
            "parameters": {
                "type": "object",
                "properties": {
                    "color": {"type": "string", "description": "Hex color, e.g. '#FF0000'"},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
                "required": ["color"],
            },
            "handler": lambda color, start=None, end=None: self.uno_bridge.set_text_color(color, start, end),
        }

        self.tools["set_background_color"] = {
            "description": "Set character background (highlight) color. Pass start/end for explicit char range.",
            "parameters": {
                "type": "object",
                "properties": {
                    "color": {"type": "string", "description": "Hex color, e.g. '#FFFF00'. Use -1 for 'no fill'."},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
                "required": ["color"],
            },
            "handler": lambda color, start=None, end=None: self.uno_bridge.set_background_color(color, start, end),
        }

        self.tools["set_paragraph_alignment"] = {
            "description": "Set paragraph alignment. Pass start/end to target paragraphs by char range; otherwise uses selection / view cursor.",
            "parameters": {
                "type": "object",
                "properties": {
                    "alignment": {"type": "string", "enum": ["left", "center", "right", "justify"]},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
                "required": ["alignment"],
            },
            "handler": lambda alignment, start=None, end=None: self.uno_bridge.set_paragraph_alignment(alignment, start, end),
        }

        self.tools["set_paragraph_indent"] = {
            "description": "Set paragraph indents (mm). Pass start/end for explicit range. Omit a field to leave it unchanged.",
            "parameters": {
                "type": "object",
                "properties": {
                    "left_mm": {"type": "number"},
                    "right_mm": {"type": "number"},
                    "first_line_mm": {"type": "number"},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
            },
            "handler": lambda **kw: self.uno_bridge.set_paragraph_indent(**kw),
        }

        self.tools["set_paragraph_spacing"] = {
            "description": (
                "Set paragraph above/below spacing (ParaTopMargin / ParaBottomMargin) "
                "and ParaContextMargin (when True, adjacent paragraphs of the same "
                "style collapse top/bottom — controls whether spacings stack). "
                "Values in mm. Read context_margin via get_paragraphs and replicate it; "
                "without it, target may add visible gaps between same-style paragraphs."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "top_mm": {"type": "number"},
                    "bottom_mm": {"type": "number"},
                    "context_margin": {"type": "boolean"},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
            },
            "handler": lambda **kw: self.uno_bridge.set_paragraph_spacing(**kw),
        }

        self.tools["set_paragraph_tabs"] = {
            "description": (
                "Replace ParaTabStops on a paragraph range. stops is a list of "
                "{position_mm, alignment: 'left'|'center'|'right'|'decimal', "
                "fill_char, decimal_char}. Use this to reproduce a right-aligned "
                "tab in lines like 'Город ... \\t ... дата' where the date is "
                "pinned to the right margin via a right-tab."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "stops": {"type": "array", "items": {"type": "object"}},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
                "required": ["stops"],
            },
            "handler": lambda stops, start=None, end=None:
                self.uno_bridge.set_paragraph_tabs(stops, start, end),
        }

        self.tools["set_footer_page_number"] = {
            "description": (
                "Replace the footer of a page-style with a single PageNumber "
                "field at the given alignment ('left'|'center'|'right'). "
                "Use to replicate page numbers across every page that uses "
                "this style — much simpler than anchoring TextFrames per page."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "page_style": {"type": "string", "default": "Default Page Style"},
                    "alignment": {"type": "string", "enum": ["left", "center", "right"], "default": "center"},
                    "font_size": {"type": "number"},
                },
            },
            "handler": lambda **kw: self.uno_bridge.set_footer_page_number(**kw),
        }

        self.tools["set_paragraph_text_flow"] = {
            "description": (
                "Set per-paragraph text-flow properties (widows, orphans, "
                "keep_together, split_paragraph, keep_with_next). These "
                "control how the paragraph breaks across pages and prevent "
                "single-word/single-line spillover onto the next page. "
                "Word→ODT often sets these per paragraph — replicate via "
                "get_paragraphs (which now exposes them) → this setter."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "widows": {"type": "integer"},
                    "orphans": {"type": "integer"},
                    "keep_together": {"type": "boolean"},
                    "split_paragraph": {"type": "boolean"},
                    "keep_with_next": {"type": "boolean"},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
            },
            "handler": lambda **kw: self.uno_bridge.set_paragraph_text_flow(**kw),
        }

        self.tools["set_paragraph_breaks"] = {
            "description": (
                "Set BreakType / PageDescName / PageNumberOffset on paragraphs. "
                "break_type: int 0..6 or name (NONE, COLUMN_BEFORE, COLUMN_AFTER, "
                "COLUMN_BOTH, PAGE_BEFORE, PAGE_AFTER, PAGE_BOTH). PAGE_BEFORE=4 "
                "forces a page break before the paragraph — needed to replicate "
                "the visual page layout of Word-imported documents (which encode "
                "such breaks as a paragraph property, NOT as a control character). "
                "page_desc_name optionally switches page-style at the break."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "break_type": {"type": ["integer", "string"]},
                    "page_desc_name": {"type": "string"},
                    "page_number_offset": {"type": "integer"},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
            },
            "handler": lambda **kw: self.uno_bridge.set_paragraph_breaks(**kw),
        }

        self.tools["set_line_spacing"] = {
            "description": "Set line spacing. proportional: value=100 single, 150=1.5x, 200=double. fix/minimum/leading: value in mm. Pass start/end for explicit range.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mode": {"type": "string", "enum": ["proportional", "minimum", "leading", "fix"], "default": "proportional"},
                    "value": {"type": "number"},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
                "required": ["value"],
            },
            "handler": lambda mode="proportional", value=100, start=None, end=None:
                self.uno_bridge.set_line_spacing(mode, value, start, end),
        }

        self.tools["apply_paragraph_style"] = {
            "description": (
                "Apply a paragraph style by name (e.g. 'Heading 1', 'Heading 2', "
                "'Quotations', 'Default Paragraph Style'). "
                "TARGET SELECTION: pass target='last' to style the LAST paragraph "
                "(use after insert_text(... position='end') — this is the bug-free "
                "pattern for batch generation). Or pass start/end to style every "
                "paragraph touching that char range. Without either, uses the "
                "current selection / view cursor (legacy; depends on UI state)."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "style_name": {"type": "string"},
                    "target": {"type": "string", "enum": ["last"], "description": "'last' = style last paragraph"},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
                "required": ["style_name"],
            },
            "handler": lambda style_name, target=None, start=None, end=None:
                self.uno_bridge.apply_paragraph_style(style_name, start, end, target),
        }

        self.tools["find_and_replace"] = {
            "description": "Replace all matches of `search` with `replace` in the active document. Set regex=true for regex search.",
            "parameters": {
                "type": "object",
                "properties": {
                    "search": {"type": "string"},
                    "replace": {"type": "string", "default": ""},
                    "regex": {"type": "boolean", "default": False},
                    "case_sensitive": {"type": "boolean", "default": False},
                },
                "required": ["search"],
            },
            "handler": lambda search, replace="", regex=False, case_sensitive=False:
                self.uno_bridge.find_and_replace(search, replace, regex, case_sensitive),
        }

        self.tools["delete_range"] = {
            "description": "Delete characters in the active Writer document between `start` (inclusive) and `end` (exclusive), counted from the beginning of the document body.",
            "parameters": {
                "type": "object",
                "properties": {
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
                "required": ["start", "end"],
            },
            "handler": lambda start, end: self.uno_bridge.delete_range(start, end),
        }

        self.tools["select_range"] = {
            "description": "Programmatically select characters [start, end) in the active Writer document so subsequent format_text_live / set_text_color / etc. apply to that range.",
            "parameters": {
                "type": "object",
                "properties": {
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
                "required": ["start", "end"],
            },
            "handler": lambda start, end: self.uno_bridge.select_range(start, end),
        }

        # ---- Read-only inspection tools (Writer) -----------------------

        self.tools["get_paragraphs"] = {
            "description": "List paragraphs of the active Writer document with their absolute char-range, style, alignment, indents and a text preview. Use this to plan what to format.",
            "parameters": {
                "type": "object",
                "properties": {
                    "start": {"type": "integer", "default": 0, "description": "Index of the first paragraph to return"},
                    "count": {"type": "integer", "description": "Max paragraphs to return (omit for all)"},
                    "include_format": {"type": "boolean", "default": True},
                    "preview_chars": {"type": "integer", "default": 80},
                },
            },
            "handler": lambda **kw: self.uno_bridge.get_paragraphs(**kw),
        }

        self.tools["get_paragraph_format_at"] = {
            "description": "Get full paragraph format at a given char position (style, alignment, indents, line spacing).",
            "parameters": {
                "type": "object",
                "properties": {"position": {"type": "integer"}},
                "required": ["position"],
            },
            "handler": lambda position: self.uno_bridge.get_paragraph_format_at(position),
        }

        self.tools["get_outline"] = {
            "description": "Return the document outline — only headings (paragraphs with OutlineLevel>0 or 'Heading*'/'Title' style). Cheap way to build a TOC or section map without reading the whole body.",
            "parameters": {
                "type": "object",
                "properties": {
                    "max_level": {"type": "integer", "default": 10, "description": "Skip headings deeper than this level"},
                },
            },
            "handler": lambda **kw: self.uno_bridge.get_outline(**kw),
        }

        self.tools["get_paragraphs_with_runs"] = {
            "description": "Like get_paragraphs but also returns inline character runs per paragraph (text portions with uniform formatting): font, size, bold/italic/underline/strike, color, hyperlink URL, char style, char_posture (NONE/OBLIQUE/ITALIC), and per-character spacing (`kerning` in 1/100 mm — surfaces only when non-zero; Word imports often set this on space portions between bold names to widen the line; without replicating, justify-wraps will diverge), `scale_width` (CharScaleWidth percent — only when ≠100). Use for faithful diff/replication when inline formatting matters.",
            "parameters": {
                "type": "object",
                "properties": {
                    "start": {"type": "integer", "default": 0},
                    "count": {"type": "integer", "description": "Max paragraphs to return (omit for all)"},
                    "include_para_format": {"type": "boolean", "default": True},
                },
            },
            "handler": lambda **kw: self.uno_bridge.get_paragraphs_with_runs(**kw),
        }

        self.tools["list_body_elements"] = {
            "description": (
                "Return the ordered sequence of paragraphs AND tables in the "
                "document body. Each entry has 'kind' = 'paragraph' or 'table'. "
                "Tables include 'name', 'rows', 'columns', and "
                "'after_paragraph_index' — the paragraph the table sits "
                "immediately after. Use this to faithfully replicate documents "
                "where tables interleave with text — get_paragraphs and "
                "get_paragraphs_with_runs silently skip tables, so a batch "
                "rebuilder loses their positions otherwise."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "start": {"type": "integer", "default": 0},
                    "count": {"type": "integer"},
                    "preview_chars": {"type": "integer", "default": 60},
                },
            },
            "handler": lambda **kw: self.uno_bridge.list_body_elements(**kw),
        }

        self.tools["get_page_layout"] = {
            "description": (
                "Map every body element (paragraph, table, table row, frame) "
                "to its actual page using ViewCursor.getPage(). Each element "
                "gets start_page+end_page; if they differ, the element is "
                "split between pages (paragraph wrapping multiple lines, "
                "table row split with Split=True, whole table spanning "
                "pages). Returns ordered 'elements' AND inverse 'pages[]' — "
                "for each page the list of paragraphs/tables/frames touching "
                "it with is_start/is_end/spans_pages markers. Also returns "
                "'table_groups[]': adjacent tables with matching column "
                "count and only empty paragraphs between them are grouped "
                "as one logical table — ODT models a split table as two "
                "TextTable objects between which the layout engine "
                "redistributes rows when content shifts, so this surfaces "
                "what the user perceives as 'one table broken between pages'. "
                "Each table in a group gets group_id/group_position/"
                "group_size and each row gets group_row_index (cumulative "
                "across the group). This is the only tool that surfaces "
                "actual page breaks (UNO does not expose them statically — "
                "they are layout-computed). REQUIRES the document to be "
                "visible: hidden docs return 0 from getPage(). Each lookup "
                "does ctrl.select(range) so a visible window will flicker "
                "briefly."
            ),
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.get_page_layout(),
        }

        self.tools["get_character_format"] = {
            "description": "Read character formatting (font, size, bold/italic/underline, color, bg color, kerning, scale_width) over a char range. `kerning` is per-char extra horizontal spacing in 1/100 mm — Word imports often set non-zero kerning on individual portions (e.g. spaces between bold names) to widen justify rendering. Pass only `start` to read a single character.",
            "parameters": {
                "type": "object",
                "properties": {
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
                "required": ["start"],
            },
            "handler": lambda start, end=None: self.uno_bridge.get_character_format(start, end),
        }

        self.tools["list_paragraph_styles"] = {
            "description": "List all paragraph styles available in the active document (e.g. 'Heading 1', 'Quotations').",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.list_paragraph_styles(),
        }

        self.tools["apply_numbering"] = {
            "description": (
                "Configure auto-numbering on a paragraph (or paragraphs in a "
                "char range). Use this when apply_paragraph_style attaches a "
                "style by name but the resulting effective_label is empty — "
                "the style exists in target doc but its numbering rules are "
                "not configured. Pair with clone_numbering_rule first if the "
                "rule itself doesn't exist in target. "
                "level: 0 = top, 1 = sub, etc. (UNO 0-indexed). "
                "rule_name: NumberingStyle to attach (omit to keep current). "
                "restart=True + start_value=N to begin a new counter."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "level": {"type": "integer", "default": 0},
                    "rule_name": {"type": "string"},
                    "restart": {"type": "boolean", "default": False},
                    "start_value": {"type": "integer"},
                    "is_number": {"type": "boolean", "default": True},
                    "target": {"type": "string", "enum": ["last"]},
                    "start": {"type": "integer"},
                    "end": {"type": "integer"},
                },
            },
            "handler": lambda **kw: self.uno_bridge.apply_numbering(**kw),
        }

        self.tools["list_numbering_styles"] = {
            "description": "List NumberingStyle names available in the active doc. Use to discover what auto-numbering rules exist before apply_numbering / clone_numbering_rule.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.list_numbering_styles(),
        }


        self.tools["list_text_frames"] = {
            "description": "List TextFrames (free-floating boxes) in a doc with text, size, position, anchor and any TextFields inside them. Page numbers visible at page bottoms are often inside TextFrames anchored to body paragraphs, NOT in the page-style's footer slot — so list_text_frames reveals what list_page_styles + dump_page_style_footer cannot.",
            "parameters": {
                "type": "object",
                "properties": {
                    "source_path": {"type": "string"},
                    "doc_title": {"type": "string"},
                },
            },
            "handler": lambda source_path=None, doc_title=None:
                self.uno_bridge.list_text_frames(source_path=source_path, doc_title=doc_title),
        }

        self.tools["list_text_fields"] = {
            "description": "List all TextFields in a doc (body + headers/footers + frames) with their service names, presentations, and anchor text. Reveals PageNumber/Date/etc. fields invisible to getString() — useful when reproducing footers that show '1' but FooterText.getString() returns empty.",
            "parameters": {
                "type": "object",
                "properties": {
                    "source_path": {"type": "string"},
                    "doc_title": {"type": "string"},
                },
            },
            "handler": lambda source_path=None, doc_title=None:
                self.uno_bridge.list_text_fields(source_path=source_path, doc_title=doc_title),
        }

        self.tools["dump_char_style"] = {
            "description": "Read a character style's properties (font/size/weight/etc) from any open doc. Compare label-rendering chars styles between source and target.",
            "parameters": {
                "type": "object",
                "properties": {
                    "style_name": {"type": "string"},
                    "source_path": {"type": "string"},
                    "doc_title": {"type": "string"},
                },
                "required": ["style_name"],
            },
            "handler": lambda style_name, source_path=None, doc_title=None:
                self.uno_bridge.dump_char_style(style_name=style_name, source_path=source_path, doc_title=doc_title),
        }

        self.tools["dump_doc_paragraph"] = {
            "description": "Read a paragraph's full props + numbering-rule level props as JSON-safe primitives from any open doc by source_path or doc_title. Use to diff source vs target without switching active document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "source_path": {"type": "string"},
                    "doc_title": {"type": "string"},
                    "paragraph_index": {"type": "integer", "default": 0},
                },
            },
            "handler": lambda source_path=None, doc_title=None, paragraph_index=0:
                self.uno_bridge.dump_doc_paragraph(source_path=source_path, doc_title=doc_title, paragraph_index=paragraph_index),
        }

        self.tools["clone_numbering_rule"] = {
            "description": (
                "Copy a NumberingStyle (with all its level shapes, prefixes, "
                "separators) from a currently-open source doc into the active "
                "doc. Required when the source doc uses a custom numbering rule "
                "that doesn't exist in the freshly-created target. After cloning, "
                "call apply_numbering(rule_name=...) on each paragraph that "
                "should follow the rule. The source doc must be already open "
                "(use open_document_live first)."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "source_path": {"type": "string", "description": "Filesystem path to the open source doc."},
                    "rule_name": {"type": "string"},
                    "target_name": {"type": "string", "description": "Name to use in target doc (defaults to rule_name)."},
                },
                "required": ["source_path", "rule_name"],
            },
            "handler": lambda source_path, rule_name, target_name=None:
                self.uno_bridge.clone_numbering_rule(source_path, rule_name, target_name),
        }

        self.tools["clone_paragraph_style"] = {
            "description": (
                "Copy a ParagraphStyle (font, size, weight, alignment, indents, "
                "spacing, line spacing, tab stops, outline level, parent) from a "
                "currently-open source doc into the active doc. Use when the "
                "source's paragraph-style name exists in target by name only "
                "(e.g. fresh LO 'Heading 1' has different defaults than a Word "
                "import) — without cloning, apply_paragraph_style inherits target "
                "defaults. Source doc must be open (use open_document_live first)."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "source_path": {"type": "string"},
                    "style_name": {"type": "string"},
                    "target_name": {"type": "string", "description": "Name in target doc (defaults to style_name)."},
                    "overwrite": {"type": "boolean", "default": True},
                },
                "required": ["source_path", "style_name"],
            },
            "handler": lambda source_path, style_name, target_name=None, overwrite=True:
                self.uno_bridge.clone_paragraph_style(source_path, style_name, target_name, overwrite),
        }

        self.tools["clone_page_style"] = {
            "description": (
                "Copy a PageStyle (page size, orientation, all 4 margins, "
                "header/footer enabled+text+heights+margins, columns, footnote "
                "area, borders, background) from a currently-open source doc "
                "into the active doc. If source_style is omitted, uses the "
                "page style of source's first paragraph. Source doc must be open."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "source_path": {"type": "string"},
                    "source_style": {"type": "string"},
                    "target_style": {"type": "string", "default": "Default Page Style"},
                },
                "required": ["source_path"],
            },
            "handler": lambda source_path, source_style=None, target_style="Default Page Style":
                self.uno_bridge.clone_page_style(source_path, source_style, target_style),
        }

        self.tools["get_paragraph_style_def"] = {
            "description": (
                "Full paragraph-style snapshot: font_name, font_size, bold, italic, "
                "underline, color, char_word_mode, alignment, left/right/first_line/"
                "top/bottom margins (mm), context_margin, line_spacing {mode, value}, "
                "tab_stops, outline_level, parent, follow, keep_together, "
                "split_paragraph, orphans, widows, kerning (CharKerning, 1/100 mm; "
                "only when ≠0), scale_width (CharScaleWidth %; only when ≠100). "
                "Pair with set_paragraph_style_props to write any subset back."
            ),
            "parameters": {
                "type": "object",
                "properties": {"style_name": {"type": "string"}},
                "required": ["style_name"],
            },
            "handler": lambda style_name: self.uno_bridge.get_paragraph_style_def(style_name),
        }

        self.tools["set_paragraph_style_props"] = {
            "description": (
                "Symmetric writer for get_paragraph_style_def — modifies the style "
                "in place (propagates to every paragraph using it). Accepts any "
                "subset of: font_name, font_size, bold, italic, underline, color "
                "('#RRGGBB'), char_word_mode, alignment ('left'|'right'|'justify'|"
                "'center'), left_mm, right_mm, first_line_mm, top_mm, bottom_mm, "
                "context_margin, line_spacing ({mode, value}), tab_stops "
                "(list of {position_mm, alignment, fill_char, decimal_char}), "
                "outline_level, keep_together, split_paragraph, orphans, widows, "
                "kerning (1/100 mm), scale_width (%), parent, follow."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "style_name": {"type": "string"},
                    "font_name": {"type": "string"},
                    "font_size": {"type": "number"},
                    "bold": {"type": "boolean"},
                    "italic": {"type": "boolean"},
                    "underline": {"type": "boolean"},
                    "color": {"type": "string"},
                    "char_word_mode": {"type": "boolean"},
                    "alignment": {"type": "string"},
                    "left_mm": {"type": "number"},
                    "right_mm": {"type": "number"},
                    "first_line_mm": {"type": "number"},
                    "top_mm": {"type": "number"},
                    "bottom_mm": {"type": "number"},
                    "context_margin": {"type": "boolean"},
                    "line_spacing": {"type": "object"},
                    "tab_stops": {"type": "array"},
                    "outline_level": {"type": "integer"},
                    "keep_together": {"type": "boolean"},
                    "split_paragraph": {"type": "boolean"},
                    "orphans": {"type": "integer"},
                    "widows": {"type": "integer"},
                    "kerning": {"type": "integer", "description": "Per-char extra horizontal spacing in 1/100 mm"},
                    "scale_width": {"type": "integer", "description": "Horizontal char scale percent (100=normal)"},
                    "parent": {"type": "string"},
                    "follow": {"type": "string"},
                },
                "required": ["style_name"],
            },
            "handler": lambda **kw: self.uno_bridge.set_paragraph_style_props(**kw),
        }

        self.tools["set_page_style_props"] = {
            "description": (
                "Symmetric writer for get_page_info — modifies a page style. "
                "Accepts any subset of: page_width_mm, page_height_mm, "
                "orientation ('portrait'|'landscape'), top/bottom/left/right_margin_mm, "
                "header_enabled, header_height_mm, header_body_distance_mm, "
                "header_left/right_margin_mm, header_text, "
                "footer_enabled, footer_height_mm, footer_body_distance_mm, "
                "footer_left/right_margin_mm, footer_text. Header/footer text and "
                "margin writes only take effect when the slot is enabled — pass "
                "header_enabled=True alongside header_text in the same call."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "page_style": {"type": "string", "default": "Default Page Style"},
                    "page_width_mm": {"type": "number"},
                    "page_height_mm": {"type": "number"},
                    "orientation": {"type": "string"},
                    "top_margin_mm": {"type": "number"},
                    "bottom_margin_mm": {"type": "number"},
                    "left_margin_mm": {"type": "number"},
                    "right_margin_mm": {"type": "number"},
                    "header_enabled": {"type": "boolean"},
                    "header_height_mm": {"type": "number"},
                    "header_body_distance_mm": {"type": "number"},
                    "header_left_margin_mm": {"type": "number"},
                    "header_right_margin_mm": {"type": "number"},
                    "header_text": {"type": "string"},
                    "footer_enabled": {"type": "boolean"},
                    "footer_height_mm": {"type": "number"},
                    "footer_body_distance_mm": {"type": "number"},
                    "footer_left_margin_mm": {"type": "number"},
                    "footer_right_margin_mm": {"type": "number"},
                    "footer_text": {"type": "string"},
                },
            },
            "handler": lambda **kw: self.uno_bridge.set_page_style_props(**kw),
        }

        self.tools["list_character_styles"] = {
            "description": "List all character styles available in the active document.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.list_character_styles(),
        }

        self.tools["find_all"] = {
            "description": "Find all matches of `search` in the active document and return their absolute [start, end) char positions. Does NOT modify the document.",
            "parameters": {
                "type": "object",
                "properties": {
                    "search": {"type": "string"},
                    "regex": {"type": "boolean", "default": False},
                    "case_sensitive": {"type": "boolean", "default": False},
                    "max_results": {"type": "integer", "default": 200},
                },
                "required": ["search"],
            },
            "handler": lambda search, regex=False, case_sensitive=False, max_results=200:
                self.uno_bridge.find_all(search, regex, case_sensitive, max_results),
        }

        self.tools["get_page_info"] = {
            "description": (
                "Returns full page metrics of the active Writer doc: page_count, "
                "current_page, page_width_mm/page_height_mm, top/bottom/left/right "
                "margins in mm, orientation, page_style name, header/footer enabled "
                "and dimensions, column_count. If page_style is omitted, picks the "
                "PageDescName of the first paragraph — important for Word imports "
                "where page1 uses a Master-Page style (e.g. 'MP0') that differs from "
                "'Default Page Style'/'Standard'. Use BEFORE replicating layout into "
                "a fresh target — agent must read source page margins and call "
                "set_page_margins on the target, otherwise text wraps differently."
            ),
            "parameters": {"type": "object", "properties": {
                "page_style": {"type": "string", "description": "Defaults to first paragraph's PageDescName."},
            }},
            "handler": lambda page_style=None: self.uno_bridge.get_page_info(page_style),
        }

        self.tools["set_page_margins"] = {
            "description": (
                "Set page margins (mm) on a page style. Only fields you pass are changed. "
                "Affects every paragraph using that page style. Pair with get_page_info to "
                "copy source layout into a target doc."
            ),
            "parameters": {"type": "object", "properties": {
                "top_mm": {"type": "number"},
                "bottom_mm": {"type": "number"},
                "left_mm": {"type": "number"},
                "right_mm": {"type": "number"},
                "page_style": {"type": "string", "default": "Default Page Style"},
            }},
            "handler": lambda top_mm=None, bottom_mm=None, left_mm=None, right_mm=None,
                              page_style="Default Page Style":
                self.uno_bridge.set_page_margins(top_mm, bottom_mm, left_mm, right_mm, page_style),
        }

        # ---- Open / Recent ---------------------------------------------

        self.tools["open_document_live"] = {
            "description": "Open an existing document on disk in LibreOffice and keep it open. Path can be absolute filesystem path or file:// URL.",
            "parameters": {"type": "object", "properties": {
                "path": {"type": "string"},
                "readonly": {"type": "boolean", "default": False},
            }, "required": ["path"]},
            "handler": lambda path, readonly=False: self.uno_bridge.open_document_live(path, readonly),
        }
        self.tools["list_recent_documents"] = {
            "description": "List documents from LibreOffice's Recent Documents (File → Recent) with their URLs and titles.",
            "parameters": {"type": "object", "properties": {
                "max_items": {"type": "integer", "default": 25}
            }},
            "handler": lambda max_items=25: self.uno_bridge.list_recent_documents(max_items),
        }
        self.tools["open_recent_document"] = {
            "description": "Open one of the Recent Documents by its index (0 = most recent).",
            "parameters": {"type": "object", "properties": {
                "index": {"type": "integer", "default": 0},
                "readonly": {"type": "boolean", "default": False},
            }},
            "handler": lambda index=0, readonly=False: self.uno_bridge.open_recent_document(index, readonly),
        }

        # ---- Document inspection / metadata ----------------------------

        self.tools["get_document_metadata"] = {
            "description": "Read document properties: title, subject, author, keywords, creation/modification dates, generator, etc.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.get_document_metadata(),
        }
        self.tools["set_document_metadata"] = {
            "description": "Update document properties (any subset).",
            "parameters": {"type": "object", "properties": {
                "title": {"type": "string"}, "subject": {"type": "string"},
                "author": {"type": "string"}, "description": {"type": "string"},
                "keywords": {"type": "array", "items": {"type": "string"}},
            }},
            "handler": lambda **kw: self.uno_bridge.set_document_metadata(**kw),
        }
        self.tools["get_document_summary"] = {
            "description": "One-shot overview: counts of paragraphs/pages/tables/images/sections/bookmarks/fields/annotations/hyperlinks plus title and char/word counts.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.get_document_summary(),
        }
        self.tools["get_text_at"] = {
            "description": "Return text in [start, end) without modifying anything.",
            "parameters": {"type": "object", "properties": {
                "start": {"type": "integer"}, "end": {"type": "integer"},
            }, "required": ["start", "end"]},
            "handler": lambda start, end: self.uno_bridge.get_text_at(start, end),
        }
        self.tools["get_selection"] = {
            "description": "Return text the user has currently selected in the editor (with approx char positions).",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.get_selection(),
        }

        # ---- Bookmarks --------------------------------------------------

        self.tools["list_bookmarks"] = {
            "description": "List bookmarks in the active document.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.list_bookmarks(),
        }
        self.tools["add_bookmark"] = {
            "description": "Add a bookmark at position `start` (or covering [start, end) if end is given).",
            "parameters": {"type": "object", "properties": {
                "name": {"type": "string"},
                "start": {"type": "integer"},
                "end": {"type": "integer"},
            }, "required": ["name", "start"]},
            "handler": lambda name, start, end=None: self.uno_bridge.add_bookmark(name, start, end),
        }
        self.tools["remove_bookmark"] = {
            "description": "Remove the bookmark with the given name.",
            "parameters": {"type": "object", "properties": {"name": {"type": "string"}},
                           "required": ["name"]},
            "handler": lambda name: self.uno_bridge.remove_bookmark(name),
        }

        # ---- Hyperlinks -------------------------------------------------

        self.tools["list_hyperlinks"] = {
            "description": "List all hyperlinks in the active document with their text and target URLs.",
            "parameters": {"type": "object", "properties": {"max_items": {"type": "integer", "default": 200}}},
            "handler": lambda max_items=200: self.uno_bridge.list_hyperlinks(max_items),
        }
        self.tools["add_hyperlink"] = {
            "description": "Turn characters [start, end) into a hyperlink pointing to `url`.",
            "parameters": {"type": "object", "properties": {
                "start": {"type": "integer"}, "end": {"type": "integer"},
                "url": {"type": "string"}, "target": {"type": "string", "default": ""},
            }, "required": ["start", "end", "url"]},
            "handler": lambda start, end, url, target="": self.uno_bridge.add_hyperlink(start, end, url, target),
        }

        # ---- Comments / annotations ------------------------------------

        self.tools["list_comments"] = {
            "description": "List comment-annotations in the document with author, date, text and anchor preview.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.list_comments(),
        }
        self.tools["add_comment"] = {
            "description": "Add an annotation/comment anchored at `start` (or covering [start, end)).",
            "parameters": {"type": "object", "properties": {
                "start": {"type": "integer"},
                "text": {"type": "string"},
                "author": {"type": "string", "default": "Claude"},
                "initials": {"type": "string", "default": "AI"},
                "end": {"type": "integer"},
            }, "required": ["start", "text"]},
            "handler": lambda start, text, author="Claude", initials="AI", end=None:
                self.uno_bridge.add_comment(start, text, author, initials, end),
        }

        # ---- Images / shapes -------------------------------------------

        self.tools["list_images"] = {
            "description": "List embedded images and shapes in the document.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.list_images(),
        }
        self.tools["insert_text_frame"] = {
            "description": "Insert a TextFrame anchored AT_PARAGRAPH containing either plain text or a PageNumber field (page_number=true). Use to replicate Word docshape page numbers that sit at page bottoms outside FooterText. Pair with list_text_frames to discover what to replicate.",
            "parameters": {
                "type": "object",
                "properties": {
                    "paragraph_index": {"type": "integer", "description": "Anchor paragraph index in body"},
                    "width_mm": {"type": "number", "default": 6.7},
                    "height_mm": {"type": "number", "default": 4.94},
                    "text": {"type": "string", "description": "Plain text content (mutex with page_number)"},
                    "page_number": {"type": "boolean", "default": False, "description": "Insert PageNumber field (auto-renders 1,2,3,...)"},
                    "hori_orient": {"type": "string", "default": "center"},
                    "vert_orient": {"type": "string", "default": "bottom"},
                    "hori_relation": {"type": "string", "default": "page"},
                    "vert_relation": {"type": "string", "default": "page"},
                    "x_mm": {"type": "number", "description": "Manual x offset (only when hori_orient='none')"},
                    "y_mm": {"type": "number", "description": "Manual y offset (only when vert_orient='none')"},
                    "back_transparent": {"type": "boolean", "default": True},
                    "remove_borders": {"type": "boolean", "default": True},
                },
            },
            "handler": lambda **kw: self.uno_bridge.insert_text_frame(**kw),
        }

        self.tools["insert_image"] = {
            "description": "Insert an image from the given file path at `position` (or end of document). Optional width_mm/height_mm to size it.",
            "parameters": {"type": "object", "properties": {
                "path": {"type": "string"},
                "position": {"type": "integer"},
                "width_mm": {"type": "number"},
                "height_mm": {"type": "number"},
            }, "required": ["path"]},
            "handler": lambda path, position=None, width_mm=None, height_mm=None:
                self.uno_bridge.insert_image(path, position, width_mm, height_mm),
        }

        # ---- Sections ---------------------------------------------------

        self.tools["list_sections"] = {
            "description": "List text sections in the document with name, protected flag and preview.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.list_sections(),
        }

        # ---- Tables (read + write cells, insert / delete) --------------

        self.tools["read_table_cells"] = {
            "description": "Return a 2D grid of cell strings for a table (by name or index, default = first table).",
            "parameters": {"type": "object", "properties": {
                "table_name": {"type": "string"},
                "table_index": {"type": "integer"},
            }},
            "handler": lambda table_name=None, table_index=None:
                self.uno_bridge.read_table_cells(table_name, table_index),
        }
        self.tools["write_table_cell"] = {
            "description": "Set the string value of a single table cell (e.g. cell='B3').",
            "parameters": {"type": "object", "properties": {
                "table_name": {"type": "string"},
                "cell": {"type": "string"},
                "value": {"type": "string"},
            }, "required": ["table_name", "cell", "value"]},
            "handler": lambda table_name, cell, value:
                self.uno_bridge.write_table_cell(table_name, cell, value),
        }
        self.tools["read_table_rich"] = {
            "description": (
                "Return each cell of a table as a list of paragraphs with runs "
                "(the same shape as get_paragraphs_with_runs entries). Use this "
                "instead of read_table_cells when you need to replicate font / "
                "bold / alignment / indent / line-spacing inside cells."
            ),
            "parameters": {"type": "object", "properties": {
                "table_name": {"type": "string"},
                "table_index": {"type": "integer"},
            }},
            "handler": lambda table_name=None, table_index=None:
                self.uno_bridge.read_table_rich(table_name, table_index),
        }
        self.tools["write_table_cell_rich"] = {
            "description": (
                "Write a list of paragraphs (with per-run formatting) into a "
                "single cell. Accepts the structure read_table_rich emits "
                "for that cell. Replaces existing cell contents. Use this to "
                "faithfully replicate a table that has formatting inside cells."
            ),
            "parameters": {"type": "object", "properties": {
                "table_name": {"type": "string"},
                "cell": {"type": "string"},
                "paragraphs": {"type": "array", "items": {"type": "object"}},
            }, "required": ["table_name", "cell", "paragraphs"]},
            "handler": lambda table_name, cell, paragraphs:
                self.uno_bridge.write_table_cell_rich(table_name, cell, paragraphs),
        }
        self.tools["insert_table"] = {
            "description": (
                "Insert a new text table with given rows × columns at "
                "`position` (or end of doc). Pass column_widths_mm to "
                "replicate non-uniform column widths from a source table — "
                "without it all columns are equal width, narrow-text cells "
                "get awkwardly wrapped (especially when ParaAdjust=block_line)."
            ),
            "parameters": {"type": "object", "properties": {
                "rows": {"type": "integer"},
                "columns": {"type": "integer"},
                "position": {"type": "integer"},
                "name": {"type": "string"},
                "column_widths_mm": {"type": "array", "items": {"type": "number"}},
                "table_width_mm": {"type": "number"},
                "split": {"type": "boolean"},
                "repeat_headline": {"type": "boolean"},
                "header_row_count": {"type": "integer"},
                "keep_together": {"type": "boolean"},
            }, "required": ["rows", "columns"]},
            "handler": lambda **kw: self.uno_bridge.insert_table(**kw),
        }
        self.tools["remove_table"] = {
            "description": "Delete a table by name.",
            "parameters": {"type": "object", "properties": {"table_name": {"type": "string"}},
                           "required": ["table_name"]},
            "handler": lambda table_name: self.uno_bridge.remove_table(table_name),
        }

        # ---- Undo / Redo / Dispatch -----------------------------------

        self.tools["undo"] = {
            "description": "Undo the last N edits (default 1). Equivalent to Cmd+Z pressed N times.",
            "parameters": {"type": "object", "properties": {
                "steps": {"type": "integer", "default": 1}
            }},
            "handler": lambda steps=1: self.uno_bridge.undo(steps),
        }
        self.tools["redo"] = {
            "description": "Redo the last N undone edits (default 1).",
            "parameters": {"type": "object", "properties": {
                "steps": {"type": "integer", "default": 1}
            }},
            "handler": lambda steps=1: self.uno_bridge.redo(steps),
        }
        self.tools["get_undo_history"] = {
            "description": "Return the list of actions on the undo and redo stacks (most recent first).",
            "parameters": {"type": "object", "properties": {
                "limit": {"type": "integer", "default": 20}
            }},
            "handler": lambda limit=20: self.uno_bridge.get_undo_history(limit),
        }
        self.tools["dispatch_uno_command"] = {
            "description": (
                "Execute a vetted built-in LibreOffice UNO command on the active document. "
                "Server-side WHITELIST: only ~50 commands proven safe on macOS are accepted; "
                "anything else is refused with an explanatory error (NEVER hangs the server). "
                "The whitelist covers the categories below; for everything outside use "
                "dedicated wrapper tools or the LibreOffice menu manually.\n\n"
                "**Allowed categories (all safe — won't hang):**\n"
                " • Character formatting: .uno:Bold, .uno:Italic, .uno:Underline, "
                ".uno:UnderlineDouble, .uno:Strikeout, .uno:Overline, "
                ".uno:Subscript, .uno:Superscript, .uno:Shadowed, .uno:Outline, "
                ".uno:UppercaseSelection, .uno:LowercaseSelection, "
                ".uno:Grow, .uno:Shrink (font ±1pt), "
                ".uno:DefaultCharStyle, .uno:ResetAttributes\n"
                " • Paragraph formatting: .uno:LeftPara, .uno:RightPara, .uno:CenterPara, "
                ".uno:JustifyPara, .uno:DefaultBullet, .uno:DefaultNumbering, "
                ".uno:DecrementIndent, .uno:IncrementIndent, "
                ".uno:DecrementSubLevels, .uno:IncrementSubLevels, "
                ".uno:ParaspaceIncrease, .uno:ParaspaceDecrease\n"
                " • Insertion (no dialog): .uno:InsertPagebreak, .uno:InsertColumnBreak, "
                ".uno:InsertLinebreak, .uno:InsertNonBreakingSpace, "
                ".uno:InsertNarrowNoBreakSpace, .uno:InsertHardHyphen, .uno:InsertSoftHyphen\n"
                " • Navigation: .uno:GoToStartOfDoc, .uno:GoToEndOfDoc, "
                ".uno:GoToStartOfLine, .uno:GoToEndOfLine, "
                ".uno:GoToNextPara, .uno:GoToPrevPara, "
                ".uno:GoToNextPage, .uno:GoToPreviousPage, "
                ".uno:GoToNextWord, .uno:GoToPrevWord, "
                ".uno:GoUp, .uno:GoDown, .uno:GoLeft, .uno:GoRight\n"
                " • Selection: .uno:SelectAll, .uno:SelectWord, .uno:SelectSentence, "
                ".uno:SelectParagraph, .uno:SelectLine\n"
                " • Editing: .uno:Cut, .uno:Copy, .uno:Paste, .uno:Undo, .uno:Redo, "
                ".uno:Delete, .uno:DelToStartOfWord, .uno:DelToEndOfWord, "
                ".uno:DelToStartOfLine, .uno:DelToEndOfLine, "
                ".uno:DelToStartOfPara, .uno:DelToEndOfPara\n\n"
                "**For Save / Export / Print** — there is no MCP equivalent on macOS "
                "(UNO blocks on UI thread). User must press Cmd+S in LibreOffice. "
                "Any attempt via this tool returns an error.\n\n"
                "Prefer the dedicated wrapper tools where they exist "
                "(e.g. set_paragraph_alignment, undo/redo, select_range) — they have "
                "richer arguments and clearer return values."
            ),
            "parameters": {"type": "object", "properties": {
                "command": {"type": "string", "description": "Command name, with or without '.uno:' prefix (e.g. 'Bold' or '.uno:Bold')"},
                "properties": {"type": "object", "description": "Optional command arguments as a dict (PropertyValue tuple is built automatically)"},
            }, "required": ["command"]},
            "handler": lambda command, properties=None: self.uno_bridge.dispatch_uno_command(command, properties),
        }

        self.tools["enable_header"] = {
            "description": "Enable or disable header on a page style.",
            "parameters": {"type": "object", "properties": {
                "enabled": {"type": "boolean", "default": True},
                "page_style": {"type": "string", "default": "Default Page Style"},
            }},
            "handler": lambda enabled=True, page_style="Default Page Style":
                self.uno_bridge.enable_header(enabled, page_style),
        }
        self.tools["enable_footer"] = {
            "description": "Enable or disable footer on a page style.",
            "parameters": {"type": "object", "properties": {
                "enabled": {"type": "boolean", "default": True},
                "page_style": {"type": "string", "default": "Default Page Style"},
            }},
            "handler": lambda enabled=True, page_style="Default Page Style":
                self.uno_bridge.enable_footer(enabled, page_style),
        }
        self.tools["set_header"] = {
            "description": "Set header text on a page style. Enables the header automatically.",
            "parameters": {"type": "object", "properties": {
                "text": {"type": "string"},
                "page_style": {"type": "string", "default": "Default Page Style"},
            }, "required": ["text"]},
            "handler": lambda text, page_style="Default Page Style":
                self.uno_bridge.set_header(text, page_style),
        }
        self.tools["set_footer"] = {
            "description": "Set footer text on a page style. Enables the footer automatically.",
            "parameters": {"type": "object", "properties": {
                "text": {"type": "string"},
                "page_style": {"type": "string", "default": "Default Page Style"},
            }, "required": ["text"]},
            "handler": lambda text, page_style="Default Page Style":
                self.uno_bridge.set_footer(text, page_style),
        }
        self.tools["get_header"] = {
            "description": "Read header text and enabled state for a page style.",
            "parameters": {"type": "object", "properties": {
                "page_style": {"type": "string", "default": "Default Page Style"}}},
            "handler": lambda page_style="Default Page Style": self.uno_bridge.get_header(page_style),
        }
        self.tools["get_footer"] = {
            "description": "Read footer text and enabled state for a page style.",
            "parameters": {"type": "object", "properties": {
                "page_style": {"type": "string", "default": "Default Page Style"}}},
            "handler": lambda page_style="Default Page Style": self.uno_bridge.get_footer(page_style),
        }

        self.tools["get_tables_info"] = {
            "description": "List text tables in the active Writer document with their dimensions.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.get_tables_info(),
        }

        self.tools["clone_document"] = {
            "description": "Convert a file on disk from one format to another via a transient hidden LibreOffice component. Bypasses the macOS UI-thread save deadlock — does NOT touch any currently-open document. target_format is auto-derived from target_path extension if omitted. Supports docx, doc, odt, rtf, txt, html, xhtml, pdf, epub, xlsx, xls, ods, csv, pptx, ppt, odp.",
            "parameters": {
                "type": "object",
                "properties": {
                    "source_path": {"type": "string", "description": "Absolute path of the source file"},
                    "target_path": {"type": "string", "description": "Absolute path of the target file (overwrites)"},
                    "target_format": {"type": "string", "description": "Optional explicit format key (docx/odt/pdf/...). Defaults to target_path extension."},
                },
                "required": ["source_path", "target_path"],
            },
            "handler": lambda **kw: self.uno_bridge.clone_document(**kw),
        }

        self.tools["read_paragraph_xml"] = {
            "description": (
                "Read raw ODT XML for a paragraph at index `paragraph_index` "
                "in `source_path`, plus all referenced styles (paragraph "
                "style chain + per-span T-styles from automatic-styles). "
                "Use when UNO API does NOT surface a property you need to "
                "replicate — Word→ODT exporters often emit fo:* attributes "
                "(letter-spacing, break-before, keep-with-next, hyphenate, "
                "padding) on automatic styles which never round-trip "
                "through pyuno. paragraph_index matches get_paragraphs "
                "indexing (body paragraphs only, table cells skipped)."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "source_path": {"type": "string"},
                    "paragraph_index": {"type": "integer"},
                    "include_styles": {"type": "boolean", "default": True},
                },
                "required": ["source_path", "paragraph_index"],
            },
            "handler": lambda **kw: self.uno_bridge.read_paragraph_xml(**kw),
        }

        # export_active_document REMOVED — storeToURL on a visible component blocks
        # the HTTP worker thread on macOS (AppKit UI-thread deadlock) and wedges the
        # entire server until LibreOffice is restarted. Use clone_document for
        # file-on-disk conversion (hidden component, no UI-thread contact).

        self.tools["execute_batch"] = {
            "description": (
                "Run a list of tool invocations sequentially in a single HTTP "
                "round-trip. Each operation is {tool: <name>, args: {...}}. "
                "Returns a parallel list of results in the same order. "
                "By default wraps the whole batch in lockControllers / "
                "unlockControllers so view updates are deferred to the end. "
                "ALSO auto-detects 'heavy' operations (clone_page_style, "
                "clone_paragraph_style, clone_numbering_rule, set_page_*_props, "
                "set_paragraph_style_props) and TEMPORARILY hides the active "
                "window for the duration of the batch, restoring visibility "
                "afterwards — required on macOS, where AppKit paint cycles on a "
                "visible doc can hold SolarMutex and deadlock the HTTP-worker. "
                "Set lock_view=false / auto_hide=false to opt out (e.g. demos)."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "operations": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "tool": {"type": "string"},
                                "args": {"type": "object"},
                            },
                            "required": ["tool"],
                        },
                    },
                    "stop_on_error": {"type": "boolean", "default": False},
                    "lock_view": {"type": "boolean", "default": True,
                                  "description": "Freeze view updates during the batch (recommended)."},
                    "auto_hide": {"type": "string", "default": "auto",
                                  "description": "'auto' (hide if any heavy op present), 'always' (hide unconditionally), 'never' (do not hide). Window visibility restored after batch."},
                },
                "required": ["operations"],
            },
            "handler": lambda operations, stop_on_error=False, lock_view=True, auto_hide="auto":
                self._execute_batch(operations, stop_on_error, lock_view, auto_hide),
        }

        self.tools["lock_view"] = {
            "description": "Freeze view updates of the active doc (lockControllers). Pair with unlock_view. Manual escape hatch — execute_batch does this automatically.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.lock_view(),
        }

        self.tools["unlock_view"] = {
            "description": "Resume view updates of the active doc (unlockControllers).",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.unlock_view(),
        }

        logger.info(f"Registered {len(self.tools)} MCP tools")

    # Operations known to mutate page-style / paragraph-style structures and
    # hit the SolarMutex-paint deadlock on macOS when run on a visible doc.
    _HEAVY_OPS = frozenset({
        "clone_page_style", "clone_paragraph_style", "clone_numbering_rule",
        "set_page_style_props", "set_paragraph_style_props",
        "set_page_margins", "apply_paragraph_style",
        "write_table_cell_rich",
    })

    def _execute_batch(self, operations, stop_on_error: bool = False,
                       lock_view: bool = True, auto_hide="auto"):
        results = []
        ok = 0
        locked = False
        # Decide whether to hide the window.
        mode = str(auto_hide).lower() if auto_hide is not None else "auto"
        if mode == "always":
            should_hide = True
        elif mode == "never" or mode in ("false", "0"):
            should_hide = False
        else:  # 'auto' / 'true' — detect heavy ops
            should_hide = any(
                isinstance(op, dict) and op.get("tool") in self._HEAVY_OPS
                for op in operations
            )
        was_visible = None
        if should_hide:
            try:
                hr = self.uno_bridge.hide_window()
                if hr.get("success"):
                    was_visible = bool(hr.get("was_visible", True))
            except Exception:
                was_visible = None
        if lock_view:
            try:
                lr = self.uno_bridge.lock_view()
                locked = bool(lr.get("success"))
            except Exception:
                locked = False
        try:
            for i, op in enumerate(operations):
                tool_name = op.get("tool") if isinstance(op, dict) else None
                if not tool_name:
                    results.append({"success": False, "error": "missing 'tool' field", "index": i})
                    if stop_on_error: break
                    continue
                args = op.get("args") or {}
                tool = self.tools.get(tool_name)
                if tool is None:
                    results.append({"success": False, "error": f"unknown tool: {tool_name}", "index": i})
                    if stop_on_error: break
                    continue
                if tool_name == "execute_batch":
                    results.append({"success": False, "error": "execute_batch cannot be nested", "index": i})
                    if stop_on_error: break
                    continue
                try:
                    r = tool["handler"](**args)
                except Exception as e:
                    r = {"success": False, "error": f"{type(e).__name__}: {e}"}
                results.append(r)
                if isinstance(r, dict) and r.get("success", True):
                    ok += 1
                elif stop_on_error:
                    break
        finally:
            if locked:
                try: self.uno_bridge.unlock_view()
                except Exception: pass
            if was_visible:
                try: self.uno_bridge.show_window()
                except Exception: pass
        return {"success": True, "results": results, "ran": len(results), "ok": ok,
                "view_locked": locked, "window_hidden": was_visible is not None,
                "restored_visibility": bool(was_visible)}
    
    async def execute_tool(self, tool_name: str, parameters: Dict[str, Any]) -> Dict[str, Any]:
        """
        Execute an MCP tool
        
        Args:
            tool_name: Name of the tool to execute
            parameters: Parameters for the tool
            
        Returns:
            Result dictionary
        """
        try:
            if tool_name not in self.tools:
                return {
                    "success": False,
                    "error": f"Unknown tool: {tool_name}",
                    "available_tools": list(self.tools.keys())
                }
            
            tool = self.tools[tool_name]
            handler = tool["handler"]
            
            # Execute the tool handler
            result = handler(**parameters)
            
            logger.info(f"Executed tool '{tool_name}' successfully")
            return result
            
        except Exception as e:
            logger.error(f"Error executing tool '{tool_name}': {e}")
            return {
                "success": False,
                "error": str(e),
                "tool": tool_name,
                "parameters": parameters
            }
    
    def get_tool_list(self) -> List[Dict[str, Any]]:
        """Get list of available tools with their descriptions"""
        return [
            {
                "name": name,
                "description": tool["description"],
                "parameters": tool["parameters"]
            }
            for name, tool in self.tools.items()
        ]
    
    # Tool handler methods
    
    def create_document_live(self, doc_type: str = "writer", visible: bool = True) -> Dict[str, Any]:
        """Create a new document. visible=False → hidden window (use for batch writes)."""
        try:
            doc = self.uno_bridge.create_document(doc_type, visible=visible)
            url = ""
            try:
                if hasattr(doc, "getURL"):
                    url = doc.getURL() or ""
            except Exception:
                url = ""
            return {
                "success": True,
                "message": f"Created new {doc_type} document",
                "document_type": doc_type,
                "url": url,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def insert_text_live(self, text: str, position=None) -> Dict[str, Any]:
        """Insert text into the currently active document. position: 'end'|'cursor'|int."""
        return self.uno_bridge.insert_text(text, position)
    
    def get_document_info_live(self) -> Dict[str, Any]:
        """Get information about the currently active document"""
        doc_info = self.uno_bridge.get_document_info()
        if "error" in doc_info:
            return {"success": False, **doc_info}
        else:
            return {"success": True, "document_info": doc_info}
    
    def format_text_live(self, **formatting) -> Dict[str, Any]:
        """Apply formatting. start/end (if given) → explicit char range, else current selection."""
        start = formatting.pop("start", None)
        end = formatting.pop("end", None)
        return self.uno_bridge.format_text(formatting, start=start, end=end)
    
    def get_text_content_live(self) -> Dict[str, Any]:
        """Get text content of the currently active document"""
        return self.uno_bridge.get_text_content()
    
    def list_open_documents(self) -> Dict[str, Any]:
        """List all open documents in LibreOffice"""
        try:
            desktop = self.uno_bridge.desktop
            documents = []
            
            # Get all open documents
            frames = desktop.getFrames()
            for i in range(frames.getCount()):
                frame = frames.getByIndex(i)
                controller = frame.getController()
                if controller:
                    doc = controller.getModel()
                    if doc:
                        doc_info = self.uno_bridge.get_document_info(doc)
                        documents.append(doc_info)
            
            return {
                "success": True,
                "documents": documents,
                "count": len(documents)
            }
            
        except Exception as e:
            return {"success": False, "error": str(e)}


# Global instance
mcp_server = None

def get_mcp_server() -> LibreOfficeMCPServer:
    """Get or create the global MCP server instance"""
    global mcp_server
    if mcp_server is None:
        mcp_server = LibreOfficeMCPServer()
    return mcp_server
