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
            "description": "Create a new document in LibreOffice",
            "parameters": {
                "type": "object",
                "properties": {
                    "doc_type": {
                        "type": "string",
                        "enum": ["writer", "calc", "impress", "draw"],
                        "description": "Type of document to create",
                        "default": "writer"
                    }
                }
            },
            "handler": self.create_document_live
        }
        
        # Text manipulation tools
        self.tools["insert_text_live"] = {
            "description": "Insert text into the currently active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "text": {
                        "type": "string",
                        "description": "Text to insert"
                    },
                    "position": {
                        "type": "integer",
                        "description": "Position to insert at (optional, defaults to cursor position)"
                    }
                },
                "required": ["text"]
            },
            "handler": self.insert_text_live
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
            "description": "Apply formatting to selected text in active document",
            "parameters": {
                "type": "object",
                "properties": {
                    "bold": {
                        "type": "boolean",
                        "description": "Apply bold formatting"
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Apply italic formatting"
                    },
                    "underline": {
                        "type": "boolean",
                        "description": "Apply underline formatting"
                    },
                    "font_size": {
                        "type": "number",
                        "description": "Font size in points"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "Font family name"
                    }
                }
            },
            "handler": self.format_text_live
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
            "description": "Set font color of the current selection (or view-cursor paragraph). Color is hex like '#FF0000' or 'FF0000' or an int.",
            "parameters": {
                "type": "object",
                "properties": {"color": {"type": "string", "description": "Hex color, e.g. '#FF0000'"}},
                "required": ["color"],
            },
            "handler": lambda color: self.uno_bridge.set_text_color(color),
        }

        self.tools["set_background_color"] = {
            "description": "Set character background (highlight) color of the current selection.",
            "parameters": {
                "type": "object",
                "properties": {"color": {"type": "string", "description": "Hex color, e.g. '#FFFF00'. Use -1 for 'no fill'."}},
                "required": ["color"],
            },
            "handler": lambda color: self.uno_bridge.set_background_color(color),
        }

        self.tools["set_paragraph_alignment"] = {
            "description": "Set paragraph alignment for current selection / cursor paragraph.",
            "parameters": {
                "type": "object",
                "properties": {
                    "alignment": {"type": "string", "enum": ["left", "center", "right", "justify"]}
                },
                "required": ["alignment"],
            },
            "handler": lambda alignment: self.uno_bridge.set_paragraph_alignment(alignment),
        }

        self.tools["set_paragraph_indent"] = {
            "description": "Set paragraph indents in millimeters (omit a field to leave it unchanged).",
            "parameters": {
                "type": "object",
                "properties": {
                    "left_mm": {"type": "number", "description": "Left indent in mm"},
                    "right_mm": {"type": "number", "description": "Right indent in mm"},
                    "first_line_mm": {"type": "number", "description": "First-line indent in mm"},
                },
            },
            "handler": lambda **kw: self.uno_bridge.set_paragraph_indent(**kw),
        }

        self.tools["set_line_spacing"] = {
            "description": "Set line spacing for current paragraph(s). proportional: value=100 single, 150=1.5x, 200=double. fix/minimum/leading: value in mm.",
            "parameters": {
                "type": "object",
                "properties": {
                    "mode": {"type": "string", "enum": ["proportional", "minimum", "leading", "fix"], "default": "proportional"},
                    "value": {"type": "number", "description": "% for proportional, mm otherwise"},
                },
                "required": ["value"],
            },
            "handler": lambda mode="proportional", value=100: self.uno_bridge.set_line_spacing(mode, value),
        }

        self.tools["apply_paragraph_style"] = {
            "description": "Apply a paragraph style by name (e.g. 'Heading 1', 'Heading 2', 'Quotations', 'Default Paragraph Style').",
            "parameters": {
                "type": "object",
                "properties": {"style_name": {"type": "string"}},
                "required": ["style_name"],
            },
            "handler": lambda style_name: self.uno_bridge.apply_paragraph_style(style_name),
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
            "description": "Like get_paragraphs but also returns inline character runs per paragraph (text portions with uniform formatting): font, size, bold/italic/underline/strike, color, hyperlink URL, char style. Use this for faithful Markdown/HTML export when inline formatting matters.",
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

        self.tools["get_character_format"] = {
            "description": "Read character formatting (font, size, bold/italic/underline, color, bg color) over a char range. Pass only `start` to read a single character.",
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
            "description": "Page count, current page, page size (mm) of the active Writer document.",
            "parameters": {"type": "object", "properties": {}},
            "handler": lambda: self.uno_bridge.get_page_info(),
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
        self.tools["insert_table"] = {
            "description": "Insert a new text table with given rows × columns at `position` (or end of doc).",
            "parameters": {"type": "object", "properties": {
                "rows": {"type": "integer"},
                "columns": {"type": "integer"},
                "position": {"type": "integer"},
                "name": {"type": "string"},
            }, "required": ["rows", "columns"]},
            "handler": lambda rows, columns, position=None, name=None:
                self.uno_bridge.insert_table(rows, columns, position, name),
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

        logger.info(f"Registered {len(self.tools)} MCP tools")
    
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
    
    def create_document_live(self, doc_type: str = "writer") -> Dict[str, Any]:
        """Create a new document in LibreOffice"""
        try:
            doc = self.uno_bridge.create_document(doc_type)
            return {
                "success": True,
                "message": f"Created new {doc_type} document",
                "document_type": doc_type,
                "url": doc.getURL() if hasattr(doc, "getURL") else "",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def insert_text_live(self, text: str, position: Optional[int] = None) -> Dict[str, Any]:
        """Insert text into the currently active document"""
        return self.uno_bridge.insert_text(text, position)
    
    def get_document_info_live(self) -> Dict[str, Any]:
        """Get information about the currently active document"""
        doc_info = self.uno_bridge.get_document_info()
        if "error" in doc_info:
            return {"success": False, **doc_info}
        else:
            return {"success": True, "document_info": doc_info}
    
    def format_text_live(self, **formatting) -> Dict[str, Any]:
        """Apply formatting to selected text"""
        return self.uno_bridge.format_text(formatting)
    
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
