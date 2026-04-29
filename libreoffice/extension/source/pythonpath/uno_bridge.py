"""
LibreOffice MCP Extension - UNO Bridge Module

This module provides a bridge between MCP operations and LibreOffice UNO API,
enabling direct manipulation of LibreOffice documents.
"""

import uno
import unohelper
from com.sun.star.beans import PropertyValue
from typing import Any, Optional, Dict, List
import logging
import traceback

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class UNOBridge:
    """Bridge between MCP operations and LibreOffice UNO API"""
    
    def __init__(self):
        """Initialize the UNO bridge"""
        try:
            self.ctx = uno.getComponentContext()
            self.smgr = self.ctx.ServiceManager
            self.desktop = self.smgr.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)
            self._last_active_doc = None
            logger.info("UNO Bridge initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize UNO Bridge: {e}")
            raise
    
    def create_document(self, doc_type: str = "writer", visible: bool = True) -> Any:
        """
        Create new document using UNO API
        
        Args:
            doc_type: Type of document ('writer', 'calc', 'impress', 'draw')
            
        Returns:
            Document object
        """
        try:
            url_map = {
                "writer": "private:factory/swriter",
                "calc": "private:factory/scalc", 
                "impress": "private:factory/simpress",
                "draw": "private:factory/sdraw"
            }
            
            url = url_map.get(doc_type, "private:factory/swriter")
            # Create the document hidden to avoid UI-thread deadlock when
            # called from a background HTTP-server thread, then show its
            # window as a separate, non-blocking operation.
            hidden = PropertyValue()
            hidden.Name = "Hidden"
            hidden.Value = True
            doc = self.desktop.loadComponentFromURL(url, "_blank", 0, (hidden,))
            if visible:
                try:
                    ctrl = doc.getCurrentController()
                    if ctrl is not None:
                        frame = ctrl.getFrame()
                        if frame is not None:
                            win = frame.getContainerWindow()
                            if win is not None:
                                win.setVisible(True)
                            # NOTE: do NOT call frame.activate() / setActiveFrame()
                            # here — both block on AppKit UI thread on macOS when
                            # invoked from a background HTTP-server thread.
                except Exception as e:
                    logger.warning(f"Created document but could not show window: {e}")
            # Remember the most recently created/opened doc as a fallback
            # anchor for get_active_document — survives cases where
            # setActiveFrame doesn't take effect immediately.
            self._last_active_doc = doc
            logger.info(f"Created new {doc_type} document")
            return doc
            
        except Exception as e:
            logger.error(f"Failed to create document: {e}")
            raise
    
    def get_active_document(self) -> Optional[Any]:
        """Get currently active document.

        Falls back to scanning open Components if `getCurrentComponent`
        returns nothing useful (happens after creating a doc with Hidden=True
        and re-showing it — its frame is not the active one yet).
        """
        # Prefer the most recently created/opened document if it's still alive
        # — this beats both getCurrentComponent (returns Start Center after
        # Hidden=True load) and frame-scan (returns the wrong writer).
        try:
            cached = self._last_active_doc
            if cached is not None and hasattr(cached, "supportsService"):
                # Liveness probe — disposed components throw on any call.
                _ = cached.getURL() if hasattr(cached, "getURL") else None
                return cached
        except Exception:
            self._last_active_doc = None  # disposed — drop it

        # First try the truly-active document via getCurrentComponent — but
        # only accept it if it's a real document (not the Start Center).
        try:
            doc = self.desktop.getCurrentComponent()
            if doc is not None and hasattr(doc, "supportsService") and (
                doc.supportsService("com.sun.star.text.TextDocument")
                or doc.supportsService("com.sun.star.sheet.SpreadsheetDocument")
                or doc.supportsService("com.sun.star.presentation.PresentationDocument")
                or doc.supportsService("com.sun.star.drawing.DrawingDocument")
            ):
                return doc
        except Exception as e:
            logger.warning(f"getCurrentComponent failed: {e}")

        # Fall back to scanning frames (same path used by list_open_documents).
        try:
            writers, others = [], []
            frames = self.desktop.getFrames()
            for i in range(frames.getCount()):
                frame = frames.getByIndex(i)
                controller = frame.getController() if frame else None
                doc = controller.getModel() if controller else None
                if doc is None or not hasattr(doc, "supportsService"):
                    continue
                if doc.supportsService("com.sun.star.text.TextDocument"):
                    writers.append(doc)
                elif (doc.supportsService("com.sun.star.sheet.SpreadsheetDocument")
                      or doc.supportsService("com.sun.star.presentation.PresentationDocument")
                      or doc.supportsService("com.sun.star.drawing.DrawingDocument")):
                    others.append(doc)
            if writers:
                return writers[0]
            if others:
                return others[0]
        except Exception as e:
            logger.error(f"Frame enumeration failed: {e}")
        return None
    
    def get_document_info(self, doc: Any = None) -> Dict[str, Any]:
        """Get information about a document"""
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"error": "No document available"}
            
            info = {
                "title": getattr(doc, 'Title', 'Unknown') if hasattr(doc, 'Title') else "Unknown",
                "url": doc.getURL() if hasattr(doc, 'getURL') else "",
                "modified": doc.isModified() if hasattr(doc, 'isModified') else False,
                "type": self._get_document_type(doc),
                "has_selection": self._has_selection(doc)
            }
            
            # Add document-specific information
            try:
                if doc.supportsService("com.sun.star.text.TextDocument"):
                    text = doc.getText()
                    info["word_count"] = len(text.getString().split())
                    info["character_count"] = len(text.getString())
                elif doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                    sheets = doc.getSheets()
                    info["sheet_count"] = sheets.getCount()
                    info["sheet_names"] = [sheets.getByIndex(i).getName()
                                         for i in range(sheets.getCount())]
            except Exception as e:
                logger.warning(f"Could not enrich document info: {e}")
            
            return info
            
        except Exception as e:
            logger.error(f"Failed to get document info: {e}")
            return {"error": str(e)}
    
    def insert_text(self, text: str, position=None, doc: Any = None) -> Dict[str, Any]:
        """Insert text into the active Writer document.

        position:
          - "end" (default) → append at the end of the document body. Safe for
            batch generation; not affected by prior select_range calls.
          - "cursor"        → insert at the current view-cursor position.
            Note: select_range() moves the view cursor onto the selection,
            so "cursor" after select_range will insert/replace there.
          - int             → absolute char offset from the document start.

        '\\n' in `text` is converted to a real paragraph break.
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            if not doc:
                return {"success": False, "error": "No active document"}
            if not doc.supportsService("com.sun.star.text.TextDocument"):
                return {"success": False, "error": f"Text insertion not supported for {self._get_document_type(doc)}"}

            text_obj = doc.getText()
            if position is None or position == "end":
                cursor = text_obj.createTextCursor()
                cursor.gotoEnd(False)
                where = "end"
            elif position == "cursor":
                cursor = doc.getCurrentController().getViewCursor()
                where = "cursor"
            else:
                try:
                    pos_int = int(position)
                except (TypeError, ValueError):
                    return {"success": False, "error": f"position must be int|'end'|'cursor', got {position!r}"}
                cursor = text_obj.createTextCursor()
                cursor.gotoStart(False)
                cursor.goRight(pos_int, False)
                where = pos_int

            parts = text.split("\n")
            for i, part in enumerate(parts):
                if i > 0:
                    text_obj.insertControlCharacter(cursor, 0, False)
                if part:
                    text_obj.insertString(cursor, part, False)
            logger.info(f"Inserted {len(text)} characters into Writer document at {where}")
            return {"success": True, "message": f"Inserted {len(text)} characters at {where}"}

        except Exception as e:
            logger.error(f"Failed to insert text: {e}")
            return {"success": False, "error": str(e)}
    
    def format_text(self, formatting: Dict[str, Any], doc: Any = None,
                    start=None, end=None) -> Dict[str, Any]:
        """Apply character formatting to a range.

        If `start` and `end` are given, format that explicit char-range
        (end-exclusive). Otherwise fall back to the current selection.
        Pass start/end for batch ops — it doesn't depend on selection state.
        """
        try:
            if doc is None:
                doc = self.get_active_document()

            if not doc or not doc.supportsService("com.sun.star.text.TextDocument"):
                return {"success": False, "error": "No Writer document available"}

            if start is not None and end is not None:
                text_range = self._resolve_range(doc, start, end)
            else:
                selection = doc.getCurrentController().getSelection()
                if selection.getCount() == 0:
                    return {"success": False, "error": "No text selected (and no start/end provided)"}
                text_range = selection.getByIndex(0)

            # Apply various formatting options
            if "bold" in formatting:
                text_range.CharWeight = 150.0 if formatting["bold"] else 100.0
            
            if "italic" in formatting:
                text_range.CharPosture = 2 if formatting["italic"] else 0
            
            if "underline" in formatting:
                text_range.CharUnderline = 1 if formatting["underline"] else 0
            
            if "font_size" in formatting:
                text_range.CharHeight = formatting["font_size"]
            
            if "font_name" in formatting:
                text_range.CharFontName = formatting["font_name"]
            
            logger.info("Applied formatting to selected text")
            return {"success": True, "message": "Formatting applied successfully"}
            
        except Exception as e:
            logger.error(f"Failed to format text: {e}")
            return {"success": False, "error": str(e)}
    
    # save_document() removed: blocks the HTTP server thread waiting on
    # the macOS UI thread. User must press Cmd+S in LibreOffice instead.
    def _removed_save_document(self, doc: Any = None, file_path: Optional[str] = None) -> Dict[str, Any]:
        """
        Save a document
        
        Args:
            doc: Document to save (None for active document)
            file_path: Path to save to (None to save to current location)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document to save"}
            
            # Save UNO calls block on the main UI thread on macOS Sequoia
            # when invoked from our background HTTP-server thread. Run in a
            # separate daemon thread and join with a timeout so the HTTP
            # response returns promptly. Actual disk write completes in the
            # background; clients can verify by checking the file's mtime.
            import threading

            overwrite = PropertyValue(); overwrite.Name = "Overwrite"; overwrite.Value = True

            if file_path:
                target_url = self._path_to_url(file_path)
            elif doc.hasLocation():
                target_url = doc.getURL()
            else:
                return {"success": False, "error": "Document has no location, specify file_path"}

            result = {"state": "pending"}

            def _store():
                try:
                    doc.storeToURL(target_url, (overwrite,))
                    result["state"] = "success"
                except Exception as e:
                    result["state"] = "error"
                    result["err"] = str(e)

            t = threading.Thread(target=_store, daemon=True, name="mcp-save")
            t.start()
            t.join(timeout=4.0)
            if result["state"] == "success":
                logger.info(f"Saved document synchronously to {target_url}")
                return {"success": True, "message": "Document saved", "url": target_url, "async": False}
            if result["state"] == "error":
                return {"success": False, "error": result.get("err", "unknown")}
            logger.info(f"Save still running in background for {target_url}")
            return {"success": True, "message": "Save initiated; completing in background",
                    "url": target_url, "async": True}
                    
        except Exception as e:
            logger.error(f"Failed to save document: {e}")
            return {"success": False, "error": str(e)}
    
    # export_document() removed: same UI-thread block as save_document.
    def _removed_export_document(self, export_format: str, file_path: str, doc: Any = None) -> Dict[str, Any]:
        """
        Export document to different format
        
        Args:
            export_format: Target format ('pdf', 'docx', 'odt', 'txt', etc.)
            file_path: Path to export to
            doc: Document to export (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document to export"}
            
            # Filter map for different formats
            filter_map = {
                'pdf': 'writer_pdf_Export',
                'docx': 'MS Word 2007 XML',
                'doc': 'MS Word 97',
                'odt': 'writer8',
                'txt': 'Text',
                'rtf': 'Rich Text Format',
                'html': 'HTML (StarWriter)'
            }
            
            filter_name = filter_map.get(export_format.lower())
            if not filter_name:
                return {"success": False, "error": f"Unsupported export format: {export_format}"}
            
            # Prepare export properties
            properties = (
                PropertyValue("FilterName", 0, filter_name, 0),
                PropertyValue("Overwrite", 0, True, 0),
            )
            
            # Export document
            url = uno.systemPathToFileUrl(file_path)
            doc.storeToURL(url, properties)
            
            logger.info(f"Exported document to {file_path} as {export_format}")
            return {"success": True, "message": f"Document exported to {file_path}"}
            
        except Exception as e:
            logger.error(f"Failed to export document: {e}")
            return {"success": False, "error": str(e)}
    
    def get_text_content(self, doc: Any = None) -> Dict[str, Any]:
        """Get text content from a document"""
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No document available"}
            
            if doc.supportsService("com.sun.star.text.TextDocument"):
                text = doc.getText().getString()
                return {"success": True, "content": text, "length": len(text)}
            return {"success": False, "error": f"Text extraction not supported for {self._get_document_type(doc)}"}
                
        except Exception as e:
            logger.error(f"Failed to get text content: {e}")
            return {"success": False, "error": str(e)}
    
    # ---- Extended editing helpers ----------------------------------------

    @staticmethod
    def _hex_to_int(color):
        """Accept '#RRGGBB' / 'RRGGBB' / int → int (0xRRGGBB)."""
        if isinstance(color, int):
            return color
        return int(str(color).lstrip("#"), 16)

    @staticmethod
    def _encode_tab_stops(stops):
        align_rev = {0: "left", 1: "center", 2: "right", 3: "decimal"}
        out = []
        if not stops:
            return out
        for t in stops:
            try:
                out.append({
                    "position_mm": t.Position / 100.0,
                    "alignment": align_rev.get(t.Alignment, "left"),
                    "fill_char": chr(t.FillChar) if t.FillChar else " ",
                    "decimal_char": chr(t.DecimalChar) if t.DecimalChar else ".",
                })
            except Exception:
                continue
        return out

    def _selected_range_or_view_cursor(self, doc):
        """Prefer current selection; fall back to view cursor (paragraph context)."""
        try:
            sel = doc.getCurrentController().getSelection()
            if sel.getCount() > 0:
                rng = sel.getByIndex(0)
                # An empty selection is still a range — fine for paragraph properties.
                return rng
        except Exception:
            pass
        return doc.getCurrentController().getViewCursor()

    def _resolve_range(self, doc, start=None, end=None):
        """Resolve an explicit char-range into a text cursor, or fall back
        to the current selection / view cursor.

        - start, end (int) → cursor over [start, end). End-exclusive.
        - start only       → cursor at single char position (paragraph context).
        - "end" sentinel for `end` means up to document end.
        - both None        → selection (if any) else view cursor.
        """
        if start is None and end is None:
            return self._selected_range_or_view_cursor(doc)
        text = doc.getText()
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        if start is not None:
            cursor.goRight(int(start), False)
        if end == "end":
            cursor.gotoEnd(True)
        elif end is not None:
            length = int(end) - int(start or 0)
            if length < 0:
                length = 0
            cursor.goRight(length, True)
        return cursor

    def _require_writer(self):
        doc = self.get_active_document()
        if not doc or not doc.supportsService("com.sun.star.text.TextDocument"):
            return None, {"success": False, "error": "No Writer document active"}
        return doc, None

    def set_text_color(self, color, start=None, end=None) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            rng = self._resolve_range(doc, start, end)
            rng.CharColor = self._hex_to_int(color)
            return {"success": True, "color": color}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_background_color(self, color, start=None, end=None) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            rng = self._resolve_range(doc, start, end)
            rng.CharBackColor = self._hex_to_int(color)
            return {"success": True, "color": color}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_alignment(self, alignment: str, start=None, end=None) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        # com.sun.star.style.ParagraphAdjust: LEFT=0, RIGHT=1, BLOCK=2, CENTER=3
        mapping = {"left": 0, "right": 1, "justify": 2, "block": 2, "center": 3}
        val = mapping.get(alignment.lower())
        if val is None:
            return {"success": False, "error": "Unknown alignment, use: left|center|right|justify"}
        try:
            rng = self._resolve_range(doc, start, end)
            rng.ParaAdjust = val
            return {"success": True, "alignment": alignment}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_indent(self, left_mm=None, right_mm=None, first_line_mm=None,
                             start=None, end=None) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            rng = self._resolve_range(doc, start, end)
            applied = {}
            if left_mm is not None:
                rng.ParaLeftMargin = int(float(left_mm) * 100)  # 1/100 mm
                applied["left_mm"] = left_mm
            if right_mm is not None:
                rng.ParaRightMargin = int(float(right_mm) * 100)
                applied["right_mm"] = right_mm
            if first_line_mm is not None:
                rng.ParaFirstLineIndent = int(float(first_line_mm) * 100)
                applied["first_line_mm"] = first_line_mm
            return {"success": True, "applied": applied}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_spacing(self, top_mm=None, bottom_mm=None,
                              context_margin=None,
                              start=None, end=None) -> Dict[str, Any]:
        """Set ParaTopMargin / ParaBottomMargin / ParaContextMargin.
        top_mm/bottom_mm in mm. context_margin (bool): when True, adjacent paragraphs
        of the same style collapse top/bottom — affects whether spacings stack."""
        doc, err = self._require_writer()
        if err:
            return err
        try:
            rng = self._resolve_range(doc, start, end)
            applied = {}
            if top_mm is not None:
                rng.ParaTopMargin = int(float(top_mm) * 100)
                applied["top_mm"] = top_mm
            if bottom_mm is not None:
                rng.ParaBottomMargin = int(float(bottom_mm) * 100)
                applied["bottom_mm"] = bottom_mm
            if context_margin is not None:
                rng.ParaContextMargin = bool(context_margin)
                applied["context_margin"] = bool(context_margin)
            return {"success": True, "applied": applied}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_tabs(self, stops, start=None, end=None) -> Dict[str, Any]:
        """Set ParaTabStops on a paragraph range.

        stops: list of {position_mm: float, alignment: 'left'|'right'|'center'|'decimal',
                        fill_char: str = ' ', decimal_char: str = '.'}
        Replaces all existing tab stops for the range.
        """
        doc, err = self._require_writer()
        if err:
            return err
        if not isinstance(stops, list):
            return {"success": False, "error": "stops must be a list"}
        align_map = {"left": 0, "center": 1, "right": 2, "decimal": 3}
        try:
            rng = self._resolve_range(doc, start, end)
            tab_structs = []
            for s in stops:
                t = uno.createUnoStruct("com.sun.star.style.TabStop")
                t.Position = int(float(s.get("position_mm", 0)) * 100)
                t.Alignment = align_map.get(str(s.get("alignment","left")).lower(), 0)
                fill = s.get("fill_char", " ") or " "
                t.FillChar = ord(fill[0]) if isinstance(fill, str) and fill else 32
                dec = s.get("decimal_char", ".") or "."
                t.DecimalChar = ord(dec[0]) if isinstance(dec, str) and dec else 46
                tab_structs.append(t)
            rng.ParaTabStops = tuple(tab_structs)
            return {"success": True, "stops_count": len(tab_structs)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_line_spacing(self, mode: str = "proportional", value: float = 100,
                         start=None, end=None) -> Dict[str, Any]:
        """mode: proportional|minimum|leading|fix; value: % for proportional, mm otherwise."""
        doc, err = self._require_writer()
        if err:
            return err
        try:
            mode_map = {"proportional": 0, "minimum": 1, "leading": 2, "fix": 3}
            mode_val = mode_map.get(mode.lower(), 0)
            ls = uno.createUnoStruct("com.sun.star.style.LineSpacing")
            ls.Mode = mode_val
            ls.Height = int(value) if mode_val == 0 else int(float(value) * 100)
            rng = self._resolve_range(doc, start, end)
            rng.ParaLineSpacing = ls
            return {"success": True, "mode": mode, "value": value}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def apply_paragraph_style(self, style_name: str, start=None, end=None,
                              target: str = None) -> Dict[str, Any]:
        """Apply paragraph style.

        target:
          - "last"   → style the LAST paragraph in the document body (use this
            right after insert_text(... position='end'); avoids the
            off-by-one-paragraph problem with view-cursor).
          - None     → use start/end if given, else current selection / view cursor.

        start, end: char range (end-exclusive). All paragraphs touching the
        range will get the style applied.
        """
        doc, err = self._require_writer()
        if err:
            return err
        # Pre-check style exists — gives agent a useful error with the
        # available list instead of an empty exception.
        try:
            para_styles = doc.getStyleFamilies().getByName("ParagraphStyles")
            if not para_styles.hasByName(style_name):
                return {"success": False,
                        "error": f"paragraph style not found: {style_name!r}",
                        "available": list(para_styles.getElementNames())}
        except Exception:
            pass
        try:
            if target == "last":
                rng = None
                enum = doc.getText().createEnumeration()
                while enum.hasMoreElements():
                    el = enum.nextElement()
                    if el.supportsService("com.sun.star.text.Paragraph"):
                        rng = el
            else:
                rng = self._resolve_range(doc, start, end)
            if rng is None:
                return {"success": False, "error": "No paragraph to style"}
            rng.ParaStyleName = style_name
            # Report what numbering actually attached — agent can detect when a style
            # exists by name but has no numbering rules, and decide to call apply_numbering.
            label = getattr(rng, "ListLabelString", "") or ""
            try:
                nr = rng.NumberingRules
                rule_name = getattr(nr, "Name", "") if nr else ""
            except Exception:
                rule_name = ""
            return {"success": True, "style": style_name, "target": target,
                    "effective_label": label,
                    "numbering_rule": rule_name}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def apply_numbering(self, level: int = 0, rule_name: str = None,
                        restart: bool = False, start_value: int = None,
                        is_number: bool = True,
                        start=None, end=None, target: str = None) -> Dict[str, Any]:
        """Configure auto-numbering on a paragraph (or range of paragraphs).

        - level: numbering depth (0 = top, 1 = sub, ...). UNO uses 0-indexed levels.
        - rule_name: name of a NumberingStyle (NumberingStyles family) to attach.
          If None, keeps the current rule (e.g. inherited from paragraph style).
        - restart: True to restart the counter at this paragraph.
        - start_value: explicit number to start from (only when restart=True).
        - is_number: False to skip numbering for this paragraph (in a list).
        - target: 'last' to operate on the LAST paragraph; otherwise start/end
          select paragraphs by char range; otherwise current selection / view cursor.
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            # Resolve target paragraph(s)
            if target == "last":
                rng = None
                enum = doc.getText().createEnumeration()
                while enum.hasMoreElements():
                    el = enum.nextElement()
                    if el.supportsService("com.sun.star.text.Paragraph"):
                        rng = el
            else:
                rng = self._resolve_range(doc, start, end)
            if rng is None:
                return {"success": False, "error": "No paragraph to apply numbering to"}

            applied = {}
            if rule_name is not None:
                try:
                    num_styles = doc.getStyleFamilies().getByName("NumberingStyles")
                except Exception:
                    num_styles = None
                if num_styles is None or not num_styles.hasByName(rule_name):
                    available = list(num_styles.getElementNames()) if num_styles else []
                    return {"success": False,
                            "error": f"numbering rule not found: {rule_name!r}",
                            "available": available}
                rng.NumberingRules = num_styles.getByName(rule_name).NumberingRules
                applied["rule_name"] = rule_name

            rng.NumberingLevel = int(level)
            rng.NumberingIsNumber = bool(is_number)
            applied["level"] = int(level)
            applied["is_number"] = bool(is_number)
            if restart:
                rng.ParaIsNumberingRestart = True
                applied["restart"] = True
                if start_value is not None:
                    rng.NumberingStartValue = int(start_value)
                    applied["start_value"] = int(start_value)

            # Read back the rendered label so the agent can verify
            label = getattr(rng, "ListLabelString", "") or ""
            return {"success": True, "applied": applied, "effective_label": label}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_numbering_styles(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            fam = doc.getStyleFamilies().getByName("NumberingStyles")
            return {"success": True, "styles": list(fam.getElementNames())}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def clone_numbering_rule(self, source_path: str, rule_name: str,
                             target_name: str = None) -> Dict[str, Any]:
        """Copy a NumberingStyle from a currently-open source doc into the active doc.

        After cloning, you can `apply_numbering(rule_name=target_name, ...)` on a
        paragraph in the active doc and the auto-numbering will render exactly
        as in the source doc (same level shapes, prefixes, separators).
        """
        import os, unicodedata
        from urllib.parse import unquote
        doc, err = self._require_writer()
        if err:
            return err
        try:
            try:
                want_real = os.path.realpath(source_path)
                want_nfc = unicodedata.normalize("NFC", want_real)
            except Exception:
                want_nfc = source_path

            src_doc = None
            comps = self.desktop.getComponents()
            it = comps.createEnumeration()
            while it.hasMoreElements():
                c = it.nextElement()
                u = ""
                try:
                    u = c.getURL() if hasattr(c, "getURL") else ""
                except Exception:
                    pass
                if not u or not u.startswith("file://"):
                    continue
                try:
                    local = unquote(u[len("file://"):])
                    local_real = os.path.realpath(local)
                    local_nfc = unicodedata.normalize("NFC", local_real)
                except Exception:
                    continue
                if local_nfc == want_nfc:
                    src_doc = c
                    break
            if src_doc is None:
                return {"success": False,
                        "error": f"source doc not currently open: {source_path!r}. "
                                 "Open it first with open_document_live."}

            src_fam = src_doc.getStyleFamilies().getByName("NumberingStyles")
            if not src_fam.hasByName(rule_name):
                return {"success": False,
                        "error": f"rule not found in source: {rule_name!r}",
                        "source_available": list(src_fam.getElementNames())}

            tgt_name = target_name or rule_name
            tgt_fam = doc.getStyleFamilies().getByName("NumberingStyles")
            if tgt_fam.hasByName(tgt_name):
                tgt_style = tgt_fam.getByName(tgt_name)
                created = False
            else:
                tgt_style = doc.createInstance("com.sun.star.style.NumberingStyle")
                tgt_fam.insertByName(tgt_name, tgt_style)
                created = True

            tgt_style.NumberingRules = src_fam.getByName(rule_name).NumberingRules
            return {"success": True, "rule_name": tgt_name, "created": created,
                    "source_rule": rule_name, "source_path": source_path}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _find_open_doc(self, source_path: str):
        """Walk Desktop.getComponents and find an open doc whose realpath
        matches source_path (NFC-normalized for macOS). Returns the model or None."""
        import os, unicodedata
        from urllib.parse import unquote
        try:
            want_real = os.path.realpath(source_path)
            want_nfc = unicodedata.normalize("NFC", want_real)
        except Exception:
            want_nfc = source_path
        comps = self.desktop.getComponents()
        it = comps.createEnumeration()
        while it.hasMoreElements():
            c = it.nextElement()
            try:
                u = c.getURL() if hasattr(c, "getURL") else ""
            except Exception:
                u = ""
            if not u or not u.startswith("file://"):
                continue
            try:
                local = unquote(u[len("file://"):])
                local_real = os.path.realpath(local)
                local_nfc = unicodedata.normalize("NFC", local_real)
            except Exception:
                continue
            if local_nfc == want_nfc:
                return c
        return None

    def clone_paragraph_style(self, source_path: str, style_name: str,
                              target_name: str = None,
                              overwrite: bool = True) -> Dict[str, Any]:
        """Copy a ParagraphStyle's properties from a currently-open source doc
        into the active doc.

        Copies font, size, bold/italic/underline, color, alignment, line spacing,
        paragraph margins/indents, tab stops, outline level, parent style.
        After cloning, apply_paragraph_style(target_name) inherits the source's
        layout — useful when source style names exist in target by name only
        (e.g. fresh LO 'Heading 1' has different defaults than a Word import).
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            src_doc = self._find_open_doc(source_path)
            if src_doc is None:
                return {"success": False,
                        "error": f"source doc not currently open: {source_path!r}. "
                                 "Open it first with open_document_live."}
            src_fam = src_doc.getStyleFamilies().getByName("ParagraphStyles")
            if not src_fam.hasByName(style_name):
                return {"success": False,
                        "error": f"style not found in source: {style_name!r}",
                        "source_available": list(src_fam.getElementNames())}
            src = src_fam.getByName(style_name)
            tgt_name = target_name or style_name
            tgt_fam = doc.getStyleFamilies().getByName("ParagraphStyles")
            if tgt_fam.hasByName(tgt_name):
                if not overwrite:
                    return {"success": False,
                            "error": f"target style exists: {tgt_name!r}; pass overwrite=True"}
                tgt = tgt_fam.getByName(tgt_name)
                created = False
            else:
                tgt = doc.createInstance("com.sun.star.style.ParagraphStyle")
                tgt_fam.insertByName(tgt_name, tgt)
                created = True

            # Copy a curated set of properties. Avoid blind setPropertyValue loop —
            # some props are read-only or interrelated (e.g. CharColor + CharColorTheme),
            # and writing them in arbitrary order can throw.
            props = [
                # Char
                "CharFontName", "CharHeight", "CharWeight", "CharPosture",
                "CharUnderline", "CharUnderlineColor", "CharUnderlineHasColor",
                "CharStrikeout", "CharOverline",
                "CharColor", "CharBackColor", "CharBackTransparent",
                "CharContoured", "CharShadowed", "CharRelief",
                "CharCaseMap", "CharWordMode", "CharKerning", "CharAutoKerning",
                "CharFontNameAsian", "CharHeightAsian", "CharWeightAsian", "CharPostureAsian",
                "CharFontNameComplex", "CharHeightComplex", "CharWeightComplex", "CharPostureComplex",
                "CharLocale", "CharLocaleAsian", "CharLocaleComplex",
                # Para
                "ParaAdjust", "ParaLastLineAdjust",
                "ParaLeftMargin", "ParaRightMargin",
                "ParaTopMargin", "ParaBottomMargin", "ParaContextMargin",
                "ParaFirstLineIndent", "ParaIsAutoFirstLineIndent",
                "ParaLineSpacing",
                "ParaTabStops",
                "ParaOrphans", "ParaWidows", "ParaKeepTogether",
                "ParaSplit",
                "ParaIsHyphenation", "ParaHyphenationMaxHyphens",
                "ParaHyphenationMaxLeadingChars", "ParaHyphenationMaxTrailingChars",
                "ParaRegisterModeActive",
                # Outline & numbering linkage
                "OutlineLevel",
                # Borders & background
                "TopBorder", "BottomBorder", "LeftBorder", "RightBorder",
                "BorderDistance", "TopBorderDistance", "BottomBorderDistance",
                "LeftBorderDistance", "RightBorderDistance",
                "ParaBackColor", "ParaBackTransparent",
                # Page break behavior
                "BreakType", "PageDescName", "PageNumberOffset",
                # Drop caps
                "DropCapFormat", "DropCapWholeWord",
            ]
            copied = []
            failed = []
            for name in props:
                try:
                    if not src.getPropertySetInfo().hasPropertyByName(name):
                        continue
                    if not tgt.getPropertySetInfo().hasPropertyByName(name):
                        continue
                    val = src.getPropertyValue(name)
                    tgt.setPropertyValue(name, val)
                    copied.append(name)
                except Exception as ex:
                    failed.append({"prop": name, "error": str(ex)})
            # Parent style — separate, set last
            try:
                parent = getattr(src, "ParentStyle", "") or ""
                if parent:
                    tgt_fam_names = list(tgt_fam.getElementNames())
                    if parent in tgt_fam_names:
                        tgt.ParentStyle = parent
            except Exception:
                pass

            return {"success": True, "style_name": tgt_name, "created": created,
                    "source_style": style_name, "source_path": source_path,
                    "copied_count": len(copied), "failed_count": len(failed),
                    "failed_props": failed[:5]}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def clone_page_style(self, source_path: str,
                         source_style: str = None,
                         target_style: str = "Default Page Style") -> Dict[str, Any]:
        """Copy a PageStyle from a currently-open source doc into the active doc.

        Copies page size/orientation, all 4 margins, header/footer enabled+text+margins,
        column count, footnote area, background. After cloning, the active doc's
        target page-style renders pages identically to the source's source_style
        (including header/footer zone reservation visible on the ruler).
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            src_doc = self._find_open_doc(source_path)
            if src_doc is None:
                return {"success": False,
                        "error": f"source doc not currently open: {source_path!r}. "
                                 "Open it first with open_document_live."}
            src_fam = src_doc.getStyleFamilies().getByName("PageStyles")
            if source_style is None:
                # Pick page-style of source's first paragraph (handles Word-import
                # 'MP0' etc.) — fall back to the doc's default if unset.
                pdn = self._first_paragraph_page_style(src_doc)
                source_style = pdn or "Default Page Style"
            if not src_fam.hasByName(source_style):
                # try locale fallbacks
                for cand in ("Standard", "Default Page Style", "Default Style"):
                    if src_fam.hasByName(cand):
                        source_style = cand; break
                if not src_fam.hasByName(source_style):
                    return {"success": False,
                            "error": f"page style not found in source: {source_style!r}",
                            "source_available": list(src_fam.getElementNames())}
            src = src_fam.getByName(source_style)
            tgt_fam = doc.getStyleFamilies().getByName("PageStyles")
            if not tgt_fam.hasByName(target_style):
                # locale fallback for target too
                for cand in ("Default Page Style", "Standard", "Default Style"):
                    if tgt_fam.hasByName(cand):
                        target_style = cand; break
                if not tgt_fam.hasByName(target_style):
                    return {"success": False,
                            "error": f"target page style not found: {target_style!r}",
                            "target_available": list(tgt_fam.getElementNames())}
            tgt = tgt_fam.getByName(target_style)

            # Header / Footer have to be enabled BEFORE copying their text /
            # margins, otherwise the slot is null and writes throw.
            try:
                if getattr(src, "HeaderIsOn", False):
                    tgt.HeaderIsOn = True
            except Exception:
                pass
            try:
                if getattr(src, "FooterIsOn", False):
                    tgt.FooterIsOn = True
            except Exception:
                pass

            props = [
                "Size", "IsLandscape",
                "TopMargin", "BottomMargin", "LeftMargin", "RightMargin",
                "BorderDistance",
                "BackColor", "BackTransparent",
                # Header
                "HeaderIsOn", "HeaderIsDynamicHeight", "HeaderIsShared",
                "HeaderHeight", "HeaderBodyDistance",
                "HeaderLeftMargin", "HeaderRightMargin",
                "HeaderBackColor", "HeaderBackTransparent",
                # Footer
                "FooterIsOn", "FooterIsDynamicHeight", "FooterIsShared",
                "FooterHeight", "FooterBodyDistance",
                "FooterLeftMargin", "FooterRightMargin",
                "FooterBackColor", "FooterBackTransparent",
                # Borders
                "TopBorder", "BottomBorder", "LeftBorder", "RightBorder",
                "TopBorderDistance", "BottomBorderDistance",
                "LeftBorderDistance", "RightBorderDistance",
                # Footnote area / columns
                "FootnoteHeight", "FootnoteLineWeight", "FootnoteLineColor",
                "FootnoteLineRelativeWidth", "FootnoteLineAdjust",
                "FootnoteLineTextDistance", "FootnoteLineDistance",
                "TextColumns",
                "PageStyleLayout",
                "RegisterModeActive",
            ]
            copied = []; failed = []
            for name in props:
                try:
                    if not src.getPropertySetInfo().hasPropertyByName(name): continue
                    if not tgt.getPropertySetInfo().hasPropertyByName(name): continue
                    val = src.getPropertyValue(name)
                    tgt.setPropertyValue(name, val)
                    copied.append(name)
                except Exception as ex:
                    failed.append({"prop": name, "error": str(ex)})

            # Header/Footer text bodies are XText objects — copy via setString
            try:
                if getattr(src, "HeaderIsOn", False) and getattr(tgt, "HeaderIsOn", False):
                    h_txt = src.HeaderText.getString()
                    tgt.HeaderText.setString(h_txt)
            except Exception as ex:
                failed.append({"prop": "HeaderText", "error": str(ex)})
            try:
                if getattr(src, "FooterIsOn", False) and getattr(tgt, "FooterIsOn", False):
                    f_txt = src.FooterText.getString()
                    tgt.FooterText.setString(f_txt)
            except Exception as ex:
                failed.append({"prop": "FooterText", "error": str(ex)})

            return {"success": True, "source_style": source_style,
                    "target_style": target_style,
                    "header_enabled": bool(getattr(tgt, "HeaderIsOn", False)),
                    "footer_enabled": bool(getattr(tgt, "FooterIsOn", False)),
                    "copied_count": len(copied), "failed_count": len(failed),
                    "failed_props": failed[:5]}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def find_and_replace(self, search: str, replace: str = "",
                         regex: bool = False, case_sensitive: bool = False) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            desc = doc.createReplaceDescriptor()
            desc.SearchString = search
            desc.ReplaceString = replace
            desc.SearchRegularExpression = bool(regex)
            desc.SearchCaseSensitive = bool(case_sensitive)
            count = doc.replaceAll(desc)
            return {"success": True, "replacements": count}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def delete_range(self, start: int, end: int) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        if end <= start:
            return {"success": False, "error": "end must be > start"}
        try:
            text = doc.getText()
            cursor = text.createTextCursor()
            cursor.gotoStart(False)
            cursor.goRight(int(start), False)
            cursor.goRight(int(end - start), True)
            cursor.setString("")
            return {"success": True, "deleted_chars": end - start}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ---- Read-only inspection helpers ------------------------------------

    @staticmethod
    def _int_to_hex(color_int):
        if color_int is None or color_int < 0:
            return None
        return f"#{int(color_int) & 0xFFFFFF:06X}"

    def _iter_paragraphs(self, doc):
        """Yield (paragraph, char_offset_start, char_length) over the doc body."""
        text = doc.getText()
        enum = text.createEnumeration()
        offset = 0
        while enum.hasMoreElements():
            elem = enum.nextElement()
            if elem.supportsService("com.sun.star.text.Paragraph"):
                s = elem.getString()
                yield elem, offset, len(s)
                offset += len(s) + 1  # +1 for the implicit paragraph break
            else:
                # Tables and other contents — skip but advance
                try:
                    s = elem.getString()
                    offset += len(s) + 1
                except Exception:
                    offset += 1

    def get_paragraphs(self, start: int = 0, count: int = None,
                       include_format: bool = True, preview_chars: int = 80) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            out = []
            for idx, (para, off, length) in enumerate(self._iter_paragraphs(doc)):
                if idx < start:
                    continue
                if count is not None and len(out) >= count:
                    break
                s = para.getString()
                entry = {
                    "index": idx,
                    "start": off,
                    "end": off + length,
                    "length": length,
                    "preview": s[:preview_chars] + ("…" if len(s) > preview_chars else ""),
                }
                if include_format:
                    try:
                        entry["style"] = para.ParaStyleName
                        entry["alignment"] = ["left", "right", "justify", "center"][para.ParaAdjust] if para.ParaAdjust in (0,1,2,3) else str(para.ParaAdjust)
                        entry["left_mm"] = para.ParaLeftMargin / 100.0
                        entry["right_mm"] = para.ParaRightMargin / 100.0
                        entry["first_line_mm"] = para.ParaFirstLineIndent / 100.0
                        entry["top_mm"] = getattr(para, "ParaTopMargin", 0) / 100.0
                        entry["bottom_mm"] = getattr(para, "ParaBottomMargin", 0) / 100.0
                        try:
                            entry["tab_stops"] = self._encode_tab_stops(para.ParaTabStops)
                        except Exception:
                            entry["tab_stops"] = []
                        ls = para.ParaLineSpacing
                        entry["line_spacing"] = {"mode": ["proportional","minimum","leading","fix"][ls.Mode] if ls.Mode in (0,1,2,3) else ls.Mode,
                                                 "value": ls.Height if ls.Mode == 0 else ls.Height / 100.0}
                        try:
                            entry["context_margin"] = bool(getattr(para, "ParaContextMargin", False))
                        except Exception:
                            entry["context_margin"] = False
                        # Page style override: paragraphs that start a new page section
                        # (PageDescName != "") force a different page-style on the
                        # following pages — agent must read this to reproduce
                        # multi-page-style docs (e.g. 'First Page' for cover, then
                        # 'Default').
                        try:
                            entry["page_desc_name"] = getattr(para, "PageDescName", "") or ""
                        except Exception:
                            entry["page_desc_name"] = ""
                        try:
                            bt = getattr(para, "BreakType", 0)
                            # BreakType is a UNO Enum (com.sun.star.style.BreakType).
                            # 0=NONE, 1=COLUMN_BEFORE, 2=COLUMN_AFTER, 3=COLUMN_BOTH,
                            # 4=PAGE_BEFORE, 5=PAGE_AFTER, 6=PAGE_BOTH
                            if hasattr(bt, "value"):
                                entry["break_type"] = int(bt.value)
                            else:
                                entry["break_type"] = int(bt) if isinstance(bt, (int, float)) else 0
                        except Exception:
                            entry["break_type"] = 0
                        entry["list_label"] = getattr(para, "ListLabelString", "") or ""
                        entry["numbering_level"] = int(getattr(para, "NumberingLevel", 0) or 0)
                        entry["numbering_is_number"] = bool(getattr(para, "NumberingIsNumber", False))
                        try:
                            nr = para.NumberingRules
                            entry["numbering_rule_name"] = getattr(nr, "Name", "") if nr else ""
                        except Exception:
                            entry["numbering_rule_name"] = ""
                    except Exception as fe:
                        entry["format_error"] = str(fe)
                out.append(entry)
            return {"success": True, "paragraphs": out, "returned": len(out)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_paragraph_format_at(self, position: int) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            for idx, (para, off, length) in enumerate(self._iter_paragraphs(doc)):
                if off <= position <= off + length:
                    ls = para.ParaLineSpacing
                    return {"success": True, "paragraph": {
                        "index": idx, "start": off, "end": off + length, "length": length,
                        "style": para.ParaStyleName,
                        "alignment": ["left", "right", "justify", "center"][para.ParaAdjust] if para.ParaAdjust in (0,1,2,3) else str(para.ParaAdjust),
                        "left_mm": para.ParaLeftMargin / 100.0,
                        "right_mm": para.ParaRightMargin / 100.0,
                        "first_line_mm": para.ParaFirstLineIndent / 100.0,
                        "line_spacing_mode": ["proportional","minimum","leading","fix"][ls.Mode] if ls.Mode in (0,1,2,3) else ls.Mode,
                        "line_spacing_value": ls.Height if ls.Mode == 0 else ls.Height / 100.0,
                    }}
            return {"success": False, "error": "position out of range"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_outline(self, max_level: int = 10) -> Dict[str, Any]:
        """Return headings of the active Writer doc — paragraphs with OutlineLevel>0
        or whose style name starts with 'Heading'/'Заголовок'/'Title'. Cheap way to
        build a TOC without scanning the full body."""
        doc, err = self._require_writer()
        if err:
            return err
        try:
            out = []
            for idx, (para, off, length) in enumerate(self._iter_paragraphs(doc)):
                level = 0
                try:
                    level = int(getattr(para, "OutlineLevel", 0) or 0)
                except Exception:
                    level = 0
                style = ""
                try:
                    style = para.ParaStyleName or ""
                except Exception:
                    pass
                style_lc = style.lower()
                heading_by_style = (
                    style_lc.startswith("heading")
                    or style_lc.startswith("заголовок")
                    or style_lc == "title"
                    or style_lc == "subtitle"
                )
                if level <= 0 and not heading_by_style:
                    continue
                if level > 0 and level > max_level:
                    continue
                # Derive level from style name if OutlineLevel is missing
                if level <= 0 and heading_by_style:
                    digits = "".join(c for c in style if c.isdigit())
                    level = int(digits) if digits else 1
                text = para.getString()
                out.append({
                    "index": idx,
                    "level": level,
                    "style": style,
                    "start": off,
                    "end": off + length,
                    "text": text,
                })
            return {"success": True, "outline": out, "count": len(out)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_paragraphs_with_runs(self, start: int = 0, count: int = None,
                                 include_para_format: bool = True) -> Dict[str, Any]:
        """Like get_paragraphs but also returns inline character runs (text
        portions with uniform formatting). Each run carries font, size,
        bold/italic/underline, color, and hyperlink URL when present.
        Use this for faithful Markdown/HTML export when inline formatting matters."""
        doc, err = self._require_writer()
        if err:
            return err
        try:
            out = []
            for idx, (para, off, length) in enumerate(self._iter_paragraphs(doc)):
                if idx < start:
                    continue
                if count is not None and len(out) >= count:
                    break
                entry = {
                    "index": idx,
                    "start": off,
                    "end": off + length,
                    "text": para.getString(),
                }
                if include_para_format:
                    try:
                        entry["style"] = para.ParaStyleName
                        entry["outline_level"] = int(getattr(para, "OutlineLevel", 0) or 0)
                        entry["alignment"] = ["left", "right", "justify", "center"][para.ParaAdjust] if para.ParaAdjust in (0,1,2,3) else str(para.ParaAdjust)
                        entry["left_mm"] = para.ParaLeftMargin / 100.0
                        entry["right_mm"] = para.ParaRightMargin / 100.0
                        entry["first_line_mm"] = para.ParaFirstLineIndent / 100.0
                        entry["top_mm"] = getattr(para, "ParaTopMargin", 0) / 100.0
                        entry["bottom_mm"] = getattr(para, "ParaBottomMargin", 0) / 100.0
                        try:
                            entry["tab_stops"] = self._encode_tab_stops(para.ParaTabStops)
                        except Exception:
                            entry["tab_stops"] = []
                        # Numbering: agent needs the rendered label and rule name to faithfully
                        # replicate auto-numbered paragraphs (Heading 1 → '1.', List Paragraph → '2.1.1.')
                        entry["list_label"] = getattr(para, "ListLabelString", "") or ""
                        entry["numbering_level"] = int(getattr(para, "NumberingLevel", 0) or 0)
                        entry["numbering_is_number"] = bool(getattr(para, "NumberingIsNumber", False))
                        try:
                            nr = para.NumberingRules
                            entry["numbering_rule_name"] = getattr(nr, "Name", "") if nr else ""
                        except Exception:
                            entry["numbering_rule_name"] = ""
                    except Exception as fe:
                        entry["format_error"] = str(fe)
                # Enumerate text portions inside the paragraph
                runs = []
                try:
                    pen = para.createEnumeration()
                    while pen.hasMoreElements():
                        portion = pen.nextElement()
                        try:
                            ptype = getattr(portion, "TextPortionType", "Text")
                        except Exception:
                            ptype = "Text"
                        s = portion.getString()
                        if not s and ptype == "Text":
                            continue
                        run = {"type": ptype, "text": s}
                        try:
                            run["font_name"] = portion.CharFontName
                            run["font_size"] = portion.CharHeight
                            run["bold"] = portion.CharWeight >= 150
                            # CharPosture: NONE=0, OBLIQUE=1, ITALIC=2, DONTKNOW=4, REVERSE_OBLIQUE=5, REVERSE_ITALIC=6
                            # Treat only true italic values as italic. OBLIQUE renders slanted but is rare and
                            # was producing false positives where every body run came back italic=true.
                            run["italic"] = portion.CharPosture in (2, 6)
                            run["underline"] = portion.CharUnderline != 0
                            run["strike"] = bool(getattr(portion, "CharStrikeout", 0))
                            run["color"] = self._int_to_hex(portion.CharColor)
                            if getattr(portion, "CharBackColor", -1) not in (-1, 0xFFFFFFFF):
                                run["background_color"] = self._int_to_hex(portion.CharBackColor)
                            url = getattr(portion, "HyperLinkURL", "")
                            if url:
                                run["hyperlink"] = url
                            cstyle = getattr(portion, "CharStyleName", "")
                            if cstyle:
                                run["char_style"] = cstyle
                        except Exception as re:
                            run["run_error"] = str(re)
                        runs.append(run)
                except Exception as ee:
                    entry["runs_error"] = str(ee)
                entry["runs"] = runs
                out.append(entry)
            return {"success": True, "paragraphs": out, "returned": len(out)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_character_format(self, start: int, end: int = None) -> Dict[str, Any]:
        """Read character format on [start, end). If end is None or end==start, samples one char at start."""
        doc, err = self._require_writer()
        if err:
            return err
        try:
            text = doc.getText()
            cursor = text.createTextCursor()
            cursor.gotoStart(False)
            cursor.goRight(int(start), False)
            length = 1 if end is None or end == start else int(end - start)
            cursor.goRight(length, True)
            return {"success": True, "format": {
                "start": start,
                "end": start + length,
                "text": cursor.getString(),
                "font_name": cursor.CharFontName,
                "font_size": cursor.CharHeight,
                "bold": cursor.CharWeight >= 150,
                "italic": cursor.CharPosture in (2, 6),
                "underline": cursor.CharUnderline != 0,
                "color": self._int_to_hex(cursor.CharColor),
                "background_color": self._int_to_hex(cursor.CharBackColor),
            }}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_paragraph_styles(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            families = doc.getStyleFamilies()
            para = families.getByName("ParagraphStyles")
            names = list(para.getElementNames())
            return {"success": True, "styles": names, "count": len(names)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_paragraph_style_def(self, style_name: str) -> Dict[str, Any]:
        """Read the resolved properties of a paragraph style.

        Use this to figure out what 'Heading 1' / 'Body Text' actually looks
        like in the current doc (font, size, weight, alignment, indents).
        Lets the agent replicate a style's effect via direct format ops when
        the source doc's style name doesn't exist in the target doc.
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            families = doc.getStyleFamilies()
            para = families.getByName("ParagraphStyles")
            if not para.hasByName(style_name):
                return {"success": False,
                        "error": f"style not found: {style_name!r}",
                        "available": list(para.getElementNames())}
            st = para.getByName(style_name)
            align_map_rev = {0: "left", 1: "right", 2: "justify", 3: "center"}
            posture = getattr(st, "CharPosture", 0)
            d = {
                "name": st.Name,
                "display_name": getattr(st, "DisplayName", st.Name),
                "parent": getattr(st, "ParentStyle", "") or "",
                "follow": getattr(st, "FollowStyle", "") or "",
                "font_name": getattr(st, "CharFontName", None),
                "font_size": getattr(st, "CharHeight", None),
                "bold": getattr(st, "CharWeight", 100) >= 150,
                "italic": posture in (2, 6),
                "underline": getattr(st, "CharUnderline", 0) != 0,
                "char_word_mode": bool(getattr(st, "CharWordMode", False)),
                "alignment": align_map_rev.get(getattr(st, "ParaAdjust", 0), "left"),
                "left_mm": getattr(st, "ParaLeftMargin", 0) / 100.0,
                "right_mm": getattr(st, "ParaRightMargin", 0) / 100.0,
                "first_line_mm": getattr(st, "ParaFirstLineIndent", 0) / 100.0,
                "top_mm": getattr(st, "ParaTopMargin", 0) / 100.0,
                "bottom_mm": getattr(st, "ParaBottomMargin", 0) / 100.0,
                "context_margin": bool(getattr(st, "ParaContextMargin", False)),
                "outline_level": getattr(st, "OutlineLevel", 0),
                "keep_together": bool(getattr(st, "ParaKeepTogether", False)),
                "split_paragraph": bool(getattr(st, "ParaSplit", True)),
                "orphans": int(getattr(st, "ParaOrphans", 0) or 0),
                "widows": int(getattr(st, "ParaWidows", 0) or 0),
            }
            try:
                d["color"] = self._int_to_hex(st.CharColor)
            except Exception:
                pass
            try:
                ls = st.ParaLineSpacing
                d["line_spacing"] = {
                    "mode": ["proportional","minimum","leading","fix"][ls.Mode] if ls.Mode in (0,1,2,3) else ls.Mode,
                    "value": ls.Height if ls.Mode == 0 else ls.Height / 100.0,
                }
            except Exception:
                pass
            try:
                d["tab_stops"] = self._encode_tab_stops(st.ParaTabStops)
            except Exception:
                d["tab_stops"] = []
            return {"success": True, "style": d}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_style_props(self, style_name: str, **props) -> Dict[str, Any]:
        """Symmetric writer for get_paragraph_style_def. Accepts any subset of:
        font_name, font_size, bold, italic, underline, color (#RRGGBB), char_word_mode,
        alignment ('left'|'right'|'justify'|'center'),
        left_mm, right_mm, first_line_mm, top_mm, bottom_mm, context_margin,
        line_spacing ({mode, value}), tab_stops (list of {position_mm, alignment, ...}),
        outline_level, keep_together, split_paragraph, orphans, widows,
        parent, follow.

        Modifies the style in place — propagates to every paragraph using it.
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            fam = doc.getStyleFamilies().getByName("ParagraphStyles")
            if not fam.hasByName(style_name):
                return {"success": False, "error": f"style not found: {style_name!r}",
                        "available": list(fam.getElementNames())}
            st = fam.getByName(style_name)
            applied = {}
            failed = {}
            def _try(fn, key, val):
                try: fn(val); applied[key] = val
                except Exception as ex: failed[key] = str(ex)
            if "font_name" in props:
                _try(lambda v: setattr(st, "CharFontName", v), "font_name", props["font_name"])
            if "font_size" in props:
                _try(lambda v: setattr(st, "CharHeight", float(v)), "font_size", props["font_size"])
            if "bold" in props:
                _try(lambda v: setattr(st, "CharWeight", 150.0 if v else 100.0), "bold", bool(props["bold"]))
            if "italic" in props:
                _try(lambda v: setattr(st, "CharPosture", 2 if v else 0), "italic", bool(props["italic"]))
            if "underline" in props:
                _try(lambda v: setattr(st, "CharUnderline", 1 if v else 0), "underline", bool(props["underline"]))
            if "char_word_mode" in props:
                _try(lambda v: setattr(st, "CharWordMode", bool(v)), "char_word_mode", bool(props["char_word_mode"]))
            if "color" in props:
                _try(lambda v: setattr(st, "CharColor", int(str(v).lstrip("#"), 16)), "color", props["color"])
            if "alignment" in props:
                a_map = {"left":0,"right":1,"justify":2,"center":3}
                v = a_map.get(str(props["alignment"]).lower())
                if v is not None: _try(lambda x: setattr(st, "ParaAdjust", x), "alignment", v)
            for k_in, k_out, scale in [
                ("left_mm","ParaLeftMargin",100),
                ("right_mm","ParaRightMargin",100),
                ("first_line_mm","ParaFirstLineIndent",100),
                ("top_mm","ParaTopMargin",100),
                ("bottom_mm","ParaBottomMargin",100),
            ]:
                if k_in in props:
                    _try(lambda v: setattr(st, k_out, int(float(v)*scale)), k_in, props[k_in])
            if "context_margin" in props:
                _try(lambda v: setattr(st, "ParaContextMargin", bool(v)), "context_margin", props["context_margin"])
            if "outline_level" in props:
                _try(lambda v: setattr(st, "OutlineLevel", int(v)), "outline_level", props["outline_level"])
            if "keep_together" in props:
                _try(lambda v: setattr(st, "ParaKeepTogether", bool(v)), "keep_together", props["keep_together"])
            if "split_paragraph" in props:
                _try(lambda v: setattr(st, "ParaSplit", bool(v)), "split_paragraph", props["split_paragraph"])
            if "orphans" in props:
                _try(lambda v: setattr(st, "ParaOrphans", int(v)), "orphans", props["orphans"])
            if "widows" in props:
                _try(lambda v: setattr(st, "ParaWidows", int(v)), "widows", props["widows"])
            if "parent" in props:
                _try(lambda v: setattr(st, "ParentStyle", v or ""), "parent", props["parent"])
            if "follow" in props:
                _try(lambda v: setattr(st, "FollowStyle", v or ""), "follow", props["follow"])
            if "line_spacing" in props:
                ls_in = props["line_spacing"] or {}
                mode_map = {"proportional":0,"minimum":1,"leading":2,"fix":3}
                mode = mode_map.get(str(ls_in.get("mode","proportional")).lower(), 0)
                val = float(ls_in.get("value", 100))
                ls = uno.createUnoStruct("com.sun.star.style.LineSpacing")
                ls.Mode = mode
                ls.Height = int(val) if mode == 0 else int(val * 100)
                _try(lambda v: setattr(st, "ParaLineSpacing", v), "line_spacing", ls_in)
                try: st.ParaLineSpacing = ls
                except Exception as ex: failed["line_spacing"] = str(ex)
            if "tab_stops" in props:
                stops = props["tab_stops"] or []
                a_map = {"left":0,"center":1,"right":2,"decimal":3}
                tab_structs = []
                try:
                    for s in stops:
                        t = uno.createUnoStruct("com.sun.star.style.TabStop")
                        t.Position = int(float(s.get("position_mm", 0)) * 100)
                        t.Alignment = a_map.get(str(s.get("alignment","left")).lower(), 0)
                        fill = s.get("fill_char", " ") or " "
                        t.FillChar = ord(fill[0]) if isinstance(fill,str) and fill else 32
                        dec = s.get("decimal_char", ".") or "."
                        t.DecimalChar = ord(dec[0]) if isinstance(dec,str) and dec else 46
                        tab_structs.append(t)
                    st.ParaTabStops = tuple(tab_structs)
                    applied["tab_stops"] = stops
                except Exception as ex:
                    failed["tab_stops"] = str(ex)
            return {"success": True, "style_name": style_name,
                    "applied": list(applied.keys()), "failed": failed}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_character_styles(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            families = doc.getStyleFamilies()
            char = families.getByName("CharacterStyles")
            names = list(char.getElementNames())
            return {"success": True, "styles": names, "count": len(names)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def find_all(self, search: str, regex: bool = False, case_sensitive: bool = False,
                 max_results: int = 200) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            desc = doc.createSearchDescriptor()
            desc.SearchString = search
            desc.SearchRegularExpression = bool(regex)
            desc.SearchCaseSensitive = bool(case_sensitive)
            found = doc.findAll(desc)
            results = []
            full_text = doc.getText().getString()
            for i in range(found.getCount()):
                if i >= max_results:
                    break
                rng = found.getByIndex(i)
                snippet = rng.getString()
                # Best-effort start offset by string scan (UNO doesn't give absolute index directly)
                # For accurate indices, scan once globally:
                results.append({"text": snippet, "length": len(snippet)})
            # Also compute absolute offsets via single text scan
            if results:
                positions = []
                idx = 0
                import re
                if regex:
                    flags = 0 if case_sensitive else re.IGNORECASE
                    for m in re.finditer(search, full_text, flags=flags):
                        positions.append({"start": m.start(), "end": m.end(), "text": m.group(0)})
                        if len(positions) >= max_results:
                            break
                else:
                    needle = search if case_sensitive else search.lower()
                    hay = full_text if case_sensitive else full_text.lower()
                    while True:
                        i = hay.find(needle, idx)
                        if i < 0:
                            break
                        positions.append({"start": i, "end": i + len(search), "text": full_text[i:i+len(search)]})
                        idx = i + len(search) if len(search) else i + 1
                        if len(positions) >= max_results:
                            break
                return {"success": True, "matches": positions, "count": len(positions)}
            return {"success": True, "matches": [], "count": 0}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _first_paragraph_page_style(self, doc) -> str:
        """Return the PageDescName of the doc's first paragraph, or '' if not set.
        Word imports often anchor a Master-Page style (e.g. 'MP0') here that
        differs from 'Default Page Style' / 'Standard'."""
        try:
            enum = doc.getText().createEnumeration()
            while enum.hasMoreElements():
                el = enum.nextElement()
                if el.supportsService("com.sun.star.text.Paragraph"):
                    pdn = getattr(el, "PageDescName", "") or ""
                    return pdn
        except Exception:
            pass
        return ""

    def get_page_info(self, page_style: str = None) -> Dict[str, Any]:
        """Page-style metrics. If page_style is None, picks the page-style of
        the first paragraph (PageDescName) — required for Word-imports where
        page1 uses a Master-Page (e.g. 'MP0') with different margins from
        'Default Page Style'/'Standard'."""
        doc, err = self._require_writer()
        if err:
            return err
        if page_style is None or page_style == "":
            pdn = self._first_paragraph_page_style(doc)
            page_style = pdn or "Default Page Style"
        try:
            ctrl = doc.getCurrentController()
            page_count = None
            try:
                page_count = doc.getPropertyValue("PageCount")
            except Exception:
                if hasattr(ctrl, "getPageCount"):
                    try:
                        page_count = ctrl.getPageCount()
                    except Exception:
                        pass
            # Fallback: walk page-bound view-cursor jumps
            if page_count is None:
                try:
                    vc = ctrl.getViewCursor()
                    if hasattr(vc, "jumpToLastPage") and hasattr(vc, "getPage"):
                        vc.jumpToLastPage()
                        page_count = vc.getPage()
                        vc.jumpToFirstPage()
                except Exception:
                    pass
            view_cursor = ctrl.getViewCursor()
            current_page = None
            try:
                current_page = view_cursor.getPage() if hasattr(view_cursor, "getPage") else None
            except Exception:
                pass
            out = {"success": True, "page_count": page_count, "current_page": current_page}
            try:
                ps = self._page_style(doc, page_style)
                out["page_style"] = getattr(ps, "Name", None)
                size = ps.Size
                out["page_width_mm"] = size.Width / 100.0
                out["page_height_mm"] = size.Height / 100.0
                out["top_margin_mm"] = getattr(ps, "TopMargin", 0) / 100.0
                out["bottom_margin_mm"] = getattr(ps, "BottomMargin", 0) / 100.0
                out["left_margin_mm"] = getattr(ps, "LeftMargin", 0) / 100.0
                out["right_margin_mm"] = getattr(ps, "RightMargin", 0) / 100.0
                try:
                    out["orientation"] = "landscape" if getattr(ps, "IsLandscape", False) else "portrait"
                except Exception:
                    pass
                # Header
                h_on = bool(getattr(ps, "HeaderIsOn", False))
                out["header_enabled"] = h_on
                if h_on:
                    out["header_height_mm"] = getattr(ps, "HeaderHeight", 0) / 100.0
                    out["header_body_distance_mm"] = getattr(ps, "HeaderBodyDistance", 0) / 100.0
                    out["header_left_margin_mm"] = getattr(ps, "HeaderLeftMargin", 0) / 100.0
                    out["header_right_margin_mm"] = getattr(ps, "HeaderRightMargin", 0) / 100.0
                    out["header_dynamic_height"] = bool(getattr(ps, "HeaderIsDynamicHeight", False))
                    out["header_shared"] = bool(getattr(ps, "HeaderIsShared", True))
                    try: out["header_text"] = ps.HeaderText.getString()
                    except Exception: out["header_text"] = ""
                # Footer
                f_on = bool(getattr(ps, "FooterIsOn", False))
                out["footer_enabled"] = f_on
                if f_on:
                    out["footer_height_mm"] = getattr(ps, "FooterHeight", 0) / 100.0
                    out["footer_body_distance_mm"] = getattr(ps, "FooterBodyDistance", 0) / 100.0
                    out["footer_left_margin_mm"] = getattr(ps, "FooterLeftMargin", 0) / 100.0
                    out["footer_right_margin_mm"] = getattr(ps, "FooterRightMargin", 0) / 100.0
                    out["footer_dynamic_height"] = bool(getattr(ps, "FooterIsDynamicHeight", False))
                    out["footer_shared"] = bool(getattr(ps, "FooterIsShared", True))
                    try: out["footer_text"] = ps.FooterText.getString()
                    except Exception: out["footer_text"] = ""
                # Columns
                try:
                    cols = ps.TextColumns
                    out["column_count"] = int(getattr(cols, "ColumnCount", 1) or 1)
                except Exception:
                    out["column_count"] = 1
            except Exception:
                pass
            return out
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_page_style_props(self, page_style: str = "Default Page Style", **props) -> Dict[str, Any]:
        """Symmetric writer for get_page_info. Accepts any subset of:
        page_width_mm, page_height_mm, orientation ('portrait'|'landscape'),
        top_margin_mm, bottom_margin_mm, left_margin_mm, right_margin_mm,
        header_enabled, header_height_mm, header_body_distance_mm,
        header_left_margin_mm, header_right_margin_mm, header_text,
        footer_enabled, footer_height_mm, footer_body_distance_mm,
        footer_left_margin_mm, footer_right_margin_mm, footer_text.

        For header/footer text writes — must enable first (or pass header_enabled=True).
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ps = self._page_style(doc, page_style)
            applied = {}; failed = {}
            def _try(fn, key, val):
                try: fn(val); applied[key] = val
                except Exception as ex: failed[key] = str(ex)
            # Page size: prefer setting explicit width/height through Size struct
            if "page_width_mm" in props or "page_height_mm" in props:
                try:
                    sz = ps.Size
                    if "page_width_mm" in props:
                        sz.Width = int(float(props["page_width_mm"]) * 100)
                        applied["page_width_mm"] = props["page_width_mm"]
                    if "page_height_mm" in props:
                        sz.Height = int(float(props["page_height_mm"]) * 100)
                        applied["page_height_mm"] = props["page_height_mm"]
                    ps.Size = sz
                except Exception as ex:
                    failed["size"] = str(ex)
            if "orientation" in props:
                _try(lambda v: setattr(ps, "IsLandscape", v == "landscape"),
                     "orientation", str(props["orientation"]).lower())
            for k_in, k_out in [
                ("top_margin_mm","TopMargin"), ("bottom_margin_mm","BottomMargin"),
                ("left_margin_mm","LeftMargin"), ("right_margin_mm","RightMargin"),
            ]:
                if k_in in props:
                    _try(lambda v: setattr(ps, k_out, int(float(v)*100)), k_in, props[k_in])
            # Header — must enable first; subsequent text/margin writes need the slot live
            if "header_enabled" in props:
                _try(lambda v: setattr(ps, "HeaderIsOn", bool(v)), "header_enabled", bool(props["header_enabled"]))
            for k_in, k_out in [
                ("header_height_mm","HeaderHeight"),
                ("header_body_distance_mm","HeaderBodyDistance"),
                ("header_left_margin_mm","HeaderLeftMargin"),
                ("header_right_margin_mm","HeaderRightMargin"),
            ]:
                if k_in in props and getattr(ps, "HeaderIsOn", False):
                    _try(lambda v: setattr(ps, k_out, int(float(v)*100)), k_in, props[k_in])
            if "header_text" in props and getattr(ps, "HeaderIsOn", False):
                _try(lambda v: ps.HeaderText.setString(v or ""), "header_text", props["header_text"])
            # Footer
            if "footer_enabled" in props:
                _try(lambda v: setattr(ps, "FooterIsOn", bool(v)), "footer_enabled", bool(props["footer_enabled"]))
            for k_in, k_out in [
                ("footer_height_mm","FooterHeight"),
                ("footer_body_distance_mm","FooterBodyDistance"),
                ("footer_left_margin_mm","FooterLeftMargin"),
                ("footer_right_margin_mm","FooterRightMargin"),
            ]:
                if k_in in props and getattr(ps, "FooterIsOn", False):
                    _try(lambda v: setattr(ps, k_out, int(float(v)*100)), k_in, props[k_in])
            if "footer_text" in props and getattr(ps, "FooterIsOn", False):
                _try(lambda v: ps.FooterText.setString(v or ""), "footer_text", props["footer_text"])
            return {"success": True, "page_style": ps.Name,
                    "applied": list(applied.keys()), "failed": failed}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_page_margins(self, top_mm: Optional[float] = None,
                         bottom_mm: Optional[float] = None,
                         left_mm: Optional[float] = None,
                         right_mm: Optional[float] = None,
                         page_style: str = "Default Page Style") -> Dict[str, Any]:
        """Set page margins (in mm) on a page style. Only provided fields are
        modified; others are left as-is. Affects every paragraph using that page
        style — page margins are NOT a paragraph property.
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ps = self._page_style(doc, page_style)
            applied = {}
            if top_mm is not None:
                ps.TopMargin = int(float(top_mm) * 100); applied["top_mm"] = top_mm
            if bottom_mm is not None:
                ps.BottomMargin = int(float(bottom_mm) * 100); applied["bottom_mm"] = bottom_mm
            if left_mm is not None:
                ps.LeftMargin = int(float(left_mm) * 100); applied["left_mm"] = left_mm
            if right_mm is not None:
                ps.RightMargin = int(float(right_mm) * 100); applied["right_mm"] = right_mm
            return {"success": True, "page_style": ps.Name, "applied": applied}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ---- Open / Recent documents ----------------------------------------

    @staticmethod
    def _path_to_url(path: str) -> str:
        if path.startswith(("file://", "private:")):
            return path
        try:
            return uno.systemPathToFileUrl(path)
        except Exception:
            from urllib.parse import quote
            return "file://" + quote(path)

    def open_document_live(self, path: str, readonly: bool = False) -> Dict[str, Any]:
        """Open an existing document on disk and keep it open.

        Uses the same Hidden=True workaround as create_document to avoid
        UI-thread deadlock when called from a background HTTP-server thread,
        then makes the window visible.

        Dedup: if a document with the same realpath is already open
        (NFC/NFD-normalized comparison), focus it instead of opening a duplicate.
        """
        try:
            import os, unicodedata
            try:
                real = os.path.realpath(path)
                want_nfc = unicodedata.normalize("NFC", real)
            except Exception:
                want_nfc = path
            # Walk all open components, compare normalized URLs
            try:
                comps = self.desktop.getComponents()
                it = comps.createEnumeration()
                while it.hasMoreElements():
                    c = it.nextElement()
                    try:
                        u = c.getURL() if hasattr(c, "getURL") else ""
                    except Exception:
                        u = ""
                    if not u or not u.startswith("file://"):
                        continue
                    try:
                        from urllib.parse import unquote
                        local = unquote(u[len("file://"):])
                        local_real = os.path.realpath(local)
                        local_nfc = unicodedata.normalize("NFC", local_real)
                    except Exception:
                        continue
                    if local_nfc == want_nfc:
                        self._last_active_doc = c
                        try:
                            ctrl = c.getCurrentController()
                            if ctrl is not None:
                                frame = ctrl.getFrame()
                                if frame is not None:
                                    win = frame.getContainerWindow()
                                    if win is not None:
                                        win.setVisible(True)
                        except Exception:
                            pass
                        return {
                            "success": True,
                            "url": u,
                            "type": self._get_document_type(c),
                            "readonly": readonly,
                            "deduplicated": True,
                        }
            except Exception as e:
                logger.warning(f"open_document_live dedup walk failed: {e}")

            url = self._path_to_url(path)
            props = []
            hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
            props.append(hidden)
            if readonly:
                ro = PropertyValue(); ro.Name = "ReadOnly"; ro.Value = True
                props.append(ro)
            doc = self.desktop.loadComponentFromURL(url, "_blank", 0, tuple(props))
            if doc is None:
                return {"success": False, "error": f"loadComponentFromURL returned None for {url}"}
            try:
                ctrl = doc.getCurrentController()
                if ctrl is not None:
                    frame = ctrl.getFrame()
                    if frame is not None:
                        win = frame.getContainerWindow()
                        if win is not None:
                            win.setVisible(True)
            except Exception as e:
                logger.warning(f"Opened document but could not show window: {e}")
            self._last_active_doc = doc
            return {
                "success": True,
                "url": doc.getURL() if hasattr(doc, "getURL") else url,
                "type": self._get_document_type(doc),
                "readonly": readonly,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _get_recent_pick_list(self):
        cp = self.smgr.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", self.ctx
        )
        nodepath = PropertyValue()
        nodepath.Name = "nodepath"
        nodepath.Value = "/org.openoffice.Office.Histories/Histories"
        node = cp.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess", (nodepath,)
        )
        return node.getByName("PickList")

    def list_recent_documents(self, max_items: int = 25) -> Dict[str, Any]:
        try:
            pick = self._get_recent_pick_list()
            order = list(pick.OrderList.getElementNames())
            items = pick.ItemList
            out = []
            for key in order[:max_items]:
                try:
                    entry = items.getByName(key)
                    url = entry.getPropertyValue("HistoryItemRef") if entry.getPropertySetInfo().hasPropertyByName("HistoryItemRef") else key
                    title = entry.getPropertyValue("Title") if entry.getPropertySetInfo().hasPropertyByName("Title") else ""
                    out.append({"url": url, "title": title, "key": key})
                except Exception:
                    out.append({"url": key, "title": "", "key": key})
            return {"success": True, "recent": out, "count": len(out)}
        except Exception as e:
            # Fallback: read item list directly (older configs)
            try:
                pick = self._get_recent_pick_list()
                items = pick.ItemList
                names = list(items.getElementNames())
                out = []
                for n in names[:max_items]:
                    out.append({"url": n, "title": "", "key": n})
                return {"success": True, "recent": out, "count": len(out)}
            except Exception as e2:
                return {"success": False, "error": f"{e}; fallback: {e2}"}

    def open_recent_document(self, index: int = 0, readonly: bool = False) -> Dict[str, Any]:
        rec = self.list_recent_documents()
        if not rec.get("success"):
            return rec
        items = rec.get("recent", [])
        if index < 0 or index >= len(items):
            return {"success": False, "error": f"index {index} out of range (have {len(items)} recent)"}
        return self.open_document_live(items[index]["url"], readonly=readonly)

    # ---- Document inspection (extended) ---------------------------------

    def get_document_metadata(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            p = doc.getDocumentProperties()
            def fmt(d):
                if d is None:
                    return None
                # com.sun.star.util.DateTime is a struct with Year/Month/...
                if hasattr(d, "Year"):
                    return f"{d.Year:04d}-{d.Month:02d}-{d.Day:02d}T{d.Hours:02d}:{d.Minutes:02d}:{d.Seconds:02d}"
                return str(d)
            return {"success": True, "metadata": {
                "title": p.Title,
                "subject": p.Subject,
                "author": p.Author,
                "description": p.Description,
                "keywords": list(p.Keywords) if p.Keywords else [],
                "language": str(p.Language) if p.Language else None,
                "creation_date": fmt(p.CreationDate),
                "modification_date": fmt(p.ModificationDate),
                "modified_by": p.ModifiedBy,
                "print_date": fmt(p.PrintDate),
                "printed_by": p.PrintedBy,
                "editing_cycles": p.EditingCycles,
                "editing_duration": p.EditingDuration,
                "generator": p.Generator,
            }}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_document_summary(self) -> Dict[str, Any]:
        """One-shot overview: counts of everything + metadata."""
        doc, err = self._require_writer()
        if err:
            return err
        try:
            text_str = doc.getText().getString()
            paragraphs = sum(1 for _ in self._iter_paragraphs(doc))
            tables_count = doc.getTextTables().getCount()
            try:
                images_count = doc.getGraphicObjects().getCount()
            except Exception:
                images_count = None
            try:
                bookmarks_count = doc.getBookmarks().getCount()
            except Exception:
                bookmarks_count = 0
            try:
                sections_count = doc.getTextSections().getCount()
            except Exception:
                sections_count = 0
            try:
                fields_count = doc.getTextFields().createEnumeration()
                cnt = 0
                while fields_count.hasMoreElements():
                    fields_count.nextElement()
                    cnt += 1
                fields_count = cnt
            except Exception:
                fields_count = None
            try:
                page_count = doc.getPropertyValue("PageCount")
            except Exception:
                page_count = None
            # comments are TextFields of type Annotation
            try:
                ann_count = 0
                e = doc.getTextFields().createEnumeration()
                while e.hasMoreElements():
                    f = e.nextElement()
                    if f.supportsService("com.sun.star.text.TextField.Annotation"):
                        ann_count += 1
            except Exception:
                ann_count = None
            # links — count unique URLs across paragraph portions
            try:
                links = self._collect_hyperlinks(doc, max_items=10000)
                links_count = len(links)
            except Exception:
                links_count = None
            return {"success": True, "summary": {
                "url": doc.getURL(),
                "title": doc.getDocumentProperties().Title or "",
                "char_count": len(text_str),
                "word_count": len(text_str.split()),
                "paragraph_count": paragraphs,
                "page_count": page_count,
                "table_count": tables_count,
                "image_count": images_count,
                "bookmark_count": bookmarks_count,
                "section_count": sections_count,
                "field_count": fields_count,
                "annotation_count": ann_count,
                "hyperlink_count": links_count,
            }}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_bookmarks(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            bms = doc.getBookmarks()
            full_text = doc.getText().getString()
            out = []
            for i in range(bms.getCount()):
                bm = bms.getByIndex(i)
                name = bm.getName()
                anchor = bm.getAnchor()
                snippet = anchor.getString()
                # Compute absolute char offset by string match: best-effort.
                start = None
                if snippet:
                    start = full_text.find(snippet)
                out.append({"name": name, "anchor_text": snippet, "approx_start": start})
            return {"success": True, "bookmarks": out, "count": len(out)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _collect_hyperlinks(self, doc, max_items=200):
        """Walk text portions; return [{url, text, ...}] for portions with HyperLinkURL set."""
        out = []
        text = doc.getText()
        para_enum = text.createEnumeration()
        while para_enum.hasMoreElements() and len(out) < max_items:
            para = para_enum.nextElement()
            if not para.supportsService("com.sun.star.text.Paragraph"):
                continue
            try:
                portions = para.createEnumeration()
            except Exception:
                continue
            while portions.hasMoreElements() and len(out) < max_items:
                p = portions.nextElement()
                try:
                    url = p.getPropertyValue("HyperLinkURL")
                except Exception:
                    url = ""
                if url:
                    out.append({"url": url, "text": p.getString(),
                                "target": getattr(p, "HyperLinkTarget", "") or ""})
        return out

    def list_hyperlinks(self, max_items: int = 200) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            return {"success": True, "hyperlinks": self._collect_hyperlinks(doc, max_items),
                    "count": len(self._collect_hyperlinks(doc, max_items))}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_comments(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            out = []
            e = doc.getTextFields().createEnumeration()
            while e.hasMoreElements():
                f = e.nextElement()
                if not f.supportsService("com.sun.star.text.TextField.Annotation"):
                    continue
                d = f.Date
                date_str = None
                try:
                    if hasattr(d, "Year"):
                        date_str = f"{d.Year:04d}-{d.Month:02d}-{d.Day:02d}"
                except Exception:
                    pass
                anchor_text = ""
                try:
                    anchor_text = f.getAnchor().getString()[:80]
                except Exception:
                    pass
                out.append({
                    "author": getattr(f, "Author", ""),
                    "initials": getattr(f, "Initials", ""),
                    "date": date_str,
                    "text": f.Content,
                    "anchor_preview": anchor_text,
                })
            return {"success": True, "comments": out, "count": len(out)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_images(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            out = []
            try:
                imgs = doc.getGraphicObjects()
                for i in range(imgs.getCount()):
                    g = imgs.getByIndex(i)
                    sz = getattr(g, "Size", None)
                    out.append({
                        "name": g.getName() if hasattr(g, "getName") else "",
                        "width_mm": sz.Width / 100.0 if sz else None,
                        "height_mm": sz.Height / 100.0 if sz else None,
                        "anchor_type": str(getattr(g, "AnchorType", "")),
                    })
            except Exception:
                pass
            # Also walk DrawPage for shapes/embedded images not in GraphicObjects
            try:
                dp = doc.getDrawPage()
                seen = {x["name"] for x in out if x["name"]}
                for i in range(dp.getCount()):
                    s = dp.getByIndex(i)
                    nm = s.getName() if hasattr(s, "getName") else ""
                    if nm in seen:
                        continue
                    sz = getattr(s, "Size", None)
                    out.append({
                        "name": nm,
                        "width_mm": sz.Width / 100.0 if sz else None,
                        "height_mm": sz.Height / 100.0 if sz else None,
                        "shape_type": s.supportsService("com.sun.star.drawing.GraphicObjectShape") and "image" or "shape",
                    })
            except Exception:
                pass
            return {"success": True, "images": out, "count": len(out)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_sections(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            secs = doc.getTextSections()
            out = []
            for i in range(secs.getCount()):
                s = secs.getByIndex(i)
                try:
                    snippet = s.getAnchor().getString()[:80]
                except Exception:
                    snippet = ""
                out.append({
                    "name": s.getName() if hasattr(s, "getName") else "",
                    "is_protected": getattr(s, "IsProtected", False),
                    "is_visible": getattr(s, "IsVisible", True),
                    "preview": snippet,
                })
            return {"success": True, "sections": out, "count": len(out)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def read_table_cells(self, table_name: str = None, table_index: int = None) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            tables = doc.getTextTables()
            if table_name:
                if not tables.hasByName(table_name):
                    return {"success": False, "error": f"no table named '{table_name}'"}
                t = tables.getByName(table_name)
            elif table_index is not None:
                if table_index < 0 or table_index >= tables.getCount():
                    return {"success": False, "error": f"table_index out of range"}
                t = tables.getByIndex(table_index)
            else:
                if tables.getCount() == 0:
                    return {"success": False, "error": "no tables"}
                t = tables.getByIndex(0)
            rows = t.getRows().getCount()
            cols = t.getColumns().getCount()
            grid = []
            for r in range(rows):
                row_cells = []
                for c in range(cols):
                    try:
                        cell_name = chr(ord("A") + c) + str(r + 1)
                        cell = t.getCellByName(cell_name)
                        row_cells.append(cell.getString() if cell else "")
                    except Exception:
                        row_cells.append("")
                grid.append(row_cells)
            return {"success": True, "name": t.getName(), "rows": rows, "columns": cols, "cells": grid}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_selection(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            sel = doc.getCurrentController().getSelection()
            if not hasattr(sel, "getCount") or sel.getCount() == 0:
                return {"success": True, "has_selection": False, "ranges": []}
            ranges = []
            full_text = doc.getText().getString()
            for i in range(sel.getCount()):
                r = sel.getByIndex(i)
                s = r.getString()
                start = full_text.find(s) if s else None
                ranges.append({
                    "text": s,
                    "length": len(s),
                    "approx_start": start,
                    "approx_end": (start + len(s)) if start is not None else None,
                })
            return {"success": True, "has_selection": True, "ranges": ranges, "count": len(ranges)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_text_at(self, start: int, end: int) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            full = doc.getText().getString()
            if start < 0 or end > len(full) or start > end:
                return {"success": False, "error": f"range [{start},{end}) out of [0,{len(full)}]"}
            return {"success": True, "start": start, "end": end, "text": full[start:end]}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ---- Mutating counterparts ------------------------------------------

    def set_document_metadata(self, title: str = None, subject: str = None,
                              author: str = None, description: str = None,
                              keywords=None) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            p = doc.getDocumentProperties()
            applied = {}
            if title is not None:
                p.Title = title; applied["title"] = title
            if subject is not None:
                p.Subject = subject; applied["subject"] = subject
            if author is not None:
                p.Author = author; applied["author"] = author
            if description is not None:
                p.Description = description; applied["description"] = description
            if keywords is not None:
                p.Keywords = tuple(keywords); applied["keywords"] = list(keywords)
            return {"success": True, "applied": applied}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _cursor_at(self, doc, start: int, end: int = None):
        text = doc.getText()
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        cursor.goRight(int(start), False)
        if end is not None and end > start:
            cursor.goRight(int(end - start), True)
        return text, cursor

    def add_bookmark(self, name: str, start: int, end: int = None) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            text, cursor = self._cursor_at(doc, start, end)
            bm = doc.createInstance("com.sun.star.text.Bookmark")
            bm.setName(name)
            text.insertTextContent(cursor, bm, end is not None and end > start)
            return {"success": True, "name": name, "start": start, "end": end}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def remove_bookmark(self, name: str) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            bms = doc.getBookmarks()
            if not bms.hasByName(name):
                return {"success": False, "error": f"no bookmark named '{name}'"}
            bm = bms.getByName(name)
            doc.getText().removeTextContent(bm)
            return {"success": True, "removed": name}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_hyperlink(self, start: int, end: int, url: str, target: str = "") -> Dict[str, Any]:
        """Make characters [start, end) a hyperlink pointing to url."""
        doc, err = self._require_writer()
        if err:
            return err
        if end <= start:
            return {"success": False, "error": "end must be > start"}
        try:
            _, cursor = self._cursor_at(doc, start, end)
            cursor.HyperLinkURL = url
            if target:
                cursor.HyperLinkTarget = target
            return {"success": True, "url": url, "start": start, "end": end}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_comment(self, start: int, text: str, author: str = "Claude",
                    initials: str = "AI", end: int = None) -> Dict[str, Any]:
        """Insert an annotation (comment) anchored at [start, end) (or just at start)."""
        doc, err = self._require_writer()
        if err:
            return err
        try:
            text_obj, cursor = self._cursor_at(doc, start, end if end is not None else start)
            ann = doc.createInstance("com.sun.star.text.TextField.Annotation")
            ann.Author = author
            ann.Initials = initials
            ann.Content = text
            # attach
            text_obj.insertTextContent(cursor, ann, end is not None and end > start)
            return {"success": True, "anchor_start": start, "anchor_end": end, "length": len(text)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def insert_image(self, path: str, position: int = None,
                     width_mm: float = None, height_mm: float = None) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            url = self._path_to_url(path)
            graphic = doc.createInstance("com.sun.star.text.TextGraphicObject")
            # GraphicURL deprecated; use Graphic via GraphicProvider
            try:
                gp = self.smgr.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", self.ctx)
                pv = PropertyValue(); pv.Name = "URL"; pv.Value = url
                graphic.Graphic = gp.queryGraphic((pv,))
            except Exception:
                graphic.GraphicURL = url  # legacy fallback
            if width_mm:
                sz = uno.createUnoStruct("com.sun.star.awt.Size")
                sz.Width = int(width_mm * 100)
                sz.Height = int((height_mm or width_mm) * 100)
                graphic.Size = sz
            text_obj = doc.getText()
            if position is not None:
                _, cursor = self._cursor_at(doc, position)
            else:
                cursor = text_obj.createTextCursor()
                cursor.gotoEnd(False)
            text_obj.insertTextContent(cursor, graphic, False)
            return {"success": True, "path": path, "url": url}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def insert_table(self, rows: int, columns: int, position: int = None,
                     name: str = None) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            table = doc.createInstance("com.sun.star.text.TextTable")
            table.initialize(int(rows), int(columns))
            if name:
                table.setName(name)
            text_obj = doc.getText()
            if position is not None:
                _, cursor = self._cursor_at(doc, position)
            else:
                cursor = text_obj.createTextCursor()
                cursor.gotoEnd(False)
            text_obj.insertTextContent(cursor, table, False)
            return {"success": True, "name": table.getName(), "rows": rows, "columns": columns}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def write_table_cell(self, table_name: str, cell: str, value: str) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            tables = doc.getTextTables()
            if not tables.hasByName(table_name):
                return {"success": False, "error": f"no table named '{table_name}'"}
            t = tables.getByName(table_name)
            c = t.getCellByName(cell)
            if c is None:
                return {"success": False, "error": f"cell '{cell}' not found in '{table_name}'"}
            c.setString(value)
            return {"success": True, "table": table_name, "cell": cell, "length": len(value)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def remove_table(self, table_name: str) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            tables = doc.getTextTables()
            if not tables.hasByName(table_name):
                return {"success": False, "error": f"no table named '{table_name}'"}
            t = tables.getByName(table_name)
            doc.getText().removeTextContent(t)
            return {"success": True, "removed": table_name}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ---- Undo / Redo / Dispatch -----------------------------------------

    def _dispatch(self, doc, command: str, props=()) -> None:
        """Internal helper: execute a UNO command on the doc's frame."""
        helper = self.smgr.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", self.ctx
        )
        frame = doc.getCurrentController().getFrame()
        helper.executeDispatch(frame, command, "", 0, props)

    def undo(self, steps: int = 1) -> Dict[str, Any]:
        """Undo the last N edits — equivalent to pressing Cmd+Z N times.
        Implemented via .uno:Undo dispatch (UndoManager API blocks the
        background HTTP thread on UI thread)."""
        doc, err = self._require_writer()
        if err:
            return err
        try:
            done = 0
            for _ in range(int(steps)):
                self._dispatch(doc, ".uno:Undo")
                done += 1
            return {"success": True, "undone": done}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def redo(self, steps: int = 1) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            done = 0
            for _ in range(int(steps)):
                self._dispatch(doc, ".uno:Redo")
                done += 1
            return {"success": True, "redone": done}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_undo_history(self, limit: int = 20) -> Dict[str, Any]:
        """Lightweight history check — only flags, since reading title arrays
        from UndoManager blocks on UI thread in LO 26."""
        doc, err = self._require_writer()
        if err:
            return err
        try:
            um = doc.UndoManager
            return {
                "success": True,
                "undo_possible": um.isUndoPossible(),
                "redo_possible": um.isRedoPossible(),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    # Whitelist of UNO commands considered safe to dispatch from the
    # background HTTP-server thread on macOS Sequoia. Any other command is
    # refused with an explanatory error instead of risking an indefinite UI
    # thread block. Add to this list ONLY after verifying the command does
    # not open a modal dialog, save/export/print/close, or run user code.
    _ALLOWED_COMMANDS = {
        # ---- Character formatting ----
        ".uno:Bold", ".uno:Italic", ".uno:Underline", ".uno:UnderlineDouble",
        ".uno:Strikeout", ".uno:Overline",
        ".uno:Subscript", ".uno:Superscript",
        ".uno:DefaultCharStyle", ".uno:ResetAttributes",
        ".uno:Shadowed", ".uno:Outline",
        ".uno:UppercaseSelection", ".uno:LowercaseSelection",
        ".uno:Grow", ".uno:Shrink",  # font size +/- 1
        # ---- Paragraph formatting ----
        ".uno:LeftPara", ".uno:RightPara", ".uno:CenterPara", ".uno:JustifyPara",
        ".uno:DefaultBullet", ".uno:DefaultNumbering",
        ".uno:DecrementIndent", ".uno:IncrementIndent",
        ".uno:DecrementSubLevels", ".uno:IncrementSubLevels",
        ".uno:ParaspaceIncrease", ".uno:ParaspaceDecrease",
        # ---- Insertion (no UI dialog) ----
        ".uno:InsertPagebreak", ".uno:InsertColumnBreak", ".uno:InsertLinebreak",
        ".uno:InsertNonBreakingSpace", ".uno:InsertNarrowNoBreakSpace",
        ".uno:InsertHardHyphen", ".uno:InsertSoftHyphen",
        # ---- Navigation ----
        ".uno:GoToStartOfDoc", ".uno:GoToEndOfDoc",
        ".uno:GoToStartOfLine", ".uno:GoToEndOfLine",
        ".uno:GoToNextPara", ".uno:GoToPrevPara",
        ".uno:GoToNextPage", ".uno:GoToPreviousPage",
        ".uno:GoToNextWord", ".uno:GoToPrevWord",
        ".uno:GoUp", ".uno:GoDown", ".uno:GoLeft", ".uno:GoRight",
        # ---- Selection ----
        ".uno:SelectAll", ".uno:SelectWord", ".uno:SelectSentence",
        ".uno:SelectParagraph", ".uno:SelectLine",
        # ---- Editing (Cut/Copy/Paste, no clipboard dialog) ----
        ".uno:Cut", ".uno:Copy", ".uno:Paste",
        ".uno:Undo", ".uno:Redo",  # prefer dedicated `undo`/`redo` tools
        ".uno:Delete", ".uno:DelToStartOfWord", ".uno:DelToEndOfWord",
        ".uno:DelToStartOfLine", ".uno:DelToEndOfLine",
        ".uno:DelToStartOfPara", ".uno:DelToEndOfPara",
        # ---- View toggles (display-only, no modal dialog) ----
        ".uno:ControlCodes",      # toggle formatting marks (¶, ·, →)
        ".uno:Marks",             # toggle field shadings
        ".uno:SpellOnline",       # toggle live spell-check (red waves)
        ".uno:ViewBounds",        # toggle text boundaries
        ".uno:ViewFormFields",    # toggle form-field shadings
    }

    def dispatch_uno_command(self, command: str, properties: dict = None) -> Dict[str, Any]:
        """Execute a built-in UNO command (e.g. '.uno:Bold', '.uno:CenterPara', '.uno:GoToStartOfDoc').
        Save / Export / Print / Open / Close are blocked because they hang the
        background thread on macOS — use the LibreOffice menu / Cmd+S manually."""
        doc, err = self._require_writer()
        if err:
            return err
        try:
            if not command.startswith(".uno:"):
                command = ".uno:" + command
            if command not in self._ALLOWED_COMMANDS:
                return {"success": False, "error":
                    f"{command} is not in the allowed-list of safe UNO commands. "
                    f"To avoid hanging the MCP server thread on macOS, only commands "
                    f"that don't open dialogs / save / export / run macros are allowed. "
                    f"See dispatch_uno_command tool description for the full list "
                    f"({len(self._ALLOWED_COMMANDS)} commands). If you need this "
                    f"command, do it from the LibreOffice menu manually."}
            props = []
            if properties:
                for k, v in properties.items():
                    pv = PropertyValue()
                    pv.Name = k
                    pv.Value = v
                    props.append(pv)
            self._dispatch(doc, command, tuple(props))
            return {"success": True, "command": command}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ---- Headers / Footers ----------------------------------------------

    def _page_style(self, doc, name: str = "Default Page Style"):
        ps_family = doc.getStyleFamilies().getByName("PageStyles")
        # Try the requested name first; if not found, fall back to a sensible
        # default. LibreOffice on different locales uses different display names
        # ("Default Page Style", "Default Style", localized variants…).
        try:
            return ps_family.getByName(name)
        except Exception:
            pass
        names = list(ps_family.getElementNames())
        for candidate in ("Default Page Style", "Default Style", "Standard"):
            if candidate in names:
                return ps_family.getByName(candidate)
        # last resort: first style that starts with "Default" or just the first one
        for n in names:
            if n.startswith("Default") or n.startswith("Стандарт"):
                return ps_family.getByName(n)
        if names:
            return ps_family.getByName(names[0])
        raise RuntimeError("No page styles available")

    def enable_header(self, enabled: bool = True, page_style: str = "Default Page Style") -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ps = self._page_style(doc, page_style)
            ps.HeaderIsOn = bool(enabled)
            return {"success": True, "header_enabled": ps.HeaderIsOn}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def enable_footer(self, enabled: bool = True, page_style: str = "Default Page Style") -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ps = self._page_style(doc, page_style)
            ps.FooterIsOn = bool(enabled)
            return {"success": True, "footer_enabled": ps.FooterIsOn}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_header(self, text: str, page_style: str = "Default Page Style") -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ps = self._page_style(doc, page_style)
            if not ps.HeaderIsOn:
                ps.HeaderIsOn = True
            ps.HeaderText.setString(text)
            return {"success": True, "header_text_length": len(text)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_footer(self, text: str, page_style: str = "Default Page Style") -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ps = self._page_style(doc, page_style)
            if not ps.FooterIsOn:
                ps.FooterIsOn = True
            ps.FooterText.setString(text)
            return {"success": True, "footer_text_length": len(text)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_header(self, page_style: str = "Default Page Style") -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ps = self._page_style(doc, page_style)
            return {"success": True, "enabled": ps.HeaderIsOn,
                    "text": ps.HeaderText.getString() if ps.HeaderIsOn else ""}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_footer(self, page_style: str = "Default Page Style") -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ps = self._page_style(doc, page_style)
            return {"success": True, "enabled": ps.FooterIsOn,
                    "text": ps.FooterText.getString() if ps.FooterIsOn else ""}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ---------------------------------------------------------------------

    def get_tables_info(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            tables = doc.getTextTables()
            out = []
            for i in range(tables.getCount()):
                t = tables.getByIndex(i)
                out.append({
                    "name": t.getName(),
                    "rows": t.getRows().getCount(),
                    "columns": t.getColumns().getCount(),
                })
            return {"success": True, "tables": out, "count": len(out)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ---------------------------------------------------------------------

    # Filter names for storeToURL — see https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html
    _STORE_FILTERS = {
        "docx": "MS Word 2007 XML",
        "doc":  "MS Word 97",
        "odt":  "writer8",
        "ott":  "writer8_template",
        "rtf":  "Rich Text Format",
        "txt":  "Text",
        "html": "HTML (StarWriter)",
        "xhtml": "XHTML Writer File",
        "pdf":  "writer_pdf_Export",
        "epub": "EPUB",
        "xlsx": "Calc MS Excel 2007 XML",
        "xls":  "MS Excel 97",
        "ods":  "calc8",
        "csv":  "Text - txt - csv (StarCalc)",
        "pptx": "Impress MS PowerPoint 2007 XML",
        "ppt":  "MS PowerPoint 97",
        "odp":  "impress8",
    }

    def clone_document(self, source_path: str, target_path: str,
                       target_format: str = None) -> Dict[str, Any]:
        """Convert a file from one format to another via a hidden, transient
        LibreOffice component — bypasses the macOS UI-thread save deadlock
        because the component is never visible and never bound to the main
        AppKit run loop.

        target_format defaults to the target_path extension (e.g. .docx → docx).
        Returns target URL on success.
        """
        try:
            src_url = self._path_to_url(source_path)
            dst_url = self._path_to_url(target_path)
            ext = (target_format or "").lower().lstrip(".")
            if not ext:
                ext = target_path.rsplit(".", 1)[-1].lower() if "." in target_path else ""
            filter_name = self._STORE_FILTERS.get(ext)
            if not filter_name:
                return {"success": False, "error": f"unknown target format '{ext}'. Supported: {sorted(self._STORE_FILTERS.keys())}"}

            # Load source hidden — never attached to a visible frame
            hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
            macros = PropertyValue(); macros.Name = "MacroExecutionMode"; macros.Value = 0
            doc = self.desktop.loadComponentFromURL(src_url, "_blank", 0, (hidden, macros))
            if doc is None:
                return {"success": False, "error": f"loadComponentFromURL returned None for {src_url}"}
            try:
                f = PropertyValue(); f.Name = "FilterName"; f.Value = filter_name
                ow = PropertyValue(); ow.Name = "Overwrite"; ow.Value = True
                doc.storeToURL(dst_url, (f, ow))
            finally:
                try:
                    doc.close(True)
                except Exception:
                    try:
                        doc.dispose()
                    except Exception:
                        pass
            return {"success": True, "source": src_url, "target": dst_url, "filter": filter_name}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def export_active_document(self, target_path: str, target_format: str = None) -> Dict[str, Any]:
        """Export the currently active document to a new file via storeToURL.
        Same UI-thread caveat as save_document — kept here as a building block;
        on macOS prefer clone_document(source_on_disk → target).
        """
        doc = self.get_active_document()
        if doc is None:
            return {"success": False, "error": "no active document"}
        try:
            dst_url = self._path_to_url(target_path)
            ext = (target_format or "").lower().lstrip(".")
            if not ext:
                ext = target_path.rsplit(".", 1)[-1].lower() if "." in target_path else ""
            filter_name = self._STORE_FILTERS.get(ext)
            if not filter_name:
                return {"success": False, "error": f"unknown target format '{ext}'"}
            f = PropertyValue(); f.Name = "FilterName"; f.Value = filter_name
            ow = PropertyValue(); ow.Name = "Overwrite"; ow.Value = True
            doc.storeToURL(dst_url, (f, ow))
            return {"success": True, "target": dst_url, "filter": filter_name}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def lock_view(self) -> Dict[str, Any]:
        """Freeze view updates (lockControllers).

        Stops the document from dispatching change events to its controllers
        while a worker thread is mutating it. Keeps the window visible but
        avoids the macOS SolarMutex contention that triggers a deadlock when
        a worker thread bursts many writes against a visible doc.

        Pair with unlock_view(). execute_batch() uses these automatically.
        Re-entrant: lockControllers/unlockControllers maintain a counter.
        """
        doc = self.get_active_document()
        if not doc:
            return {"success": False, "error": "No active document"}
        try:
            doc.lockControllers()
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def unlock_view(self) -> Dict[str, Any]:
        doc = self.get_active_document()
        if not doc:
            return {"success": False, "error": "No active document"}
        try:
            doc.unlockControllers()
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def shutdown_application(self, force: bool = False, delay_ms: int = 250) -> Dict[str, Any]:
        """Cleanly terminate LibreOffice via Desktop.terminate().

        Use this for hot-reload of the extension instead of `pkill -9`. A clean
        terminate writes the registry/clipboard properly and skips the
        Document-Recovery dialog on next launch.

        Implementation note: terminate() runs on a delayed background thread so
        the HTTP response can be flushed first — otherwise the LO process exits
        before the client gets the reply. force=True clears every doc's Modified
        flag so terminate() doesn't bail (DESTRUCTIVE — discards unsaved edits).
        """
        try:
            if force:
                try:
                    comps = self.desktop.getComponents()
                    it = comps.createEnumeration()
                    while it.hasMoreElements():
                        c = it.nextElement()
                        try:
                            if hasattr(c, "setModified"):
                                c.setModified(False)
                        except Exception:
                            pass
                except Exception:
                    pass
            import threading
            def _terminate():
                try:
                    self.desktop.terminate()
                except Exception:
                    pass
            t = threading.Timer(max(0, int(delay_ms)) / 1000.0, _terminate)
            t.daemon = True
            t.start()
            return {"success": True, "scheduled_in_ms": int(delay_ms), "force": force}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def show_window(self) -> Dict[str, Any]:
        """Make the active document's window visible.

        Use this AFTER finishing a batch of writes on a doc that was created
        with visible=False, to avoid the macOS SolarMutex contention that
        happens when a worker thread mutates a visible doc many times in a
        burst (visible doc → main-thread layout reflow → SolarMutex
        contention → eventual deadlock).
        """
        try:
            doc = self.get_active_document()
            if not doc:
                return {"success": False, "error": "No active document"}
            ctrl = doc.getCurrentController()
            if ctrl is None:
                return {"success": False, "error": "No controller"}
            frame = ctrl.getFrame()
            if frame is None:
                return {"success": False, "error": "No frame"}
            win = frame.getContainerWindow()
            if win is None:
                return {"success": False, "error": "No container window"}
            win.setVisible(True)
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def select_range(self, start: int, end: int) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        if end < start:
            return {"success": False, "error": "end must be >= start"}
        try:
            text = doc.getText()
            cursor = text.createTextCursor()
            cursor.gotoStart(False)
            cursor.goRight(int(start), False)
            cursor.goRight(int(end - start), True)
            doc.getCurrentController().select(cursor)
            return {"success": True, "start": start, "end": end}
        except Exception as e:
            return {"success": False, "error": str(e)}

    # ---------------------------------------------------------------------

    def _get_document_type(self, doc: Any) -> str:
        """Determine document type"""
        try:
            if doc.supportsService("com.sun.star.text.TextDocument"):
                return "writer"
            if doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                return "calc"
            if doc.supportsService("com.sun.star.presentation.PresentationDocument"):
                return "impress"
            if doc.supportsService("com.sun.star.drawing.DrawingDocument"):
                return "draw"
        except Exception:
            pass
        return "unknown"
    
    def _has_selection(self, doc: Any) -> bool:
        """Check if document has selected content"""
        try:
            if hasattr(doc, 'getCurrentController'):
                controller = doc.getCurrentController()
                if hasattr(controller, 'getSelection'):
                    selection = controller.getSelection()
                    return selection.getCount() > 0
        except:
            pass
        return False
