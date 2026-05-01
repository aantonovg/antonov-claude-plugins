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

            # Per-character kerning (extra horizontal spacing in 1/100 mm).
            # Word imports often store small kerning on space characters between
            # bold names — without this, justify-wrap differs.
            # Wrap in try/except: setting CharKerning on certain ranges has been
            # seen to deadlock under view_locked batches (LO tries to relayout).
            if "kerning" in formatting and formatting["kerning"] is not None:
                try: text_range.CharKerning = int(formatting["kerning"])
                except Exception as kex: logger.warning(f"CharKerning set failed: {kex}")

            if "scale_width" in formatting and formatting["scale_width"] is not None:
                try: text_range.CharScaleWidth = int(formatting["scale_width"])
                except Exception as sex: logger.warning(f"CharScaleWidth set failed: {sex}")

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
        # TabAlign enum: LEFT=0, CENTER=1, RIGHT=2, DECIMAL=3, DEFAULT=4
        align_name_map = {"LEFT": "left", "CENTER": "center", "RIGHT": "right",
                          "DECIMAL": "decimal", "DEFAULT": "default"}
        align_int_map = {0: "left", 1: "center", 2: "right", 3: "decimal", 4: "default"}
        out = []
        if not stops:
            return out
        for t in stops:
            entry = {}
            try:
                entry["position_mm"] = t.Position / 100.0
            except Exception:
                continue  # no position → skip
            # alignment can be enum (newer LO) or int (older) — try both
            try:
                a = t.Alignment
                a_name = getattr(a, "value", None)
                if isinstance(a_name, str):
                    entry["alignment"] = align_name_map.get(a_name, a_name.lower())
                else:
                    try: entry["alignment"] = align_int_map.get(int(a), "left")
                    except Exception: entry["alignment"] = "left"
            except Exception:
                entry["alignment"] = "left"
            try:
                entry["fill_char"] = chr(t.FillChar) if t.FillChar else " "
            except Exception:
                entry["fill_char"] = " "
            try:
                entry["decimal_char"] = chr(t.DecimalChar) if t.DecimalChar else "."
            except Exception:
                entry["decimal_char"] = "."
            out.append(entry)
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

    def set_paragraph_alignment(self, alignment, start=None, end=None) -> Dict[str, Any]:
        """alignment: string ('left'|'center'|'right'|'justify'|'block'|'stretch') or
        raw int matching com.sun.star.style.ParagraphAdjust enum:
        LEFT=0, RIGHT=1, BLOCK=2, CENTER=3, STRETCH=4, BLOCK_LINE=5.
        Word imports use STRETCH (4) or BLOCK_LINE (5) for headings/dates that
        appear "centered" via stretched whitespace — preserve via raw int."""
        doc, err = self._require_writer()
        if err:
            return err
        mapping = {"left": 0, "right": 1, "justify": 2, "block": 2, "center": 3,
                   "stretch": 4, "block_line": 5}
        if isinstance(alignment, int):
            val = alignment
        elif isinstance(alignment, str) and alignment.isdigit():
            val = int(alignment)
        else:
            val = mapping.get(str(alignment).lower())
        if val is None or not (0 <= val <= 5):
            return {"success": False,
                    "error": "Unknown alignment, use: left|center|right|justify|stretch|block_line or int 0-5"}
        try:
            rng = self._resolve_range(doc, start, end)
            rng.ParaAdjust = val
            return {"success": True, "alignment": val}
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

    _BREAK_NAMES = ["NONE", "COLUMN_BEFORE", "COLUMN_AFTER", "COLUMN_BOTH",
                    "PAGE_BEFORE", "PAGE_AFTER", "PAGE_BOTH"]

    def set_paragraph_breaks(self, break_type=None, page_desc_name=None,
                             page_number_offset=None,
                             start=None, end=None) -> Dict[str, Any]:
        """Set BreakType / PageDescName / PageNumberOffset on paragraphs in [start, end).

        break_type: int 0..6 or string name from com.sun.star.style.BreakType
            (NONE=0, COLUMN_BEFORE=1, COLUMN_AFTER=2, COLUMN_BOTH=3,
             PAGE_BEFORE=4, PAGE_AFTER=5, PAGE_BOTH=6). PAGE_BEFORE forces
            a page break before the paragraph — needed to replicate the
            visual page layout of Word-imported documents.
        page_desc_name: name of a page-style assigned via PageDescName.
            Only effective on paragraphs that also carry a PAGE_* break.
        page_number_offset: int — restart numbering at this value at the
            page introduced by the break.
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            rng = self._resolve_range(doc, start, end)
            applied = {}
            if break_type is not None:
                if isinstance(break_type, str) and not break_type.lstrip("-").isdigit():
                    bt_name = break_type.upper()
                    if bt_name not in self._BREAK_NAMES:
                        return {"success": False, "error": f"unknown break_type {break_type!r}"}
                else:
                    bt_int = int(break_type)
                    if not 0 <= bt_int <= 6:
                        return {"success": False, "error": "break_type must be 0..6"}
                    bt_name = self._BREAK_NAMES[bt_int]
                rng.BreakType = uno.Enum("com.sun.star.style.BreakType", bt_name)
                applied["break_type"] = bt_name
            if page_desc_name is not None:
                rng.PageDescName = str(page_desc_name) if page_desc_name else ""
                applied["page_desc_name"] = page_desc_name
            if page_number_offset is not None:
                rng.PageNumberOffset = int(page_number_offset)
                applied["page_number_offset"] = int(page_number_offset)
            return {"success": True, "applied": applied}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_text_flow(self, widows=None, orphans=None,
                                keep_together=None, split_paragraph=None,
                                keep_with_next=None,
                                start=None, end=None) -> Dict[str, Any]:
        """Set text-flow properties on paragraphs in [start, end). These
        govern HOW the paragraph breaks across pages — the difference between
        "last word falls onto a near-empty next page" and "all 5 lines stay
        together". Word→ODT often sets these per paragraph; without
        replicating, target page-layout drifts from source.

        widows / orphans: int — minimum number of lines to keep on the
            new / old page when paragraph breaks across them.
        keep_together: bool — paragraph never splits across pages.
        split_paragraph: bool — paragraph IS allowed to split (False = same
            as keep_together but exposed as separate UNO field).
        keep_with_next: bool — paragraph stays glued to the next one
            (no break between them).
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            rng = self._resolve_range(doc, start, end)
            applied = {}
            if widows is not None:
                rng.ParaWidows = int(widows)
                applied["widows"] = int(widows)
            if orphans is not None:
                rng.ParaOrphans = int(orphans)
                applied["orphans"] = int(orphans)
            if keep_together is not None:
                rng.ParaKeepTogether = bool(keep_together)
                applied["keep_together"] = bool(keep_together)
            if split_paragraph is not None:
                rng.ParaSplit = bool(split_paragraph)
                applied["split_paragraph"] = bool(split_paragraph)
            if keep_with_next is not None:
                # ParaKeepWithNext is the modern name; older builds expose
                # KeepWithNext directly. Try both.
                for prop in ("ParaKeepWithNext", "KeepWithNext"):
                    try:
                        setattr(rng, prop, bool(keep_with_next))
                        break
                    except Exception:
                        continue
                applied["keep_with_next"] = bool(keep_with_next)
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

    def list_text_frames(self, source_path: str = None, doc_title: str = None) -> Dict[str, Any]:
        """List all TextFrames (free-floating boxes) in a doc with their text,
        size, position, anchor type. PageNumber fields visible at page bottoms
        are often inside TextFrames anchored to body paragraphs, NOT in the
        page-style's footer slot."""
        try:
            doc = None
            if source_path:
                doc = self._find_open_doc(source_path)
            if doc is None and doc_title:
                comps = self.desktop.getComponents()
                it = comps.createEnumeration()
                while it.hasMoreElements():
                    c = it.nextElement()
                    try: t = c.getTitle()
                    except Exception: t = ""
                    if t == doc_title:
                        doc = c; break
            if doc is None:
                doc, err = self._require_writer()
                if err: return err
            frames = []
            try:
                tf = doc.getTextFrames()
                for i in range(tf.Count):
                    f = tf.getByIndex(i)
                    entry = {"name": getattr(f, "Name", "")}
                    try: entry["text"] = (f.getString() or "")[:100]
                    except Exception: pass
                    try:
                        sz = f.Size
                        entry["size_mm"] = {"w": sz.Width/100.0, "h": sz.Height/100.0}
                    except Exception: pass
                    try:
                        pos = f.Position
                        entry["pos_mm"] = {"x": pos.X/100.0, "y": pos.Y/100.0}
                    except Exception: pass
                    try:
                        a = f.AnchorType
                        entry["anchor_type"] = getattr(a, "value", str(a))
                    except Exception: pass
                    try:
                        anchor = f.Anchor
                        if anchor is not None:
                            entry["anchor_text"] = (anchor.getString() or "")[:60]
                    except Exception: pass
                    # Find anchor paragraph index. anchor.getString() is empty
                    # for all 11 page-number frames so string equality fails.
                    # compareRegionStarts and gotoPreviousParagraph also fail
                    # because frame.Anchor lives in a different Text container.
                    # Last resort: pyuno proxy identity — if frame.Anchor IS
                    # one of the body paragraphs (XTextContent), `==` between
                    # them will be True for the same wrapped UNO object.
                    try:
                        anchor = f.Anchor
                        if anchor is not None:
                            body = doc.getText()
                            pe = body.createEnumeration()
                            idx = 0
                            matched = None
                            while pe.hasMoreElements():
                                p = pe.nextElement()
                                if p.supportsService("com.sun.star.text.Paragraph"):
                                    try:
                                        if p == anchor:
                                            matched = idx
                                            break
                                    except Exception: pass
                                    idx += 1
                            if matched is not None:
                                entry["anchor_para_index"] = matched
                            else:
                                entry["anchor_para_error"] = "anchor != any body paragraph"
                    except Exception as e:
                        entry["anchor_para_error"] = str(e)
                    # Vertical/horizontal positioning
                    for prop in ("HoriOrient", "VertOrient", "HoriOrientPosition",
                                 "VertOrientPosition", "HoriOrientRelation",
                                 "VertOrientRelation", "RelativeWidth", "RelativeHeight"):
                        try:
                            v = f.getPropertyValue(prop)
                            if hasattr(v, "value"): v = v.value
                            if isinstance(v, (int, float, str, bool)):
                                entry[prop] = v
                        except Exception: pass
                    # Border/transparency
                    try: entry["BackTransparent"] = bool(f.BackTransparent)
                    except Exception: pass
                    try:
                        info = f.getPropertySetInfo()
                        for bp in ("LeftBorder", "RightBorder", "TopBorder", "BottomBorder"):
                            if info.hasPropertyByName(bp):
                                try:
                                    bord = f.getPropertyValue(bp)
                                    entry[bp + "_width"] = int(getattr(bord, "OuterLineWidth", 0) or 0)
                                except Exception: pass
                    except Exception: pass
                    # walk inner portions to detect TextFields
                    fields = []
                    try:
                        pe = f.createEnumeration()
                        while pe.hasMoreElements():
                            p = pe.nextElement()
                            qe = p.createEnumeration()
                            while qe.hasMoreElements():
                                q = qe.nextElement()
                                if getattr(q, "TextPortionType", "") == "TextField":
                                    try:
                                        fld = q.TextField
                                        svc = next((s for s in (fld.SupportedServiceNames or [])
                                                    if s.startswith("com.sun.star.text.TextField.")
                                                    and s != "com.sun.star.text.TextField"), "?")
                                        fields.append({"service": svc.split(".")[-1],
                                                       "present": fld.getPresentation(False) if hasattr(fld, "getPresentation") else None})
                                    except Exception: pass
                    except Exception: pass
                    if fields: entry["fields"] = fields
                    frames.append(entry)
            except Exception as ex:
                return {"success": False, "error": str(ex)}
            return {"success": True, "frames": frames, "count": len(frames)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def list_text_fields(self, source_path: str = None, doc_title: str = None) -> Dict[str, Any]:
        """List all TextFields in a doc with their service name, presentation,
        and anchor text/paragraph. Reveals PageNumber / Date / etc. fields
        anywhere in the document (body, frames, headers, footers)."""
        try:
            doc = None
            if source_path:
                doc = self._find_open_doc(source_path)
            if doc is None and doc_title:
                comps = self.desktop.getComponents()
                it = comps.createEnumeration()
                while it.hasMoreElements():
                    c = it.nextElement()
                    try: t = c.getTitle()
                    except Exception: t = ""
                    if t == doc_title:
                        doc = c; break
            if doc is None:
                doc, err = self._require_writer()
                if err: return err
            out = []
            try:
                e = doc.getTextFields().createEnumeration()
                while e.hasMoreElements():
                    f = e.nextElement()
                    services = list(getattr(f, "SupportedServiceNames", []) or [])
                    svc = next((s for s in services if s.startswith("com.sun.star.text.TextField.")
                                and s != "com.sun.star.text.TextField"), services[0] if services else "?")
                    entry = {"service": svc}
                    try: entry["present"] = f.getPresentation(False)
                    except Exception: pass
                    # anchor text snippet
                    try:
                        a = f.getAnchor()
                        if a is not None:
                            try: entry["anchor_text"] = (a.getString() or "")[:60]
                            except Exception: pass
                    except Exception: pass
                    out.append(entry)
            except Exception as ex:
                return {"success": False, "error": str(ex)}
            return {"success": True, "fields": out, "count": len(out)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def dump_char_style(self, style_name: str, source_path: str = None,
                        doc_title: str = None) -> Dict[str, Any]:
        """Read a character style's font/size/weight/posture/etc. from any open
        doc. Compares cloned char styles between source and target."""
        try:
            doc = None
            if source_path:
                doc = self._find_open_doc(source_path)
            if doc is None and doc_title:
                comps = self.desktop.getComponents()
                it = comps.createEnumeration()
                while it.hasMoreElements():
                    c = it.nextElement()
                    try: t = c.getTitle()
                    except Exception: t = ""
                    if t == doc_title:
                        doc = c; break
            if doc is None:
                doc, err = self._require_writer()
                if err: return err
            fam = doc.getStyleFamilies().getByName("CharacterStyles")
            if not fam.hasByName(style_name):
                return {"success": False, "error": f"char style {style_name!r} not found"}
            st = fam.getByName(style_name)
            out = {}
            for prop in ("CharFontName", "CharHeight", "CharWeight",
                         "CharPosture", "CharColor", "CharUnderline",
                         "CharStrikeout", "CharKerning", "CharScaleWidth",
                         "CharBackColor", "ParentStyle", "DisplayName"):
                try:
                    v = st.getPropertyValue(prop)
                    if hasattr(v, "value"):
                        v = v.value
                    if isinstance(v, (bool, int, float, str)):
                        out[prop] = v
                    else:
                        out[prop] = str(v)
                except Exception: pass
            return {"success": True, "style": style_name, "props": out,
                    "doc_title": doc.getTitle() if hasattr(doc, "getTitle") else None}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def dump_doc_paragraph(self, source_path: str = None, doc_title: str = None,
                           paragraph_index: int = 0) -> Dict[str, Any]:
        """Return paragraph + rule level props as JSON-safe primitives. No
        struct walking — pure int/float/str only. Picks doc by source_path
        (file URL match) or doc_title (e.g. 'Untitled 1'). Falls back to active."""
        try:
            doc = None
            if source_path:
                doc = self._find_open_doc(source_path)
            if doc is None and doc_title:
                comps = self.desktop.getComponents()
                it = comps.createEnumeration()
                while it.hasMoreElements():
                    c = it.nextElement()
                    try: t = c.getTitle()
                    except Exception: t = ""
                    if t == doc_title:
                        doc = c; break
            if doc is None:
                doc, err = self._require_writer()
                if err: return err

            text = doc.getText()
            pe = text.createEnumeration()
            i = 0; target = None
            while pe.hasMoreElements():
                p = pe.nextElement()
                if not p.supportsService("com.sun.star.text.Paragraph"):
                    continue
                if i == paragraph_index:
                    target = p; break
                i += 1
            if target is None:
                return {"success": False, "error": f"paragraph {paragraph_index} not found"}

            def _i(x):
                try: return int(x)
                except Exception:
                    try: return int(getattr(x, "value", 0))
                    except Exception: return None
            def _s(x):
                try: return str(x) if x is not None else None
                except Exception: return None

            out = {
                "index": paragraph_index,
                "text": (target.getString() or "")[:120],
                "ParaStyleName": _s(target.ParaStyleName),
                "ParaAdjust": _i(target.ParaAdjust),
                "ParaLeftMargin": _i(target.ParaLeftMargin),
                "ParaRightMargin": _i(target.ParaRightMargin),
                "ParaFirstLineIndent": _i(target.ParaFirstLineIndent),
                "ParaTopMargin": _i(getattr(target, "ParaTopMargin", 0)),
                "ParaBottomMargin": _i(getattr(target, "ParaBottomMargin", 0)),
                "NumberingLevel": _i(getattr(target, "NumberingLevel", 0)),
                "NumberingIsNumber": bool(getattr(target, "NumberingIsNumber", False)),
                "ListLabelString": _s(getattr(target, "ListLabelString", "")),
            }
            try:
                ls = target.ParaLineSpacing
                out["ParaLineSpacing"] = {"mode": _i(ls.Mode), "height": _i(ls.Height)}
            except Exception: pass
            try:
                tabs = []
                for ts in (target.ParaTabStops or []):
                    try:
                        tabs.append({"position": _i(ts.Position),
                                     "alignment": _i(getattr(ts, "Alignment", 0))})
                    except Exception: pass
                out["ParaTabStops"] = tabs
            except Exception: pass
            # runs with full char props
            try:
                runs = []
                qe = target.createEnumeration()
                while qe.hasMoreElements():
                    q = qe.nextElement()
                    ptype = getattr(q, "TextPortionType", "?")
                    s = q.getString() or ""
                    if not s and ptype == "Text": continue
                    r = {"type": ptype, "text": s[:80]}
                    for cp in ("CharFontName", "CharHeight", "CharWeight",
                               "CharPosture", "CharUnderline", "CharStyleName",
                               "CharKerning", "CharScaleWidth", "CharWordMode",
                               "CharLocale", "CharFontNameAsian", "CharHeightAsian",
                               "CharWeightAsian", "CharFontNameComplex", "CharHeightComplex",
                               "CharWeightComplex", "CharNoHyphenation"):
                        try:
                            v = q.getPropertyValue(cp)
                            if hasattr(v, "value"):
                                v = v.value
                            if isinstance(v, (bool, int, float, str)):
                                r[cp] = v
                            else:
                                r[cp] = str(v)
                        except Exception: pass
                    runs.append(r)
                out["runs"] = runs
            except Exception: pass
            try:
                rules = target.getPropertyValue("NumberingRules")
                if rules is not None:
                    out["NumberingRules_Name"] = _s(getattr(rules, "Name", None))
                    lvl = _i(target.NumberingLevel) or 0
                    try:
                        props = rules.getByIndex(lvl)
                        rule_lvl = {}
                        for pv in props:
                            try:
                                v = pv.Value
                                if isinstance(v, (bool, int, float, str)):
                                    rule_lvl[pv.Name] = v
                                elif hasattr(v, "value") and isinstance(v.value, (bool, int, float, str)):
                                    rule_lvl[pv.Name] = v.value
                                else:
                                    rule_lvl[pv.Name] = type(v).__name__
                            except Exception: pass
                        out["rule_level_props"] = rule_lvl
                    except Exception as ex:
                        out["rule_level_err"] = str(ex)
            except Exception: pass
            return {"success": True, "paragraph": out,
                    "doc_title": _s(doc.getTitle() if hasattr(doc, "getTitle") else None)}
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

            src_rules_obj = src_fam.getByName(rule_name).NumberingRules
            tgt_style.NumberingRules = src_rules_obj

            # Auto-clone any character styles referenced by level CharStyleName.
            # Without this, target gets a stub char style with default font (e.g.
            # Liberation Serif 12pt) instead of source's actual style (e.g. Calibri
            # 11pt) — labels render with wrong width and tab-gap differs.
            cloned_chars = []
            try:
                src_char_fam = src_doc.getStyleFamilies().getByName("CharacterStyles")
                tgt_char_fam = doc.getStyleFamilies().getByName("CharacterStyles")
                seen = set()
                for lvl in range(src_rules_obj.Count):
                    try:
                        props = src_rules_obj.getByIndex(lvl)
                    except Exception:
                        continue
                    char_name = ""
                    for pv in props:
                        if pv.Name == "CharStyleName":
                            char_name = pv.Value or ""
                            break
                    if not char_name or char_name in seen:
                        continue
                    seen.add(char_name)
                    if not src_char_fam.hasByName(char_name):
                        continue
                    src_cs = src_char_fam.getByName(char_name)
                    if tgt_char_fam.hasByName(char_name):
                        tgt_cs = tgt_char_fam.getByName(char_name)
                    else:
                        tgt_cs = doc.createInstance("com.sun.star.style.CharacterStyle")
                        tgt_char_fam.insertByName(char_name, tgt_cs)
                    # copy a focused set of char props
                    for cprop in ("CharFontName", "CharHeight", "CharWeight",
                                  "CharPosture", "CharColor", "CharUnderline",
                                  "CharStrikeout", "CharKerning", "CharScaleWidth",
                                  "CharBackColor", "CharFontNameAsian", "CharHeightAsian",
                                  "CharWeightAsian", "CharPostureAsian",
                                  "CharFontNameComplex", "CharHeightComplex",
                                  "CharWeightComplex", "CharPostureComplex"):
                        try:
                            if not src_cs.getPropertySetInfo().hasPropertyByName(cprop): continue
                            if not tgt_cs.getPropertySetInfo().hasPropertyByName(cprop): continue
                            tgt_cs.setPropertyValue(cprop, src_cs.getPropertyValue(cprop))
                        except Exception:
                            pass
                    cloned_chars.append(char_name)
            except Exception as ex:
                cloned_chars.append(f"<err: {ex}>")

            return {"success": True, "rule_name": tgt_name, "created": created,
                    "source_rule": rule_name, "source_path": source_path,
                    "cloned_char_styles": cloned_chars}
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
                # KeepWithNext on a paragraph style (e.g. Heading 1) glues
                # heading to the following block (paragraph or table). Without
                # cloning this prop, target's heading "release" the table
                # below — both jump to the next page as one big chunk
                # instead of staying together with body-text continuing
                # naturally on the current page.
                "ParaKeepWithNext", "KeepWithNext",
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

    def _clone_xtext(self, doc, src_xtext, tgt_xtext) -> Dict[str, int]:
        """Replace tgt_xtext content with src_xtext content, preserving paragraph
        structure, character formatting, and TextFields (PageNumber, Date, etc.).

        setString() loses fields — this walks portions and recreates each one.
        """
        # Clear target. setString("") is idempotent and cheap.
        try:
            tgt_xtext.setString("")
        except Exception:
            pass

        cursor = tgt_xtext.createTextCursor()
        # Char props worth carrying onto inserted text. Para props on portion=0
        # captured when we set them on the cursor before each paragraph.
        CHAR_PROPS = (
            "CharFontName", "CharHeight", "CharWeight", "CharPosture",
            "CharUnderline", "CharStrikeout", "CharColor", "CharBackColor",
            "CharEscapement", "CharEscapementHeight",
            "CharFontNameAsian", "CharHeightAsian", "CharWeightAsian", "CharPostureAsian",
            "CharFontNameComplex", "CharHeightComplex", "CharWeightComplex",
            "CharStyleName", "CharKerning", "CharWordMode",
        )
        PARA_PROPS = (
            "ParaAdjust", "ParaStyleName",
            "ParaLeftMargin", "ParaRightMargin", "ParaFirstLineIndent",
            "ParaTopMargin", "ParaBottomMargin",
            "ParaLineSpacing", "ParaTabStops",
        )

        def _copy_props(src_obj, dst_setter, names):
            """Copy each prop one-by-one; skip on error so partial read works."""
            try:
                info = src_obj.getPropertySetInfo()
            except Exception:
                info = None
            for n in names:
                try:
                    if info is not None and not info.hasPropertyByName(n):
                        continue
                    v = src_obj.getPropertyValue(n)
                    dst_setter(n, v)
                except Exception:
                    pass

        copied_text = 0
        copied_fields = 0
        first_para = True
        try:
            para_enum = src_xtext.createEnumeration()
        except Exception as ex:
            return {"copied_text": 0, "copied_fields": 0, "error": str(ex)}

        while para_enum.hasMoreElements():
            src_para = para_enum.nextElement()
            if not first_para:
                # PARAGRAPH_BREAK = 0
                try:
                    tgt_xtext.insertControlCharacter(cursor, 0, False)
                except Exception:
                    pass
            first_para = False

            # Apply paragraph-level props on the cursor before inserting text.
            _copy_props(src_para, lambda n, v: cursor.setPropertyValue(n, v), PARA_PROPS)

            try:
                portion_enum = src_para.createEnumeration()
            except Exception:
                continue

            while portion_enum.hasMoreElements():
                portion = portion_enum.nextElement()
                try:
                    ptype = portion.TextPortionType
                except Exception:
                    ptype = "Text"

                if ptype == "Text":
                    txt = ""
                    try:
                        txt = portion.getString() or ""
                    except Exception:
                        txt = ""
                    if not txt:
                        continue
                    _copy_props(portion, lambda n, v: cursor.setPropertyValue(n, v), CHAR_PROPS)
                    try:
                        tgt_xtext.insertString(cursor, txt, False)
                        copied_text += len(txt)
                    except Exception:
                        pass
                elif ptype == "TextField":
                    try:
                        src_fld = portion.TextField
                    except Exception:
                        src_fld = None
                    if src_fld is None:
                        try:
                            tgt_xtext.insertString(cursor, portion.getString() or "", False)
                        except Exception:
                            pass
                        continue
                    try:
                        services = list(src_fld.SupportedServiceNames or [])
                    except Exception:
                        services = []
                    fld_service = None
                    for s in services:
                        if (s.startswith("com.sun.star.text.TextField.")
                                and s != "com.sun.star.text.TextField"):
                            fld_service = s
                            break
                    if not fld_service:
                        try:
                            tgt_xtext.insertString(cursor, portion.getString() or "", False)
                        except Exception:
                            pass
                        continue
                    try:
                        new_fld = doc.createInstance(fld_service)
                        # copy field-specific props (NumberingType, SubType, IsFixed, ...)
                        try:
                            for prop in src_fld.getPropertySetInfo().getProperties():
                                try:
                                    new_fld.setPropertyValue(
                                        prop.Name,
                                        src_fld.getPropertyValue(prop.Name))
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        # apply char props at insertion point
                        _copy_props(portion, lambda n, v: cursor.setPropertyValue(n, v), CHAR_PROPS)
                        tgt_xtext.insertTextContent(cursor, new_fld, False)
                        copied_fields += 1
                    except Exception:
                        try:
                            tgt_xtext.insertString(cursor, portion.getString() or "", False)
                        except Exception:
                            pass
                # other portion types (Bookmark/Footnote/Frame) — skip; rare in headers/footers

        return {"copied_text": copied_text, "copied_fields": copied_fields}

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
                # Word→ODT imports use per-section page-styles named MP0..MPN
                # to encode different headers/footers across sections. When
                # replicating such a doc we need MPx styles in target — create
                # them on demand instead of falling back to Default.
                try:
                    new_style = doc.createInstance("com.sun.star.style.PageStyle")
                    tgt_fam.insertByName(target_style, new_style)
                except Exception as ce:
                    # Locale fallback only if creation fails (older LO, etc.)
                    for cand in ("Default Page Style", "Standard", "Default Style"):
                        if tgt_fam.hasByName(cand):
                            target_style = cand; break
                    if not tgt_fam.hasByName(target_style):
                        return {"success": False,
                                "error": f"could not create or fallback target page style: {ce}",
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

            # Header/Footer text bodies are XText objects — clone full structure
            # (paragraphs, char formatting, TextFields like PageNumber).
            header_stats = None
            footer_stats = None
            try:
                if getattr(src, "HeaderIsOn", False) and getattr(tgt, "HeaderIsOn", False):
                    header_stats = self._clone_xtext(doc, src.HeaderText, tgt.HeaderText)
            except Exception as ex:
                failed.append({"prop": "HeaderText", "error": str(ex)})
            try:
                if getattr(src, "FooterIsOn", False) and getattr(tgt, "FooterIsOn", False):
                    footer_stats = self._clone_xtext(doc, src.FooterText, tgt.FooterText)
            except Exception as ex:
                failed.append({"prop": "FooterText", "error": str(ex)})

            return {"success": True, "source_style": source_style,
                    "target_style": target_style,
                    "header_enabled": bool(getattr(tgt, "HeaderIsOn", False)),
                    "footer_enabled": bool(getattr(tgt, "FooterIsOn", False)),
                    "header_stats": header_stats,
                    "footer_stats": footer_stats,
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
                        # ParagraphAdjust: LEFT=0,RIGHT=1,BLOCK=2,CENTER=3,STRETCH=4,BLOCK_LINE=5
                        # Expose raw int (paragraph_adjust) for fidelity + readable name (alignment).
                        adj = para.ParaAdjust
                        adj_int = int(getattr(adj, "value", adj)) if not isinstance(adj, int) else adj
                        try: adj_int = int(adj)
                        except Exception:
                            try: adj_int = int(getattr(adj, "value", 0))
                            except Exception: adj_int = 0
                        entry["paragraph_adjust"] = adj_int
                        entry["alignment"] = {0:"left",1:"right",2:"justify",3:"center",4:"stretch",5:"block_line"}.get(adj_int, str(adj_int))
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
                        # Paragraph-level CharHeight (font size). Important
                        # for EMPTY paragraphs: without runs, font size info
                        # gets lost — but UNO surfaces it on the paragraph
                        # itself, derived from style or override. Word→ODT
                        # often emits 1pt/2pt empty paragraphs to compress
                        # spacing; without replicating, target's 12pt empties
                        # take a full line each.
                        try: entry["char_height"] = float(getattr(para, "CharHeight", 0) or 0)
                        except Exception: entry["char_height"] = None
                        # Text-flow — Word→ODT часто кодирует "не разрывать
                        # абзац" / "не отрывать от следующего" на per-paragraph
                        # уровне. Без репликации последняя строка / слово
                        # "проваливается" на следующую страницу, ломая layout.
                        try: entry["widows"] = int(getattr(para, "ParaWidows", 0) or 0)
                        except Exception: entry["widows"] = 0
                        try: entry["orphans"] = int(getattr(para, "ParaOrphans", 0) or 0)
                        except Exception: entry["orphans"] = 0
                        try: entry["keep_together"] = bool(getattr(para, "ParaKeepTogether", False))
                        except Exception: entry["keep_together"] = False
                        try: entry["split_paragraph"] = bool(getattr(para, "ParaSplit", True))
                        except Exception: entry["split_paragraph"] = True
                        try:
                            kwn = getattr(para, "ParaKeepWithNext", None)
                            if kwn is None:
                                kwn = getattr(para, "KeepWithNext", False)
                            entry["keep_with_next"] = bool(kwn)
                        except Exception: entry["keep_with_next"] = False
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
                        # ParagraphAdjust: LEFT=0,RIGHT=1,BLOCK=2,CENTER=3,STRETCH=4,BLOCK_LINE=5
                        # Expose raw int (paragraph_adjust) for fidelity + readable name (alignment).
                        adj = para.ParaAdjust
                        adj_int = int(getattr(adj, "value", adj)) if not isinstance(adj, int) else adj
                        try: adj_int = int(adj)
                        except Exception:
                            try: adj_int = int(getattr(adj, "value", 0))
                            except Exception: adj_int = 0
                        entry["paragraph_adjust"] = adj_int
                        entry["alignment"] = {0:"left",1:"right",2:"justify",3:"center",4:"stretch",5:"block_line"}.get(adj_int, str(adj_int))
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
                        try: entry["char_height"] = float(getattr(para, "CharHeight", 0) or 0)
                        except Exception: entry["char_height"] = None
                        # Text-flow (см. комментарий в get_paragraphs)
                        try: entry["widows"] = int(getattr(para, "ParaWidows", 0) or 0)
                        except Exception: entry["widows"] = 0
                        try: entry["orphans"] = int(getattr(para, "ParaOrphans", 0) or 0)
                        except Exception: entry["orphans"] = 0
                        try: entry["keep_together"] = bool(getattr(para, "ParaKeepTogether", False))
                        except Exception: entry["keep_together"] = False
                        try: entry["split_paragraph"] = bool(getattr(para, "ParaSplit", True))
                        except Exception: entry["split_paragraph"] = True
                        try:
                            kwn = getattr(para, "ParaKeepWithNext", None)
                            if kwn is None:
                                kwn = getattr(para, "KeepWithNext", False)
                            entry["keep_with_next"] = bool(kwn)
                        except Exception: entry["keep_with_next"] = False
                        # BreakType / PageDescName тоже нужны для diff
                        try:
                            entry["page_desc_name"] = getattr(para, "PageDescName", "") or ""
                        except Exception:
                            entry["page_desc_name"] = ""
                        try:
                            bt = getattr(para, "BreakType", 0)
                            if hasattr(bt, "value"):
                                entry["break_type"] = int(bt.value)
                            else:
                                entry["break_type"] = int(bt) if isinstance(bt, (int, float)) else 0
                        except Exception:
                            entry["break_type"] = 0
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
                            # CharPosture is com.sun.star.awt.FontSlant enum. In Python UNO,
                            # `.value` returns the NAME string ("NONE","OBLIQUE","ITALIC",
                            # "DONTKNOW","REVERSE_OBLIQUE","REVERSE_ITALIC"), not the int.
                            # OBLIQUE is what Word imports use for italic.
                            cp_name = ""
                            try:
                                cp = portion.CharPosture
                                cp_name = getattr(cp, "value", None) or str(cp) or ""
                            except Exception:
                                cp_name = ""
                            run["italic"] = cp_name in ("OBLIQUE", "ITALIC",
                                                          "REVERSE_OBLIQUE", "REVERSE_ITALIC")
                            run["char_posture"] = cp_name
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
                            # Per-portion kerning (1/100 mm) and scale width (%).
                            # Source docs sometimes set per-space kerning to push justify
                            # rendering wider — without this, target wraps differently.
                            try:
                                k = int(getattr(portion, "CharKerning", 0) or 0)
                                if k != 0: run["kerning"] = k
                            except Exception: pass
                            try:
                                sw = int(getattr(portion, "CharScaleWidth", 100) or 100)
                                if sw != 100: run["scale_width"] = sw
                            except Exception: pass
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

    def list_body_elements(self, start: int = 0, count: int = None,
                           preview_chars: int = 60) -> Dict[str, Any]:
        """Iterate the document body and return ordered paragraphs **and**
        tables (TextTable) with positional metadata.

        Unlike _iter_paragraphs (which silently skips tables), this exposes
        the absolute order needed to reconstruct mixed paragraph/table flows.
        Each paragraph entry mirrors the index produced by get_paragraphs.
        Tables include their name, dimensions, and the index of the
        paragraph immediately before them — that's what the batch reproducer
        uses to insert a table at the right anchor in the rebuilt body.

        anchor_offset uses the same offset accounting as _iter_paragraphs
        (offset += len(getString()) + 1 per element) — only meaningful
        relative to other elements in this same enumeration.
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            text = doc.getText()
            enum = text.createEnumeration()
            offset = 0
            para_idx = 0
            elements = []
            last_para_idx = -1
            seq = 0
            while enum.hasMoreElements():
                elem = enum.nextElement()
                if elem.supportsService("com.sun.star.text.Paragraph"):
                    s = elem.getString()
                    if seq >= start and (count is None or len(elements) < count):
                        elements.append({
                            "kind": "paragraph",
                            "index": para_idx,
                            "start": offset,
                            "end": offset + len(s),
                            "preview": s[:preview_chars] + ("…" if len(s) > preview_chars else ""),
                        })
                    last_para_idx = para_idx
                    offset += len(s) + 1
                    para_idx += 1
                    seq += 1
                elif elem.supportsService("com.sun.star.text.TextTable"):
                    if seq >= start and (count is None or len(elements) < count):
                        try:
                            rows = elem.getRows().getCount()
                            cols = elem.getColumns().getCount()
                        except Exception:
                            rows, cols = 0, 0
                        try:
                            tname = elem.getName()
                        except Exception:
                            tname = ""
                        elements.append({
                            "kind": "table",
                            "name": tname,
                            "after_paragraph_index": last_para_idx,
                            "anchor_offset": offset,
                            "rows": rows,
                            "columns": cols,
                        })
                    try:
                        s = elem.getString()
                        offset += len(s) + 1
                    except Exception:
                        offset += 1
                    seq += 1
                else:
                    # Unknown element — keep offset accounting consistent
                    try:
                        s = elem.getString()
                        offset += len(s) + 1
                    except Exception:
                        offset += 1
                    seq += 1
            return {"success": True, "elements": elements,
                    "returned": len(elements)}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_page_layout(self) -> Dict[str, Any]:
        """Map every body element (paragraph, table, table row, frame) to its
        actual page using ViewCursor.getPage(). Each element gets start_page
        and end_page; if they differ the element is split between pages
        (paragraph wrapping, row split with Split=True, table spanning pages).

        Returns:
        - page_count
        - elements: ordered list with start_page/end_page (and per-row split
          info for tables)
        - pages: inverse view — for each page, which elements touch it (with
          is_start/is_end markers)
        - frames: TextFrames mapped to their anchor page

        Caveat: requires layout to be computed. If the document is hidden,
        getPage() may return 0 until visibility is restored. Caller should
        ensure the doc is visible (or call show_window first). Each lookup
        does ctrl.select(range) which moves the visible selection — expect
        a brief flicker on a visible window.
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ctrl = doc.getCurrentController()
            vc = ctrl.getViewCursor()

            def _page_of(rng):
                if rng is None:
                    return 0
                try:
                    ctrl.select(rng)
                    p = vc.getPage()
                    return int(p) if p else 0
                except Exception:
                    return 0

            text = doc.getText()
            enum = text.createEnumeration()
            elements = []
            para_idx = 0
            offset = 0
            while enum.hasMoreElements():
                elem = enum.nextElement()
                if elem.supportsService("com.sun.star.text.Paragraph"):
                    s = elem.getString()
                    sp = _page_of(elem.getStart())
                    ep = _page_of(elem.getEnd()) if s else sp
                    entry = {
                        "kind": "paragraph",
                        "index": para_idx,
                        "start": offset,
                        "end": offset + len(s),
                        "preview": (s[:60] + ("…" if len(s) > 60 else "")),
                        "start_page": sp,
                        "end_page": ep,
                    }
                    if sp and ep and sp != ep:
                        entry["spans_pages"] = True
                    elements.append(entry)
                    offset += len(s) + 1
                    para_idx += 1
                elif elem.supportsService("com.sun.star.text.TextTable"):
                    try: tname = elem.getName()
                    except Exception: tname = ""
                    try:
                        n_rows = elem.getRows().getCount()
                        n_cols = elem.getColumns().getCount()
                    except Exception:
                        n_rows, n_cols = 0, 0
                    rows_info = []
                    t_sp = 0
                    t_ep = 0
                    for r in range(n_rows):
                        try:
                            first_cell = elem.getCellByPosition(0, r)
                            last_cell = elem.getCellByPosition(n_cols - 1, r) if n_cols else first_cell
                            rsp = _page_of(first_cell.getStart())
                            rep = _page_of(last_cell.getEnd())
                        except Exception:
                            rsp = 0; rep = 0
                        row_entry = {"row": r, "start_page": rsp, "end_page": rep}
                        if rsp and rep and rsp != rep:
                            row_entry["spans_pages"] = True
                        rows_info.append(row_entry)
                        if r == 0:
                            t_sp = rsp
                        t_ep = rep
                    entry = {
                        "kind": "table",
                        "name": tname,
                        "rows_count": n_rows,
                        "columns_count": n_cols,
                        "anchor_offset": offset,
                        "start_page": t_sp,
                        "end_page": t_ep,
                        "rows": rows_info,
                    }
                    if t_sp and t_ep and t_sp != t_ep:
                        entry["spans_pages"] = True
                    elements.append(entry)
                    try:
                        s = elem.getString(); offset += len(s) + 1
                    except Exception:
                        offset += 1
                else:
                    try:
                        s = elem.getString(); offset += len(s) + 1
                    except Exception:
                        offset += 1

            # Frames mapped to anchor page
            frames_by_page = []
            try:
                tf = doc.getTextFrames()
                for i in range(tf.Count):
                    f = tf.getByIndex(i)
                    try:
                        a = f.Anchor
                    except Exception:
                        a = None
                    fp = _page_of(a) if a is not None else 0
                    try: fname = f.Name
                    except Exception: fname = ""
                    frames_by_page.append({"name": fname, "page": fp})
            except Exception:
                pass

            # Determine page count
            page_count = 0
            try:
                page_count = int(doc.getPropertyValue("PageCount") or 0)
            except Exception:
                pass
            if page_count == 0:
                try:
                    vc.jumpToLastPage()
                    page_count = int(vc.getPage() or 0)
                except Exception:
                    pass
            if page_count == 0:
                for e in elements:
                    page_count = max(page_count, e.get("end_page") or 0,
                                     e.get("start_page") or 0)

            # Inverse view: pages[]
            pages = [{"page": n + 1, "contents": [], "frames": []}
                     for n in range(page_count)]
            for e in elements:
                sp = e.get("start_page") or 0
                ep = e.get("end_page") or 0
                if sp == 0 or ep == 0:
                    continue
                for p in range(sp, ep + 1):
                    if 1 <= p <= page_count:
                        ref = {
                            "kind": e["kind"],
                            "is_start": p == sp,
                            "is_end": p == ep,
                            "spans_pages": bool(e.get("spans_pages")),
                        }
                        if e["kind"] == "paragraph":
                            ref["index"] = e.get("index")
                            ref["preview"] = e.get("preview")
                        else:
                            ref["name"] = e.get("name")
                            ref["rows_count"] = e.get("rows_count")
                        pages[p - 1]["contents"].append(ref)
            for fr in frames_by_page:
                p = fr.get("page") or 0
                if 1 <= p <= page_count:
                    pages[p - 1]["frames"].append({"name": fr.get("name")})

            return {
                "success": True,
                "page_count": page_count,
                "elements": elements,
                "pages": pages,
                "frames": frames_by_page,
            }
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
            fmt = {
                "start": start,
                "end": start + length,
                "text": cursor.getString(),
                "font_name": cursor.CharFontName,
                "font_size": cursor.CharHeight,
                "bold": cursor.CharWeight >= 150,
                "italic": (getattr(cursor.CharPosture, "value", None) or str(cursor.CharPosture)) in ("OBLIQUE", "ITALIC", "REVERSE_OBLIQUE", "REVERSE_ITALIC"),
                "underline": cursor.CharUnderline != 0,
                "color": self._int_to_hex(cursor.CharColor),
                "background_color": self._int_to_hex(cursor.CharBackColor),
            }
            # Per-character spacing — Word imports often set CharKerning on
            # specific portions (e.g. expanded spacing on bold names) to widen
            # justify wraps. Without surfacing it here, agents can't diff or
            # replicate it.
            try: fmt["kerning"] = int(getattr(cursor, "CharKerning", 0) or 0)
            except Exception: pass
            try: fmt["scale_width"] = int(getattr(cursor, "CharScaleWidth", 100) or 100)
            except Exception: pass
            return {"success": True, "format": fmt}
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
            posture_raw = getattr(st, "CharPosture", 0)
            posture_name = getattr(posture_raw, "value", None) or str(posture_raw) or ""
            d = {
                "name": st.Name,
                "display_name": getattr(st, "DisplayName", st.Name),
                "parent": getattr(st, "ParentStyle", "") or "",
                "follow": getattr(st, "FollowStyle", "") or "",
                "font_name": getattr(st, "CharFontName", None),
                "font_size": getattr(st, "CharHeight", None),
                "bold": getattr(st, "CharWeight", 100) >= 150,
                "italic": posture_name in ("OBLIQUE", "ITALIC", "REVERSE_OBLIQUE", "REVERSE_ITALIC"),
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
            # Per-character spacing on the style level. Word imports may set
            # CharKerning > 0 on entire styles (e.g. expanded headings).
            try:
                k = int(getattr(st, "CharKerning", 0) or 0)
                if k != 0: d["kerning"] = k
            except Exception: pass
            try:
                sw = int(getattr(st, "CharScaleWidth", 100) or 100)
                if sw != 100: d["scale_width"] = sw
            except Exception: pass
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
            if "kerning" in props:
                _try(lambda v: setattr(st, "CharKerning", int(v)), "kerning", props["kerning"])
            if "scale_width" in props:
                _try(lambda v: setattr(st, "CharScaleWidth", int(v)), "scale_width", props["scale_width"])
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

    def _extract_para_with_runs(self, para) -> Dict[str, Any]:
        """Pull paragraph metadata + per-portion runs into a JSON-serialisable
        dict. Mirrors the schema of get_paragraphs_with_runs entries; reused
        by read_table_rich and any other context that needs full paragraph
        replication data."""
        entry = {"text": para.getString()}
        try:
            entry["style"] = para.ParaStyleName
            adj = para.ParaAdjust
            try:
                adj_int = int(adj)
            except Exception:
                try: adj_int = int(getattr(adj, "value", 0))
                except Exception: adj_int = 0
            entry["paragraph_adjust"] = adj_int
            entry["alignment"] = {0:"left",1:"right",2:"justify",3:"center",
                                  4:"stretch",5:"block_line"}.get(adj_int, str(adj_int))
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
            entry["line_spacing"] = {
                "mode": ["proportional","minimum","leading","fix"][ls.Mode]
                        if ls.Mode in (0,1,2,3) else ls.Mode,
                "value": ls.Height if ls.Mode == 0 else ls.Height / 100.0}
            try:
                entry["context_margin"] = bool(getattr(para, "ParaContextMargin", False))
            except Exception:
                entry["context_margin"] = False
        except Exception as fe:
            entry["format_error"] = str(fe)
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
                    cp_name = ""
                    try:
                        cp = portion.CharPosture
                        cp_name = getattr(cp, "value", None) or str(cp) or ""
                    except Exception:
                        cp_name = ""
                    run["italic"] = cp_name in ("OBLIQUE", "ITALIC",
                                                 "REVERSE_OBLIQUE", "REVERSE_ITALIC")
                    run["underline"] = portion.CharUnderline != 0
                    run["color"] = self._int_to_hex(portion.CharColor)
                    if getattr(portion, "CharBackColor", -1) not in (-1, 0xFFFFFFFF):
                        run["background_color"] = self._int_to_hex(portion.CharBackColor)
                    url = getattr(portion, "HyperLinkURL", "")
                    if url:
                        run["hyperlink"] = url
                    cstyle = getattr(portion, "CharStyleName", "")
                    if cstyle:
                        run["char_style"] = cstyle
                    try:
                        k = int(getattr(portion, "CharKerning", 0) or 0)
                        if k != 0: run["kerning"] = k
                    except Exception: pass
                    try:
                        sw = int(getattr(portion, "CharScaleWidth", 100) or 100)
                        if sw != 100: run["scale_width"] = sw
                    except Exception: pass
                except Exception as re:
                    run["run_error"] = str(re)
                runs.append(run)
        except Exception as ee:
            entry["runs_error"] = str(ee)
        entry["runs"] = runs
        return entry

    def read_table_rich(self, table_name: str = None,
                        table_index: int = None) -> Dict[str, Any]:
        """Read each cell as a list of paragraphs with runs (the same shape
        emitted by get_paragraphs_with_runs). Use this to faithfully replicate
        a table — read_table_cells only returns plain text and loses font /
        bold / alignment / indent inside cells."""
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
                    return {"success": False, "error": "table_index out of range"}
                t = tables.getByIndex(table_index)
            else:
                if tables.getCount() == 0:
                    return {"success": False, "error": "no tables"}
                t = tables.getByIndex(0)
            rows = t.getRows().getCount()
            cols = t.getColumns().getCount()
            # Column widths — критично для replicate. UNO хранит их через
            # TableColumnSeparators (массив N-1 точек разреза 0..10000 в
            # relative units от ширины таблицы) + Width (абсолют 1/100 mm).
            # Без этого insert_table в target создаёт колонки одинаковой
            # ширины, и текст в узких ячейках получается уродливо растянут
            # (особенно при ParaAdjust=BLOCK_LINE).
            column_widths_mm = []
            table_width_mm = None
            try:
                tw = int(getattr(t, "Width", 0) or 0)
                if tw > 0:
                    table_width_mm = tw / 100.0
                    seps = list(getattr(t, "TableColumnSeparators", []) or [])
                    prev = 0
                    for sep in seps:
                        column_widths_mm.append((sep.Position - prev) * tw / 10000.0 / 100.0)
                        prev = sep.Position
                    column_widths_mm.append((10000 - prev) * tw / 10000.0 / 100.0)
            except Exception:
                pass
            # Table-level page-break behaviour. Without Split=True the table
            # cannot span pages — when it doesn't fit on the current page it
            # jumps wholesale to the next, leaving a big empty gap.
            # RepeatHeadline + HeaderRowCount control whether the header row
            # repeats at the top of every page the table spans.
            try: t_split = bool(getattr(t, "Split", True))
            except Exception: t_split = True
            try: t_repeat = bool(getattr(t, "RepeatHeadline", False))
            except Exception: t_repeat = False
            try: t_header_rows = int(getattr(t, "HeaderRowCount", 0) or 0)
            except Exception: t_header_rows = 0
            try: t_keep_together = bool(getattr(t, "KeepTogether", False))
            except Exception: t_keep_together = False
            grid = []
            for r in range(rows):
                row_arr = []
                for c in range(cols):
                    cn = chr(ord("A") + c) + str(r + 1)
                    try:
                        cell = t.getCellByName(cn)
                    except Exception:
                        cell = None
                    if cell is None:
                        row_arr.append({"name": cn, "paragraphs": []})
                        continue
                    paras = []
                    try:
                        ctext = cell.getText()
                        cenum = ctext.createEnumeration()
                        while cenum.hasMoreElements():
                            elem = cenum.nextElement()
                            if not elem.supportsService("com.sun.star.text.Paragraph"):
                                continue
                            paras.append(self._extract_para_with_runs(elem))
                    except Exception as ce:
                        row_arr.append({"name": cn, "paragraphs": [],
                                        "error": str(ce)})
                        continue
                    cell_entry = {"name": cn, "paragraphs": paras}
                    # Per-cell vertical alignment (TOP=0,CENTER=1,BOTTOM=2)
                    try:
                        va = cell.VertOrient
                        cell_entry["vert_orient"] = int(getattr(va, "value", va)) if not isinstance(va, int) else va
                    except Exception: pass
                    row_arr.append(cell_entry)
                grid.append(row_arr)
            return {"success": True, "name": t.getName(), "rows": rows,
                    "columns": cols, "table_width_mm": table_width_mm,
                    "column_widths_mm": column_widths_mm,
                    "split": t_split, "repeat_headline": t_repeat,
                    "header_row_count": t_header_rows,
                    "keep_together": t_keep_together,
                    "cells": grid}
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

    def insert_text_frame(self, paragraph_index: int = None,
                          width_mm: float = 6.7, height_mm: float = 4.94,
                          text: str = None, page_number: bool = False,
                          hori_orient: str = "center", vert_orient: str = "bottom",
                          hori_relation: str = "page", vert_relation: str = "page",
                          x_mm: float = 0.0, y_mm: float = 0.0,
                          back_transparent: bool = True,
                          remove_borders: bool = True) -> Dict[str, Any]:
        """Insert a TextFrame anchored to a paragraph (AT_PARAGRAPH).
        Supports embedding a PageNumber field via page_number=True OR plain
        text via `text`. Use this to replicate Word's docshape page-number
        boxes that sit at page bottoms outside of FooterText.

        hori_orient/vert_orient: 'left'|'center'|'right'|'top'|'bottom'|'none'
        relation: 'frame'|'paragraph'|'page'|'page-text-area'
        x_mm/y_mm used only when orient='none' (manual offset).
        """
        doc, err = self._require_writer()
        if err: return err
        try:
            text_obj = doc.getText()
            target = None
            if paragraph_index is not None:
                pe = text_obj.createEnumeration()
                i = 0
                while pe.hasMoreElements():
                    p = pe.nextElement()
                    if p.supportsService("com.sun.star.text.Paragraph"):
                        if i == paragraph_index:
                            target = p; break
                        i += 1
                if target is None:
                    return {"success": False, "error": f"paragraph {paragraph_index} not found"}
            anchor = target if target is not None else text_obj.createTextCursor()

            frame = doc.createInstance("com.sun.star.text.TextFrame")
            sz = uno.createUnoStruct("com.sun.star.awt.Size")
            sz.Width = int(float(width_mm) * 100)
            sz.Height = int(float(height_mm) * 100)
            frame.Size = sz

            text_obj.insertTextContent(anchor.getStart() if hasattr(anchor, "getStart") else text_obj.getEnd(),
                                       frame, False)

            # AnchorType set after insertion (some props lock once attached)
            try:
                from com.sun.star.text.TextContentAnchorType import AT_PARAGRAPH
                frame.AnchorType = AT_PARAGRAPH
            except Exception: pass

            # com.sun.star.text.HoriOrientation: NONE=0, RIGHT=1, CENTER=2,
            # LEFT=3, INSIDE=4, OUTSIDE=5, FULL=6.
            HORI_MAP = {"none": 0, "right": 1, "center": 2, "left": 3,
                        "inside": 4, "outside": 5, "full": 6}
            # com.sun.star.text.VertOrientation: NONE=0, TOP=1, CENTER=2,
            # BOTTOM=3, CHAR_TOP=4, ..., LINE_BOTTOM=8.
            VERT_MAP = {"none": 0, "top": 1, "center": 2, "bottom": 3,
                        "char_top": 4, "char_center": 5, "char_bottom": 6,
                        "line_top": 7, "line_center": 8, "line_bottom": 9}
            # com.sun.star.text.RelOrientation: FRAME=0, PRINT_AREA=1,
            # CHAR=2, PAGE_LEFT=3, PAGE_RIGHT=4, FRAME_LEFT=5, FRAME_RIGHT=6,
            # PAGE_FRAME=7, PAGE_PRINT_AREA=8, TEXT_LINE=9.
            # Common aliases: 'page' = PAGE_FRAME (7), 'page-text-area' =
            # PAGE_PRINT_AREA (8), 'paragraph' = FRAME (0).
            REL_MAP = {"frame": 0, "paragraph": 0, "print_area": 1,
                       "char": 2, "page-left": 3, "page-right": 4,
                       "frame-left": 5, "frame-right": 6,
                       "page": 7, "page-frame": 7,
                       "page-text-area": 8, "page-print-area": 8,
                       "page-content": 8, "text-line": 9}
            try: frame.HoriOrient = HORI_MAP.get(hori_orient.lower(), 2)
            except Exception: pass
            try: frame.VertOrient = VERT_MAP.get(vert_orient.lower(), 3)
            except Exception: pass
            try: frame.HoriOrientRelation = REL_MAP.get(hori_relation.lower(), 7)
            except Exception: pass
            try: frame.VertOrientRelation = REL_MAP.get(vert_relation.lower(), 7)
            except Exception: pass
            if hori_orient.lower() == "none":
                try: frame.HoriOrientPosition = int(float(x_mm) * 100)
                except Exception: pass
            else:
                # Some LO versions need an explicit zero offset alongside
                # orient=center, otherwise leftover positions kick in.
                try: frame.HoriOrientPosition = 0
                except Exception: pass
            if vert_orient.lower() == "none":
                try: frame.VertOrientPosition = int(float(y_mm) * 100)
                except Exception: pass
            else:
                try: frame.VertOrientPosition = 0
                except Exception: pass
            if back_transparent:
                try: frame.BackTransparent = True
                except Exception: pass
            if remove_borders:
                try:
                    bl = uno.createUnoStruct("com.sun.star.table.BorderLine2")
                    bl.OuterLineWidth = 0
                    for bp in ("LeftBorder", "RightBorder", "TopBorder", "BottomBorder"):
                        try: setattr(frame, bp, bl)
                        except Exception: pass
                except Exception: pass

            # Insert content into the frame
            ftxt = frame.getText()
            cur = ftxt.createTextCursor()
            if page_number:
                fld = doc.createInstance("com.sun.star.text.TextField.PageNumber")
                try:
                    from com.sun.star.style.NumberingType import ARABIC
                    fld.NumberingType = ARABIC
                except Exception: pass
                ftxt.insertTextContent(cur, fld, False)
            elif text:
                ftxt.insertString(cur, text, False)
            return {"success": True, "name": frame.Name,
                    "anchor_paragraph_index": paragraph_index}
        except Exception as e:
            return {"success": False, "error": str(e), "trace": traceback.format_exc()}

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
                     name: str = None, column_widths_mm: list = None,
                     table_width_mm: float = None,
                     split: bool = None, repeat_headline: bool = None,
                     header_row_count: int = None,
                     keep_together: bool = None) -> Dict[str, Any]:
        """Insert a TextTable.

        column_widths_mm: optional list of width per column (mm) — applied
            via TableColumnSeparators after insertion. Critical for tables
            where one column is much narrower/wider; without it all columns
            are equal width and text in narrow cells gets brutally wrapped.
        table_width_mm: total table width in mm. If omitted, derived from
            sum(column_widths_mm) or left at default.
        """
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
            # Apply column widths AFTER insertion (table.Width / Separators
            # don't take effect until the table is in the doc).
            applied_cols = None
            if column_widths_mm:
                try:
                    widths = [float(w) for w in column_widths_mm if w is not None]
                    if len(widths) == columns and all(w > 0 for w in widths):
                        # Set TableWidth first if requested
                        if table_width_mm is None:
                            table_width_mm = sum(widths)
                        try: table.Width = int(float(table_width_mm) * 100)
                        except Exception: pass
                        # Build TableColumnSeparators (n-1 separators)
                        total = sum(widths)
                        cum = 0
                        seps_existing = list(getattr(table, "TableColumnSeparators", []) or [])
                        # We need n-1 separators with Position in 0..10000
                        new_seps = []
                        # createUnoStruct for each separator
                        for i in range(columns - 1):
                            cum += widths[i]
                            sep = uno.createUnoStruct("com.sun.star.text.TableColumnSeparator")
                            sep.Position = int(cum / total * 10000)
                            sep.IsVisible = True
                            new_seps.append(sep)
                        table.TableColumnSeparators = tuple(new_seps)
                        applied_cols = widths
                except Exception as we:
                    logger.warning(f"insert_table column_widths failed: {we}")
            # Page-flow behaviour
            if split is not None:
                try: table.Split = bool(split)
                except Exception: pass
            if keep_together is not None:
                try: table.KeepTogether = bool(keep_together)
                except Exception: pass
            if repeat_headline is not None:
                try: table.RepeatHeadline = bool(repeat_headline)
                except Exception: pass
            if header_row_count is not None:
                try: table.HeaderRowCount = int(header_row_count)
                except Exception: pass
            return {"success": True, "name": table.getName(),
                    "rows": rows, "columns": columns,
                    "applied_column_widths": applied_cols}
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

    def _apply_run_props(self, cur, run: dict, skip_word_justify_artifacts: bool = False):
        """Apply per-portion character properties from a run dict onto cursor.

        skip_word_justify_artifacts: when True, ignore kerning and scale_width.
            Word→ODT exporter encodes pre-applied justify-stretching as
            per-portion CharKerning / CharScaleWidth on space portions.
            These values are only valid at the exact original line width;
            re-applying them on top of ParaAdjust=block_line (which itself
            stretches to fit the new container) yields double-stretching —
            visible in narrow table cells where text becomes visually
            "разорванным": "п о л н о с т ь ю". Use this flag inside table
            cells where ParaAdjust handles justification on its own.
        """
        if run.get("font_name"):
            try: cur.CharFontName = run["font_name"]
            except Exception: pass
        if run.get("font_size") is not None:
            try: cur.CharHeight = float(run["font_size"])
            except Exception: pass
        if run.get("bold"):
            try: cur.CharWeight = 150.0
            except Exception: pass
        if run.get("italic"):
            try: cur.CharPosture = uno.Enum("com.sun.star.awt.FontSlant", "ITALIC")
            except Exception: pass
        if run.get("underline"):
            try: cur.CharUnderline = 1
            except Exception: pass
        if not skip_word_justify_artifacts:
            if run.get("kerning") is not None:
                try: cur.CharKerning = int(run["kerning"])
                except Exception: pass
            if run.get("scale_width") is not None:
                try: cur.CharScaleWidth = int(run["scale_width"])
                except Exception: pass
        if run.get("color"):
            try: cur.CharColor = self._hex_to_int(run["color"])
            except Exception: pass
        if run.get("background_color"):
            try: cur.CharBackColor = self._hex_to_int(run["background_color"])
            except Exception: pass
        if run.get("char_style"):
            try: cur.CharStyleName = run["char_style"]
            except Exception: pass
        if run.get("hyperlink"):
            try: cur.HyperLinkURL = run["hyperlink"]
            except Exception: pass

    def _apply_paragraph_props(self, cur, p: dict):
        """Apply per-paragraph properties (style, alignment, indent, spacing,
        line spacing, tab stops) onto a cursor anchored inside the paragraph."""
        if p.get("style"):
            try: cur.ParaStyleName = p["style"]
            except Exception: pass
        pa = p.get("paragraph_adjust")
        if pa is not None:
            try: cur.ParaAdjust = int(pa)
            except Exception: pass
        if p.get("left_mm") is not None:
            try: cur.ParaLeftMargin = int(float(p["left_mm"]) * 100)
            except Exception: pass
        if p.get("right_mm") is not None:
            try: cur.ParaRightMargin = int(float(p["right_mm"]) * 100)
            except Exception: pass
        if p.get("first_line_mm") is not None:
            try: cur.ParaFirstLineIndent = int(float(p["first_line_mm"]) * 100)
            except Exception: pass
        if p.get("top_mm") is not None:
            try: cur.ParaTopMargin = int(float(p["top_mm"]) * 100)
            except Exception: pass
        if p.get("bottom_mm") is not None:
            try: cur.ParaBottomMargin = int(float(p["bottom_mm"]) * 100)
            except Exception: pass
        ls = p.get("line_spacing") or {}
        mode = ls.get("mode"); val = ls.get("value")
        if mode and val is not None:
            try:
                from com.sun.star.style import LineSpacing as _LS
                mode_map = {"proportional": 0, "minimum": 1, "leading": 2, "fix": 3}
                ls_mode = mode_map.get(str(mode).lower(), 0)
                spacing = _LS()
                spacing.Mode = ls_mode
                spacing.Height = int(float(val)) if ls_mode == 0 else int(float(val) * 100)
                cur.ParaLineSpacing = spacing
            except Exception: pass
        if p.get("context_margin") is not None:
            try: cur.ParaContextMargin = bool(p["context_margin"])
            except Exception: pass
        tabs = p.get("tab_stops") or []
        if tabs:
            try:
                # Reuse same encoding-back logic as set_paragraph_tabs
                self._apply_tab_stops(cur, tabs)
            except Exception: pass

    def _apply_tab_stops(self, cur, stops):
        """Encode list-of-dict tab stops into a UNO TabStop[] sequence and
        assign to cur.ParaTabStops. Mirrors set_paragraph_tabs encoding."""
        align_map = {"left": 0, "center": 1, "right": 2, "decimal": 3}
        out = []
        for s in stops:
            if not isinstance(s, dict): continue
            t = uno.createUnoStruct("com.sun.star.style.TabStop")
            t.Position = int(float(s.get("position_mm", 0)) * 100)
            t.Alignment = align_map.get(str(s.get("alignment", "left")).lower(), 0)
            fc = s.get("fill_char") or " "
            dc = s.get("decimal_char") or "."
            t.FillChar = ord(fc[0]) if isinstance(fc, str) and fc else 32
            t.DecimalChar = ord(dc[0]) if isinstance(dc, str) and dc else 46
            out.append(t)
        cur.ParaTabStops = tuple(out)

    def write_table_cell_rich(self, table_name: str, cell: str,
                              paragraphs: list) -> Dict[str, Any]:
        """Write a list of paragraphs (text + runs + paragraph props) into a
        table cell, preserving formatting that read_table_rich captured.

        paragraphs: list of dicts with optional keys 'text', 'style',
            'paragraph_adjust', 'left_mm', 'right_mm', 'first_line_mm',
            'top_mm', 'bottom_mm', 'line_spacing', 'tab_stops',
            'context_margin', 'runs'. Each run can have 'text', 'font_name',
            'font_size', 'bold', 'italic', 'underline', 'kerning',
            'scale_width', 'color', 'background_color', 'hyperlink', 'char_style'.
        """
        doc, err = self._require_writer()
        if err:
            return err
        if not isinstance(paragraphs, list):
            return {"success": False, "error": "paragraphs must be a list"}
        try:
            tables = doc.getTextTables()
            if not tables.hasByName(table_name):
                return {"success": False, "error": f"no table named '{table_name}'"}
            t = tables.getByName(table_name)
            c = t.getCellByName(cell)
            if c is None:
                return {"success": False, "error": f"cell '{cell}' not found"}
            text = c.getText()
            text.setString("")
            from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
            cursor = text.createTextCursor()
            cursor.gotoStart(False)
            for pi, p in enumerate(paragraphs):
                if pi > 0:
                    text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
                runs = p.get("runs") or []
                # Insert runs with per-run formatting; if no runs, just plain text.
                ptext = p.get("text") or ""
                if runs:
                    for run in runs:
                        rt = run.get("text") or ""
                        if not rt:
                            continue
                        anchor_end = cursor.getEnd()
                        text.insertString(cursor, rt, False)
                        sel = text.createTextCursorByRange(anchor_end)
                        sel.gotoRange(cursor.getEnd(), True)
                        # Skip kerning/scale_width inside table cells — Word
                        # exporter pre-bakes justify-stretching into them at
                        # the source column width; re-applying here on top
                        # of ParaAdjust=block_line yields visible double
                        # stretching ("п о л н о с т ь ю").
                        self._apply_run_props(sel, run, skip_word_justify_artifacts=True)
                elif ptext:
                    text.insertString(cursor, ptext, False)
                # Apply paragraph properties on a cursor inside the just-written paragraph
                pcur = text.createTextCursorByRange(cursor.getEnd())
                self._apply_paragraph_props(pcur, p)
            return {"success": True, "table": table_name, "cell": cell,
                    "paragraphs": len(paragraphs)}
        except Exception as e:
            return {"success": False, "error": str(e),
                    "trace": traceback.format_exc()}

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

    def set_footer_page_number(self, page_style: str = "Default Page Style",
                               alignment: str = "center",
                               font_size: float = None) -> Dict[str, Any]:
        """Replace footer content with a centered (or other-aligned) PageNumber
        field. Use this when source numbers each page and you need it
        replicated on every page using a given page-style — far simpler than
        anchoring per-page TextFrames. alignment: 'left'|'center'|'right'.
        """
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ps = self._page_style(doc, page_style)
            if not ps.FooterIsOn:
                ps.FooterIsOn = True
            # Match Word→ODT export footer geometry (HeaderIsOn=false,
            # FooterIsOn=true, FooterHeight=3.44mm, BodyDistance=0,
            # IsDynamicHeight=true). Default UNO footer reserves more space
            # which shifts body-area and disturbs page-break placement.
            try: ps.FooterIsDynamicHeight = True
            except Exception: pass
            try: ps.FooterHeight = 344          # 3.44 mm
            except Exception: pass
            try: ps.FooterBodyDistance = 0
            except Exception: pass
            footer = ps.FooterText
            footer.setString("")
            cursor = footer.createTextCursor()
            align_map = {"left": 0, "right": 1, "justify": 2, "center": 3}
            cursor.ParaAdjust = align_map.get(str(alignment).lower(), 3)
            if font_size is not None:
                try: cursor.CharHeight = float(font_size)
                except Exception: pass
            field = doc.createInstance("com.sun.star.text.TextField.PageNumber")
            try:
                field.NumberingType = 4  # ARABIC
                field.SubType = 1        # CURRENT
            except Exception:
                pass
            footer.insertTextContent(cursor, field, False)
            return {"success": True, "page_style": page_style,
                    "alignment": alignment}
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

    def read_paragraph_xml(self, source_path: str, paragraph_index: int,
                           include_styles: bool = True) -> Dict[str, Any]:
        """Read raw ODT XML for a paragraph and the styles it references.

        UNO paragraph properties surface only resolved values for a fixed
        set of fields — Word→ODT exporters often emit `fo:*` attributes
        (letter-spacing, break-before, keep-with-next, hyphenate, etc.)
        on automatic styles which never round-trip through pyuno. Parsing
        content.xml + styles.xml directly is the only way to see them.

        Returns:
          - paragraph_xml: raw `<text:p ... >...</text:p>` or
            `<text:h ... >...</text:h>` element as a string
          - style_name: paragraph style ref attribute
          - styles: dict {style_name: raw_style_xml} for the paragraph's
            style chain (parent-style-name walked) plus every text-span's
            referenced T-style. Use this to find fo:* attributes that
            UNO API hides.
        """
        try:
            import zipfile
            import re as _re
            with zipfile.ZipFile(source_path) as zf:
                with zf.open("content.xml") as f:
                    content = f.read().decode("utf-8")
                styles_xml = ""
                try:
                    with zf.open("styles.xml") as f:
                        styles_xml = f.read().decode("utf-8")
                except KeyError:
                    pass
            # Iterate ALL text:p and text:h in body order — same way UNO
            # _iter_paragraphs does (paragraphs in body, top-level).
            # Need to skip paragraphs nested INSIDE table cells to match
            # the index that get_paragraphs returns.
            # Strategy: strip out <table:table>...</table:table> blocks
            # before iterating paragraphs (their inner paragraphs are
            # not body-level).
            body_only = _re.sub(r'<table:table\b[^>]*>.*?</table:table>',
                                '<table:table-stub/>', content, flags=_re.DOTALL)
            # Iterate text:p / text:h — including self-closing variants
            # (`<text:p .../>` is the ODT representation of an empty
            # paragraph; without matching them the index drifts by ~44
            # for a typical Word→ODT export with empty separator paras).
            iters = list(_re.finditer(
                r'<text:(p|h)\b[^>]*?(?:/>|>.*?</text:\1>)',
                body_only, flags=_re.DOTALL))
            if paragraph_index < 0 or paragraph_index >= len(iters):
                return {"success": False,
                        "error": f"index out of range (0..{len(iters)-1})",
                        "total": len(iters)}
            para_xml = iters[paragraph_index].group()
            # Style name on the paragraph
            sm = _re.search(r'text:style-name="([^"]+)"', para_xml)
            style_name = sm.group(1) if sm else ""
            result = {"success": True, "paragraph_index": paragraph_index,
                      "total_paragraphs": len(iters),
                      "paragraph_xml": para_xml, "style_name": style_name}
            if not include_styles:
                return result
            # Walk style chain + collect T-styles referenced in spans
            referenced = {}
            seen = set()
            queue = [style_name] if style_name else []
            # Add T-styles
            for ts in set(_re.findall(r'<text:span text:style-name="([^"]+)"', para_xml)):
                queue.append(ts)
            while queue:
                sn = queue.pop()
                if sn in seen:
                    continue
                seen.add(sn)
                pat = _re.compile(rf'<style:style style:name="{_re.escape(sn)}"[^>]*>.*?</style:style>',
                                  _re.DOTALL)
                m = pat.search(content) or pat.search(styles_xml)
                if m:
                    sxml = m.group()
                    referenced[sn] = sxml
                    pm = _re.search(r'style:parent-style-name="([^"]+)"', sxml)
                    if pm:
                        queue.append(pm.group(1))
            result["styles"] = referenced
            return result
        except FileNotFoundError:
            return {"success": False, "error": f"file not found: {source_path!r}"}
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

    def _doc_window(self, doc=None):
        """Return the active doc's container window (or None)."""
        try:
            d = doc if doc is not None else self.get_active_document()
            if d is None: return None
            ctrl = d.getCurrentController()
            if ctrl is None: return None
            frame = ctrl.getFrame()
            if frame is None: return None
            return frame.getContainerWindow()
        except Exception:
            return None

    def show_window(self) -> Dict[str, Any]:
        """Make the active document's window visible. Pair with hide_window
        around bulk writes that hit PageStyle / paragraph style mutations on
        macOS — visible-window paint cycles hold SolarMutex and deadlock the
        HTTP-worker thread."""
        win = self._doc_window()
        if win is None:
            return {"success": False, "error": "No container window"}
        try:
            win.setVisible(True)
            return {"success": True}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def hide_window(self) -> Dict[str, Any]:
        """Hide the active document's window. Use BEFORE a burst of
        clone_page_style / clone_paragraph_style / set_page_style_props /
        set_paragraph_style_props on macOS, then call show_window() after.
        Document model stays alive — only the visual frame is detached.
        execute_batch with auto_hide=true does this automatically."""
        win = self._doc_window()
        if win is None:
            return {"success": False, "error": "No container window"}
        try:
            was_visible = bool(win.isVisible())
            win.setVisible(False)
            return {"success": True, "was_visible": was_visible}
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
