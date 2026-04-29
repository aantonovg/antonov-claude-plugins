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
            logger.info("UNO Bridge initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize UNO Bridge: {e}")
            raise
    
    def create_document(self, doc_type: str = "writer") -> Any:
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
            try:
                ctrl = doc.getCurrentController()
                if ctrl is not None:
                    frame = ctrl.getFrame()
                    if frame is not None:
                        win = frame.getContainerWindow()
                        if win is not None:
                            win.setVisible(True)
            except Exception as e:
                logger.warning(f"Created document but could not show window: {e}")
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
    
    def insert_text(self, text: str, position: Optional[int] = None, doc: Any = None) -> Dict[str, Any]:
        """
        Insert text into a document
        
        Args:
            text: Text to insert
            position: Position to insert at (None for current cursor position)
            doc: Document to insert into (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc:
                return {"success": False, "error": "No active document"}
            
            if doc.supportsService("com.sun.star.text.TextDocument"):
                text_obj = doc.getText()

                if position is None:
                    cursor = doc.getCurrentController().getViewCursor()
                else:
                    cursor = text_obj.createTextCursor()
                    cursor.gotoStart(False)
                    cursor.goRight(position, False)

                # com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK = 0
                # Split on '\n' so newlines become real paragraph breaks
                # instead of literal characters in one paragraph.
                parts = text.split("\n")
                for i, part in enumerate(parts):
                    if i > 0:
                        text_obj.insertControlCharacter(cursor, 0, False)
                    if part:
                        text_obj.insertString(cursor, part, False)
                logger.info(f"Inserted {len(text)} characters into Writer document")
                return {"success": True, "message": f"Inserted {len(text)} characters"}

            return {"success": False, "error": f"Text insertion not supported for {self._get_document_type(doc)}"}
                
        except Exception as e:
            logger.error(f"Failed to insert text: {e}")
            return {"success": False, "error": str(e)}
    
    def format_text(self, formatting: Dict[str, Any], doc: Any = None) -> Dict[str, Any]:
        """
        Apply formatting to selected text
        
        Args:
            formatting: Dictionary of formatting options
            doc: Document to format (None for active document)
            
        Returns:
            Result dictionary
        """
        try:
            if doc is None:
                doc = self.get_active_document()
            
            if not doc or not doc.supportsService("com.sun.star.text.TextDocument"):
                return {"success": False, "error": "No Writer document available"}
            
            # Get current selection
            selection = doc.getCurrentController().getSelection()
            if selection.getCount() == 0:
                return {"success": False, "error": "No text selected"}
            
            # Apply formatting to selection
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

    def _require_writer(self):
        doc = self.get_active_document()
        if not doc or not doc.supportsService("com.sun.star.text.TextDocument"):
            return None, {"success": False, "error": "No Writer document active"}
        return doc, None

    def set_text_color(self, color) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            rng = self._selected_range_or_view_cursor(doc)
            rng.CharColor = self._hex_to_int(color)
            return {"success": True, "color": color}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_background_color(self, color) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            rng = self._selected_range_or_view_cursor(doc)
            rng.CharBackColor = self._hex_to_int(color)
            return {"success": True, "color": color}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_alignment(self, alignment: str) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        # com.sun.star.style.ParagraphAdjust: LEFT=0, RIGHT=1, BLOCK=2, CENTER=3
        mapping = {"left": 0, "right": 1, "justify": 2, "block": 2, "center": 3}
        val = mapping.get(alignment.lower())
        if val is None:
            return {"success": False, "error": "Unknown alignment, use: left|center|right|justify"}
        try:
            rng = self._selected_range_or_view_cursor(doc)
            rng.ParaAdjust = val
            return {"success": True, "alignment": alignment}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_paragraph_indent(self, left_mm=None, right_mm=None, first_line_mm=None) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            rng = self._selected_range_or_view_cursor(doc)
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

    def set_line_spacing(self, mode: str = "proportional", value: float = 100) -> Dict[str, Any]:
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
            rng = self._selected_range_or_view_cursor(doc)
            rng.ParaLineSpacing = ls
            return {"success": True, "mode": mode, "value": value}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def apply_paragraph_style(self, style_name: str) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            rng = self._selected_range_or_view_cursor(doc)
            rng.ParaStyleName = style_name
            return {"success": True, "style": style_name}
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
                        ls = para.ParaLineSpacing
                        entry["line_spacing"] = {"mode": ["proportional","minimum","leading","fix"][ls.Mode] if ls.Mode in (0,1,2,3) else ls.Mode,
                                                 "value": ls.Height if ls.Mode == 0 else ls.Height / 100.0}
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
                            run["italic"] = portion.CharPosture != 0
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
                "italic": cursor.CharPosture != 0,
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

    def get_page_info(self) -> Dict[str, Any]:
        doc, err = self._require_writer()
        if err:
            return err
        try:
            ctrl = doc.getCurrentController()
            # PageCount is a property on the model in modern LO, not a controller method
            page_count = None
            try:
                page_count = doc.getPropertyValue("PageCount")
            except Exception:
                if hasattr(ctrl, "getPageCount"):
                    try:
                        page_count = ctrl.getPageCount()
                    except Exception:
                        pass
            view_cursor = ctrl.getViewCursor()
            current_page = None
            try:
                current_page = view_cursor.getPage() if hasattr(view_cursor, "getPage") else None
            except Exception:
                pass
            # page size — from the resolved default page style
            width_mm = height_mm = None
            try:
                ps = self._page_style(doc)
                size = ps.Size
                width_mm = size.Width / 100.0
                height_mm = size.Height / 100.0
            except Exception:
                pass
            return {"success": True, "page_count": page_count, "current_page": current_page,
                    "page_width_mm": width_mm, "page_height_mm": height_mm}
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
        """
        try:
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
