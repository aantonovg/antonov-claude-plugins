---
name: libreoffice-word
description: Read, edit, replicate, format, compare, or convert documents in any Word-style format — .odt, .doc, .docx, .pages (after manual conversion), .rtf, .txt, .html — using the two LibreOffice MCP servers (live HTTP API for editing an open document, headless subprocess for closed-file conversion). Activate whenever the user asks to open, view, modify, format, replicate, diff, or convert one of these documents.
user-invocable: false
---

# LibreOffice / Word documents — agent guide

## When to activate

Activate this skill any time the user asks to:

- **Read**: open, inspect, summarise, search, compare, diff, extract content from a `.odt` / `.doc` / `.docx` / `.rtf` / `.txt` / `.html` / `.epub` document (or a `.pages` file after they convert it — see Format Support).
- **Edit**: insert, delete, format, rename, replace text; change paragraph alignment / indents / spacing / line spacing / styles / numbering; insert / modify / remove tables, hyperlinks, bookmarks, comments, images, headers, footers, page-numbering frames; change page margins / page styles.
- **Replicate**: rebuild a document via batch operations (programmatic equivalent of clone, not a binary copy).
- **Convert**: change format (e.g. `.docx` → `.odt`, `.odt` → `.pdf`, `.doc` → `.html`).
- **Diff**: structurally compare two documents.

The user does NOT need to mention "LibreOffice" — only the format hint is enough. They might say "look at this contract", "fix the formatting in foo.docx", "extract paragraphs from bar.odt", "convert this to PDF".

---

## Two MCP servers — the most important architectural decision

There are **two** LibreOffice MCP servers wired into Claude. They look similar, but they cover different use cases and have different failure modes. **The correct choice often determines whether the task succeeds or hangs.**

### A) `libreoffice-live` (HTTP API at `localhost:8765`)

Tool prefix: `mcp__libreoffice-live__*` (or `mcp__plugin_libreoffice_libreoffice-live__*`).

- Talks to a **running LibreOffice GUI** via a Python extension.
- The user **sees changes live** in the open window — instant visual feedback.
- Best for: **editing**, **inspecting**, **replicating**, **formatting**, **interactive workflows**.
- ⚠️ **Save through UNO is broken on macOS** (UI-thread deadlock with AppKit). Always tell the user to press `Cmd+S` after edits.
- ⚠️ Operations on a **visible** window can deadlock when invoked from a background thread (SolarMutex). Use `visible=False` + `show_window` for new docs, and `auto_hide="auto"` (default) in `execute_batch`.
- ⚠️ `export_active_document` is **disabled** on macOS for the same reason — use `clone_document` instead.

### B) `libreoffice` (headless subprocess `soffice --headless`)

Tool prefix: `mcp__libreoffice__*`.

- Forks a fresh `soffice --headless` process per call. No GUI, no shared state.
- ~1–2 seconds per call (process startup overhead).
- Save **works** here — the headless process can write files.
- Best for: **format conversion of files on disk**, **batch processing**, **reading closed files without bothering the user's open session**.
- Cannot edit live; every "edit" goes through load → modify → save → close.

### Decision matrix

| Goal | Server | Why |
|---|---|---|
| Read open document | `libreoffice-live` | Instant, no fork |
| Read closed document, then close | `libreoffice` | Don't intrude on user session |
| Edit document user has open | `libreoffice-live` | They see changes; ask Cmd+S to save |
| Convert one file (e.g. `foo.docx` → `foo.pdf`) | `libreoffice-live` `clone_document` | Hidden component, no UI deadlock; works on macOS |
| Convert dozens of files | `libreoffice` `batch_convert_documents` | Headless is built for this |
| Create a new document and write to it | `libreoffice-live` `create_document_live(visible=False)` then ops, then `show_window` | Live; safe via hidden init |
| Run a `.uno:*` command | `libreoffice-live` `dispatch_uno_command` | Whitelisted, no fork |

When unsure: **start with `libreoffice-live`** if a window is open or the user wants to see results; **fall back to `libreoffice`** if you specifically need save-to-disk and don't want the user involved.

---

## Format support

| Format | Read | Write | Notes |
|---|---|---|---|
| `.odt` | ✔ | ✔ | Native; round-trips perfectly |
| `.docx` | ✔ | ✔ | Word 2007+ (`MS Word 2007 XML` filter) |
| `.doc` | ✔ | ✔ | Word 97-2003 |
| `.rtf` | ✔ | ✔ | |
| `.txt` | ✔ | ✔ | Plain text |
| `.html`, `.xhtml` | ✔ | ✔ | StarWriter HTML |
| `.epub` | ✔ | ✔ | Export only |
| `.pdf` | (visible only) | ✔ | Export only — read PDFs in another tool |
| `.pages` | ✘ | ✘ | LibreOffice does **NOT** support Apple Pages. Tell the user: "Open in Apple Pages → File → Export To → Word (.docx) → re-attach." |

For non-Writer formats (`.pptx`, `.xlsx`, `.ods`, `.csv`, `.odp`) — supported by `clone_document` for conversion, but this skill focuses on Writer (text) documents.

---

## Read inventory — how to "see" a document

A document is a hierarchy: **page-styles → body → paragraphs (and tables interleaved) → runs (uniform-format text portions)**. To replicate or analyse, you must read every level.

### High-level shape

| Tool | Returns |
|---|---|
| `get_document_info_live` | Title, URL, modified flag, type, char/word count, has_selection |
| `get_document_summary` | One-shot overview: paragraph count, table count, image count, hyperlink count, etc. — call this **first** for unfamiliar docs |
| `get_page_info` | Page size mm, margins T/B/L/R, page_style name, page_count, header/footer enabled + height + text |
| `get_outline` | Heading 1..N hierarchy with positions |
| `list_open_documents` | All open windows with titles/URLs |

### Body iteration (paragraphs + tables, in order)

| Tool | When to use |
|---|---|
| `list_body_elements` | **Always first** when the document has tables. Returns paragraphs AND tables in correct order, with `after_paragraph_index` for each table. `get_paragraphs` silently skips tables — using only it loses their positions. |
| `get_paragraphs(count, include_format=True)` | Per-paragraph metadata: index, start, end, length, style, alignment + paragraph_adjust (raw int 0–5), left/right/first_line_mm, top/bottom_mm, line_spacing, tab_stops, list_label, numbering_*, page_desc_name, break_type, **char_height**, **widows / orphans / keep_together / split_paragraph / keep_with_next** |
| `get_paragraphs_with_runs(count)` | Same as above + `runs[]`: per-portion font_name, font_size, bold, italic, underline, color, hyperlink, char_style, kerning, scale_width, char_posture |
| `get_paragraph_format_at(position)` | Paragraph format at a single offset (when iterating not needed) |
| `get_character_format(start, end)` | Character format for a range |

### Tables

| Tool | Returns |
|---|---|
| `get_tables_info` | List of tables with name + rows + columns (no positions, no widths) |
| `read_table_rich(table_name)` | **Use this for replication.** Per cell: array of paragraphs with full metadata + runs. Plus `column_widths_mm`, `table_width_mm`, `split`, `repeat_headline`, `header_row_count`, `keep_together`, per-cell `vert_orient`. |
| `read_table_cells(table_name)` | Plain string grid only — use only when formatting doesn't matter |

### Frames (page-number boxes, anchored content)

`list_text_frames` — every TextFrame with name, size_mm, pos_mm, anchor_type, **anchor_para_index**, HoriOrient/VertOrient/HoriOrientPosition/VertOrientPosition, HoriOrientRelation/VertOrientRelation, BackTransparent, borders, **fields[]** (TextField service names — e.g. `PageNumber`).

⚠️ `anchor_para_index` is best-effort. For documents with N empty-text frames anchored to N different paragraphs (typical Word-export per-page page-numbers), pyuno's identity / region-comparison APIs do not reliably distinguish anchors. **Fallback strategy**: pair the N-th frame with the N-th paragraph that has a `page_desc_name` — Word→ODT export typically emits one frame per master-page-switch.

### Sections, hyperlinks, bookmarks, comments, images

`list_hyperlinks`, `list_bookmarks`, `list_comments`, `list_images`, `list_sections`, `list_paragraph_styles`, `list_character_styles`.

### Hidden properties — when UNO is not enough

`read_paragraph_xml(source_path, paragraph_index, include_styles=True)` — opens the source file as a ZIP, parses `content.xml` + `styles.xml`, returns:
- raw `<text:p>` / `<text:h>` element XML
- paragraph style name
- styles dict: full style chain (parent walk) + every per-span text-style

Use this when you need:
- `fo:break-before="page"` on automatic paragraph styles (Word-emitted page-breaks UNO doesn't expose at paragraph level)
- `fo:letter-spacing` on text styles (Word's pre-baked justify spacing)
- `fo:keep-with-next` / `fo:keep-together` on heading styles
- `fo:hyphenate`, `fo:padding-*` and other `fo:*` attributes

The matching paragraph index aligns with `get_paragraphs` — both treat body paragraphs only (table-cell paragraphs are excluded).

### Locating

`find_all(text)`, `get_text_at(start, end)`, `get_text_content_live`, `get_selection`, `select_range(start, end)`.

---

## Write inventory — how to manipulate a document

### Text content

- `insert_text_live(text, position="end" | int)` — `\n` becomes a paragraph break. Position can be int offset or `"end"`.
- `format_text_live({font_name, font_size, bold, italic, underline, kerning, scale_width, color}, start, end)` — character formatting on a range. Range size 0 (start==end) is valid for setting CharHeight on an empty paragraph.
- `delete_range(start, end)`.
- `find_and_replace(find, replace, regex=False, ...)`.

### Paragraph properties

| Tool | What it sets |
|---|---|
| `set_paragraph_alignment(alignment, start, end)` | string `left`/`center`/`right`/`justify`/`stretch`/`block_line` OR raw int 0–5 (Word imports often use `STRETCH`=4 / `BLOCK_LINE`=5 — pass int to preserve) |
| `set_paragraph_indent(left_mm, right_mm, first_line_mm, start, end)` | All in mm |
| `set_paragraph_spacing(top_mm, bottom_mm, context_margin, start, end)` | `context_margin=True` collapses adjacent same-style spacings — read from source via `get_paragraphs` and replicate |
| `set_line_spacing(mode, value, start, end)` | `mode='proportional'` value=100 single, 150=1.5x; `'fix'/'minimum'/'leading'` value in mm |
| `set_paragraph_tabs(stops, start, end)` | List of `{position_mm, alignment, fill_char, decimal_char}`. Replaces all stops on the range. |
| `set_paragraph_breaks(break_type, page_desc_name, page_number_offset, start, end)` | `break_type` int 0–6 or name (NONE, PAGE_BEFORE=4, PAGE_AFTER=5, PAGE_BOTH=6, COLUMN_*). Use to replicate Word's `fo:break-before="page"` which UNO hides at paragraph level. |
| `set_paragraph_text_flow(widows, orphans, keep_together, split_paragraph, keep_with_next, start, end)` | Page-flow controls. Critical to prevent single words / single lines spilling onto next page. |
| `apply_paragraph_style(style_name, start, end)` | Assign existing style by name. The style must exist in target — clone it from source first if not. |
| `apply_numbering(rule_name, level, is_number, start, end, restart, start_value)` | Auto-numbering on a paragraph. |
| `set_text_color(color, start, end)`, `set_background_color(color, start, end)` | Hex `#RRGGBB`. |

### Tables

| Tool | What it does |
|---|---|
| `insert_table(rows, columns, position, name, column_widths_mm, table_width_mm, split, repeat_headline, header_row_count, keep_together)` | Insert at offset. Without `column_widths_mm` columns are equal; narrow content with `block_line` alignment will be visually shredded. |
| `write_table_cell_rich(table_name, cell, paragraphs)` | Paired with `read_table_rich`. Accepts a list of paragraph dicts (with runs). Internally **skips kerning / scale_width** on runs to avoid double-stretching pre-baked Word justify spacing. |
| `write_table_cell(table_name, cell, value)` | Plain text only. |
| `remove_table(table_name)` | Delete entire table. |

### Headers / footers / page numbering

| Tool | Use |
|---|---|
| `enable_header(enabled, page_style)` | Toggle header zone |
| `set_header(text, page_style)` | Plain text |
| `enable_footer(enabled, page_style)` | Toggle footer zone |
| `set_footer(text, page_style)` | Plain text |
| `set_footer_page_number(page_style, alignment="center", font_size)` | Centred PageNumber field in footer. Sets footer geometry to compact Word-export defaults. |
| `insert_text_frame(paragraph_index, width_mm, height_mm, page_number, hori_orient, vert_orient, hori_relation, vert_relation, x_mm, y_mm, back_transparent, remove_borders, text)` | Free-floating box. With `page_number=True` embeds a PageNumber field. For per-page numbering exactly like Word: `vert_orient="none"`, `vert_relation="page"`, `y_mm` slightly below body bottom, with one frame per page anchored to the paragraph that opens that page. |
| `get_header(page_style)`, `get_footer(page_style)` | Read |

### Page styles and global layout

| Tool | What it does |
|---|---|
| `clone_page_style(source_path, source_style, target_style)` | Copy ALL properties of a page-style from a source doc (which must be open) into the active doc. **Auto-creates target_style if missing**. Copies size, orientation, all 4 margins, header/footer enabled+height+margins, columns, footnote area, background. |
| `clone_paragraph_style(source_path, style_name)` | Same idea for paragraph styles. Includes `ParaKeepWithNext` / `KeepWithNext` (heading→table glue), break/keep/spacing, all character properties. |
| `clone_numbering_rule(source_path, rule_name)` | Copy auto-number rules; also auto-clones referenced character styles. |
| `set_page_style_props(page_style, ...)` | Set page-style properties directly without source doc. |
| `set_paragraph_style_props(style_name, ...)` | Set paragraph-style properties. |
| `set_page_margins(top_mm, bottom_mm, left_mm, right_mm, page_style)` | Set page margins on a style. Use as backup if `clone_page_style` doesn't carry per-page overrides. |

### Hyperlinks / bookmarks / comments / images / metadata

`add_hyperlink(start, end, url, target)`, `add_bookmark(name, start, end)`, `remove_bookmark(name)`, `add_comment(start, text, author, initials, end)`, `insert_image(image_path, paragraph_index, ...)`, `set_document_metadata`, `set_background_color`.

### Documents

| Tool | What it does |
|---|---|
| `create_document_live(doc_type="writer", visible=True)` | Create a new Writer/Calc/Impress/Draw window. **On macOS, when another doc is open, pass `visible=False` and call `show_window` after batch ops** — otherwise SolarMutex deadlock. |
| `open_document_live(path, readonly)` | Open existing file in a window. |
| `clone_document(source_path, target_path, target_format)` | Hidden-component file→file conversion. Bypasses macOS UI-thread deadlock. **The only reliable way to save/convert on macOS.** |
| `list_open_documents` | Enumerate windows. |
| `dispatch_uno_command(command)` | Send a `.uno:*` command (~66 whitelisted: bold, italic, alignment, page break, undo, copy/paste, navigation, etc.). |

### Power tool: `execute_batch`

```json
{
  "operations": [
    {"tool": "clone_page_style", "args": {...}},
    {"tool": "clone_paragraph_style", "args": {...}},
    {"tool": "insert_text_live", "args": {...}},
    {"tool": "apply_paragraph_style", "args": {...}}
  ],
  "stop_on_error": false,
  "lock_view": true,
  "auto_hide": "auto"
}
```

- One HTTP round-trip for many ops.
- `lock_view=true`: the UI freezes during the batch — no flicker, faster.
- `auto_hide="auto"` (default): if any "heavy" op (`clone_*`, `set_page_*`, `set_paragraph_style_props`, `apply_paragraph_style`, `write_table_cell_rich`) is present, the window is hidden during the batch, then restored. **This is the standard way to avoid macOS deadlock when modifying styles on a visible window.**
- `auto_hide="always"` / `"never"` to override.

When you have more than ~3 ops, **always use `execute_batch`**. Hundreds or thousands of ops in one batch is normal (a multi-page contract with full per-run formatting can take ~3-4 seconds for ~3000 ops).

### Undo / redo / window

`undo`, `redo`, `get_undo_history`, `hide_window`, `show_window`.

---

## Critical gotchas (read this before every session on macOS)

### 🔴 Save is broken on macOS through `libreoffice-live`

`doc.store()` / `doc.storeToURL()` / `.uno:Save` from a background HTTP-server thread block forever on macOS Sequoia (the AppKit run loop holds a barrier that UNO save needs).

**Workarounds:**

1. **For edits in user's open document**: after the last operation, tell the user:
   > Готово. Нажмите `Cmd+S` чтобы сохранить изменения в `<filename>`.
2. **For converting / writing a new file**: use `clone_document(source_path, target_path)` — it spawns a hidden component which can save. Or, if you need fresh content not derived from an existing file, drive the headless server `mcp__libreoffice__*` instead.
3. **NEVER** call `export_active_document` on macOS — it is disabled and would hang.

### 🔴 SolarMutex deadlock on heavy ops with visible window

If you create a new doc with `visible=True` and immediately run `clone_page_style` / `clone_paragraph_style` / `set_page_margins` etc., the call may hang forever. AppKit holds the SolarMutex during paint cycles, blocking the HTTP worker thread.

**Pattern:**
```python
create_document_live(doc_type="writer", visible=False)
execute_batch(operations=[clone_page_style, clone_paragraph_style, ...],
              auto_hide="auto")  # default
show_window()
```

`auto_hide="auto"` detects heavy ops and hides the window for the batch's duration even if it was visible.

### 🟡 UNO doesn't expose all `fo:*` attributes

For example:
- `fo:break-before="page"` on a paragraph **style** — UNO returns `BreakType=NONE` at paragraph level even though the page-break works. To replicate via `set_paragraph_breaks`, set `break_type=PAGE_BEFORE` (4) explicitly when the source paragraph has `page_desc_name` set.
- `fo:letter-spacing` on a span style — does map to `CharKerning` at portion level; this works.
- `fo:keep-with-next` on a paragraph style — clone via `clone_paragraph_style` (already includes the property).

When UNO read returns "everything looks normal" but the document layout differs, fall back to `read_paragraph_xml` to see the raw `fo:*` attributes.

### 🟡 Word→ODT pre-baked justify spacing

Word exports `justify` alignment by adding per-portion `CharKerning` and `CharScaleWidth` to the **space portions** between words. These values are valid only at the source's exact line width.

- In **body paragraphs**: replicate them via `format_text_live(kerning=, scale_width=)` — same line width = same look.
- In **table cells**: the column might be a different width in target. Combined with `ParaAdjust=BLOCK_LINE`, replicating the kerning gives **double-stretched** text (visible as broken-up letter spacing — "п о л н о с т ь ю"). `write_table_cell_rich` already skips these properties internally; do NOT re-apply them via `format_text_live` on cell ranges.

### 🟡 Table column widths

Source tables typically have non-uniform column widths. `insert_table` without `column_widths_mm` makes equal columns. With justify text, narrow columns shred content into single-word lines. **Always pass `column_widths_mm` from `read_table_rich`'s output.**

### 🟡 Style cloning collects only body paragraph styles

When iterating source paragraphs to gather `para_styles`, you'll miss styles used **inside table cells** (`Table Paragraph` etc.). When applied without cloning, target falls back to a default that's often bold-looking. **Collect styles from both body and cells**:

```python
para_styles = sorted(
    {p.get("style") for p in body_paragraphs if p.get("style")}
    | {p.get("style") for t in tables for row in t["cells"]
       for cell in row for p in cell.get("paragraphs", []) if p.get("style")}
)
```

### 🟡 Empty paragraphs have a font_size too

Word emits decorative empty paragraphs with `font-size=1pt..9pt` to compress vertical spacing. `get_paragraphs` returns `char_height` for these. Replicate via:

```python
if not paragraph_text and not runs and char_height:
    format_text_live(start=s, end=s, font_size=char_height)
```

Without this, target's empties become full 12pt lines and the document grows by extra pages.

### 🟡 Per-page numbering frames don't have stable anchor identity

`list_text_frames` reports frames anchored AT_PARAGRAPH but pyuno can't reliably tell **which** paragraph for empty-text anchors — `getString()` is empty for all and identity comparison fails across the proxy boundary. For Word-style per-page numbering: **N frames + N paragraphs with `page_desc_name`** is the universal pattern. Pair them by index of appearance.

### 🟡 Page-style switches are page-style hints, not breaks

Setting `page_desc_name="MP1"` alone often does NOT trigger a page break in modern LO builds — it only declares "from here on, use this layout". To force the break, set `break_type=PAGE_BEFORE` (4) on the same paragraph. (Word→ODT exporters typically emit BOTH `master-page-name` and `fo:break-before="page"` together.)

---

## Algorithms

### 1. Read whole document

```
1. get_document_summary  → know paragraph_count, table_count, ...
2. get_page_info  → page size, margins, page_style name
3. list_body_elements  → ordered sequence of {paragraph: idx, table: name, after_paragraph_index}
4. For each paragraph in body:
     get_paragraphs_with_runs(start=idx, count=1) → text, runs, alignment, indents, ...
   For each table:
     read_table_rich(table_name) → cells with paragraphs+runs, column_widths
5. list_text_frames  → page-number boxes etc.
6. list_hyperlinks, list_bookmarks, list_comments, list_images, list_sections (as needed)
7. read_paragraph_xml(paragraph_index)  → only when looking for a property UNO is hiding
```

### 2. Edit a live document (user has it open)

```
1. open_document_live(path) (or skip if already open and active)
2. Inspect with the Read algorithm above (focused on the target area).
3. Build operations as a list. Use execute_batch.
4. After the batch, tell the user: "Готово, нажмите Cmd+S".
```

### 3. Convert format (file on disk)

Single call:

```
clone_document(source_path="/path/foo.docx", target_path="/path/foo.odt")
```

(target_format auto-derived from extension; supported: docx, doc, odt, rtf, txt, html, xhtml, pdf, epub, ods, xlsx, xls, csv, pptx, ppt, odp).

Hidden component, no UI deadlock, works on macOS.

### 4. Replicate a document (build target by reading source and re-emitting via writes)

⚠️ This achieves *visually close* output, **NOT pixel-perfect**. For pixel-perfect output, use `clone_document`. Replicate is the right approach when the user wants programmatic editing or analysis-driven rebuild.

```
1. Open source readonly:
     open_document_live(path=source, readonly=True)

2. Snapshot:
     page_info = get_page_info()
     body = list_body_elements()
     paras = get_paragraphs_with_runs(count=N)
     flat  = get_paragraphs(count=N)         # for char_height / break_type / page_desc_name
     tables = [read_table_rich(name=t.name) for t in body if t.kind=='table']
     frames = list_text_frames()
     hyperlinks = list_hyperlinks()

3. Collect styles:
     para_styles = body_paragraph_styles ∪ table_cell_paragraph_styles
     numbering_rules = paragraphs with numbering_is_number
     page_styles_used = unique(p.page_desc_name for p in flat) ∪ {page_info.page_style}

4. Create target HIDDEN:
     create_document_live(doc_type="writer", visible=False)

5. Build ops list (one execute_batch):
   a. clone_page_style(source_style=page_info.page_style, target_style="Default Page Style")
   b. clone_page_style(source_style=ps, target_style=ps) for ps in page_styles_used (creates section masters)
   c. clone_paragraph_style(source_path, style_name=ps) for ps in para_styles
   d. clone_numbering_rule(source_path, rule_name=rn) for rn in numbering_rules
   e. set_page_margins(...) using page_info  # backup in case clone_page_style misses
   f. insert_text_live(text="\n".join(p.text for p in paras), position="end")
   g. For each paragraph i with text/style/format:
        if style: apply_paragraph_style(style, start, end)
        if numbered: apply_numbering(rule_name, level, start, end, restart=first_seen_for_rule, start_value=1)
        set_paragraph_indent(left, right, first_line, start, end)  # ALWAYS for numbered, else only when differs from style
        set_paragraph_spacing(top, bottom, context_margin, start, end)
        set_line_spacing(mode, value, start, end)
        set_paragraph_alignment(alignment_int, start, end)  # raw int — preserves stretch / block_line
        set_paragraph_tabs(tabs, start, end)
        set_paragraph_text_flow(widows, orphans, keep_together, split, keep_with_next, start, end)
        if char_height and empty: format_text_live(start, end=start, font_size=char_height)
        if break_type or (page_desc_name and i!=0): set_paragraph_breaks(break_type=4, start, end)  # PAGE_BEFORE
   h. For each run in each paragraph:
        format_text_live(start, end, font_name, font_size, bold, italic, underline, kerning, scale_width)
   i. For hyperlinks (extracted from runs with run.hyperlink):
        add_hyperlink(start, end, url)

6. Run execute_batch(operations=ops, stop_on_error=False, auto_hide="auto")

7. Insert tables AFTER paragraphs (sort by anchor offset DESC so each insertion doesn't shift later offsets):
   for offset, table in sorted(tables, by anchor_offset, reverse=True):
       insert_table(rows, columns, position=offset, name=table.name+"_copy",
                    column_widths_mm=table.column_widths_mm,
                    table_width_mm=table.table_width_mm,
                    split=table.split, repeat_headline=table.repeat_headline,
                    header_row_count=table.header_row_count,
                    keep_together=table.keep_together)
   for table in tables:
       cell_ops = [write_table_cell_rich(table_name, cell.name, cell.paragraphs)
                   for row in table.cells for cell in row if cell.paragraphs]
       execute_batch(operations=cell_ops)

8. Insert per-page TextFrames (from source frames, paired with page_desc_name paragraphs):
   page_starts = [p.index for p in flat if p.page_desc_name]
   for i, frame in enumerate(frames):
       paragraph_index = frame.anchor_para_index or page_starts[i]
       insert_text_frame(paragraph_index=paragraph_index,
                         width_mm=frame.size_mm.w, height_mm=frame.size_mm.h,
                         page_number=any(f.service=="PageNumber" for f in frame.fields),
                         hori_orient=..., vert_orient=...,
                         hori_relation=..., vert_relation=...,
                         x_mm=frame.HoriOrientPosition/100, y_mm=frame.VertOrientPosition/100)

9. show_window() — restore visibility for the user
10. Tell the user: "Готово. Нажмите Cmd+S чтобы сохранить как <name>."
```

### 5. Diff two documents

```
1. open_document_live(source1)
2. snap1 = get_paragraphs(count=N)
3. open_document_live(source2)
4. snap2 = get_paragraphs(count=N)
5. For each paragraph index, diff fields (alignment, indents, style, spacing, ...)
6. For tables: read_table_rich on both, diff cells + column_widths.
```

---

## Limitations the user should know about

Always be honest with the user about what is and isn't possible.

| Limitation | What to tell the user |
|---|---|
| **Save** through live MCP on macOS hangs | "После правок нажмите `Cmd+S` в окне LibreOffice — сам я сохранить не могу из-за macOS UI-thread deadlock. На Linux/Windows это работало бы автоматически." |
| **Pixel-perfect layout** is not achievable via batch writes | "Я могу воспроизвести содержимое и форматирование, но точное расположение строк по страницам зависит от LibreOffice rendering engine. Для побайтовой копии используй `clone_document`." |
| **`.pages` files** are not supported | "Apple Pages не открывается LibreOffice. Открой файл в Pages → File → Export To → Word (.docx), и я обработаю .docx." |
| **Layout simulator** doesn't exist | "Если в исходнике текст разрывается между страницами в специфичном месте, точно повторить это без полного rendering pipeline невозможно." |
| **Macros** don't run | "Макросы в .doc/.docx файлах отключены при импорте — я их не выполняю." |
| **Very large files** (>1000 paragraphs) — single `execute_batch` may timeout | "Для очень больших документов разбиваю на чанки по 200-500 параграфов." |
| **PDF content reading** | "PDF — для чтения используй другой инструмент. LibreOffice MCP читает только Writer-форматы. PDF умею только экспортировать." |
| **Track changes** | "Track changes / редакторские правки не поддерживаются — данные читаются как обычный текст." |
| **Footnotes / endnotes / fields beyond PageNumber** | "Базовые поля (PageNumber) поддерживаются. Для footnotes/citation fields нужно дополнительная работа — можем расширить мост по запросу." |

---

## Quick reference: typical user requests → tool sequence

| User says | Sequence |
|---|---|
| "Открой и покажи содержимое договора" | `open_document_live` → `get_document_summary` → `list_body_elements` → `get_paragraphs_with_runs` |
| "Замени все 'Иванов' на 'Петров'" | `find_and_replace(find="Иванов", replace="Петров")` → tell user Cmd+S |
| "Сделай заголовок жирным" | `find_all("заголовок текст")` → `format_text_live(start, end, bold=True)` → Cmd+S |
| "Конвертируй foo.docx в PDF" | `clone_document("foo.docx", "foo.pdf")` |
| "Создай отчёт по шаблону" | `create_document_live(visible=False)` → `clone_page_style` etc. → `insert_text_live` → `format_text_live` → batch → `show_window` → Cmd+S |
| "Сравни два договора" | open both → `get_paragraphs` on each → diff fields |
| "Дай таблицу из второго раздела как CSV" | `read_table_rich(table_index=1)` → emit CSV from `cells` |
| "Обнови все даты на 2026" | `find_all(regex="20[0-9]{2}")` → for each match `format_text_live` (or `find_and_replace` with regex) |
| "Скопируй вёрстку из old.odt в new.odt" | open old → snapshot → open new → `clone_page_style(source_path=old)` etc. |

---

## Final reminders

1. **Always tell the user to press Cmd+S after edits on macOS.** If you forget, the user's changes are lost.
2. **`execute_batch` is your friend.** Don't make 50 sequential HTTP calls.
3. **`clone_document` is the only safe save on macOS.** When the user wants a saved file, use it.
4. **`read_paragraph_xml` is the escape hatch** when UNO API is silent on a property.
5. **Replicate is *not* clone.** Set expectations honestly: visually close, not byte-identical.
6. **`.pages` is not supported.** Don't try; ask the user to convert.
7. **On unfamiliar documents, always start with `get_document_summary` + `list_body_elements`** to understand structure before mass operations.
