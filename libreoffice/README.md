# libreoffice

Plugin для Claude Code, дающий **живое редактирование документов LibreOffice Writer** прямо из чата. После установки агент видит **55+ MCP-инструментов** и работает с открытым в LibreOffice документом — изменения появляются в окне сразу, без перезагрузки файла.

Покрывает: текст и параграфы, стили, цвет/фон/выравнивание/отступы/межстрочный интервал, find&replace, programmatic select_range, закладки/гиперссылки/комментарии, таблицы (read+write по ячейкам), headers/footers, метаданные документа, undo/redo, dispatch к ~66 безопасным `.uno:` командам.

## Что внутри

```
libreoffice/
├── .claude-plugin/plugin.json     # манифест Claude Code
├── .mcp.json                      # запускает live_bridge.py через uv
├── live_bridge.py                 # stdio MCP-мост → HTTP localhost:8765
├── pyproject.toml                 # зависимости моста (mcp, httpx)
├── extension/
│   ├── libreoffice-mcp-extension.oxt   # готовое расширение (поставить в LibreOffice)
│   └── source/                          # исходники расширения для пересборки
└── scripts/
    ├── install.sh                 # автоустановка: re-sign + unopkg add
    └── uninstall.sh
```

Архитектура: внутри LibreOffice крутится Python-расширение, которое поднимает HTTP-сервер на `localhost:8765` и делает все действия через UNO API. Стартует **автоматически** при запуске LibreOffice. `live_bridge.py` — это тонкий stdio↔HTTP мост, который Claude Code запускает как обычный stdio MCP-сервер.

## Prerequisites

| Компонент | Минимальная версия | Зачем |
|---|---|---|
| LibreOffice | 24.2+ | основной хост — работаем с открытым документом через UNO |
| Python | 3.10+ | для `live_bridge.py` (mcp + httpx) |
| `uv` | свежий | запускает мост в изолированной venv (без `pip install` на хост) |
| `unopkg` | идёт с LibreOffice | устанавливает расширение |

Установка prerequisites:

```bash
# macOS
brew install --cask libreoffice
brew install uv

# Debian/Ubuntu
sudo apt install libreoffice
curl -LsSf https://astral.sh/uv/install.sh | sh

# Windows
winget install TheDocumentFoundation.LibreOffice
winget install astral-sh.uv
```

## Установка в Claude Code

```
/plugin marketplace add aantonovg/antonov-claude-plugins
/plugin install libreoffice@antonov-claude-plugins
```

После установки плагина — **один раз** запустить `install.sh`, который ставит расширение в LibreOffice и (на macOS) делает ad-hoc re-sign:

```bash
~/.claude/plugins/cache/antonov-claude-plugins/libreoffice/*/scripts/install.sh
```

> Реальный путь содержит версию плагина (например `.../libreoffice/1.0.0/scripts/install.sh`). Если glob не сработает — найди вручную:
> ```bash
> find ~/.claude -name "install.sh" -path "*libreoffice*" 2>/dev/null
> ```

После этого:
1. Открой LibreOffice (расширение автоматически стартует HTTP API на :8765).
2. Перезапусти Claude Code.
3. Проверь: `claude mcp list` должен показать **`libreoffice-live: ✓ Connected`**.

## Установка в Codex

Codex CLI умеет подключать кастомные marketplace-источники из GitHub. Добавь этот репозиторий как marketplace:

```bash
codex plugin marketplace add aantonovg/antonov-claude-plugins
```

После этого плагин `libreoffice` должен появиться в интерфейсе Codex среди доступных плагинов из marketplace `antonov-claude-plugins`.

Если нужно обновить marketplace после нового push:

```bash
codex plugin marketplace upgrade antonov-claude-plugins
```

После установки плагина — **один раз** запустить `install.sh`, который ставит расширение в LibreOffice и (на macOS) делает ad-hoc re-sign:

```bash
~/.codex/plugins/cache/antonov-claude-plugins/libreoffice/*/scripts/install.sh
```

Если glob не сработает — найди реальный путь:

```bash
find ~/.codex -name "install.sh" -path "*libreoffice*" 2>/dev/null
```

После этого:
1. Открой LibreOffice (расширение автоматически стартует HTTP API на :8765).
2. Перезапусти Codex.
3. Проверь, что MCP-сервер `libreoffice-live` подключился.

### Ручное подключение в Codex CLI

Если UI/marketplace в твоей версии Codex не показывает кастомный плагин, можно подключить MCP-сервер напрямую через клон репозитория и `~/.codex/config.toml`.

1. Клонировать репозиторий:
   ```bash
   git clone https://github.com/aantonovg/antonov-claude-plugins ~/.codex-plugins/antonov-claude-plugins
   ```

2. Поставить prerequisites (см. выше) и расширение в LibreOffice:
   ```bash
   ~/.codex-plugins/antonov-claude-plugins/libreoffice/scripts/install.sh
   # (в Codex путь без версии — клон лежит как обычный git checkout)
   ```

3. Добавить MCP-сервер в `~/.codex/config.toml`:
   ```toml
   [mcp_servers.libreoffice-live]
   command = "uv"
   args = [
     "run",
     "--project", "/Users/YOURNAME/.codex-plugins/antonov-claude-plugins/libreoffice",
     "python",
     "/Users/YOURNAME/.codex-plugins/antonov-claude-plugins/libreoffice/live_bridge.py",
   ]
   ```
   *(Замени `/Users/YOURNAME/...` на реальный абсолютный путь — ручной MCP-конфиг Codex не подставляет `${CLAUDE_PLUGIN_ROOT}`.)*

4. Открой LibreOffice, перезапусти Codex, проверь `codex mcp list` (или эквивалент в твоей версии — название команды менялось).

## Что делает `install.sh`

1. Проверяет наличие `soffice`/`libreoffice`, `unopkg`, `uv`.
2. **На macOS** — делает `sudo codesign --force --deep --sign - /Applications/LibreOffice.app`. Это снимает Launch Constraints в TDF-подписи, которые на macOS Sequoia убивают `LibreOfficePython.framework` через `SIGKILL CODESIGNING`. Без этого расширение не сможет стартовать HTTP-сервер. Подпись становится ad-hoc; LibreOffice работает идентично. **Повторять после каждого `brew upgrade --cask libreoffice`**.
3. Закрывает LibreOffice (он не должен быть запущен во время `unopkg add`, иначе UNO-pipe не отдаст активацию).
4. `unopkg remove` старой версии (если была) и `unopkg add` свежей.

Скрипт идемпотентен — можно перезапускать.

## Доступные MCP-инструменты (65+)

| Категория | Tools |
|---|---|
| Создание / открытие / завершение | `create_document_live`, `open_document_live`, `list_open_documents`, `list_recent_documents`, `open_recent_document`, `shutdown_application` (graceful, без recovery dialog) |
| Чтение содержимого | `get_text_content_live`, `get_text_at`, `get_paragraphs`, `get_paragraphs_with_runs`, `get_outline`, `get_paragraph_format_at`, `get_character_format`, `get_selection`, `get_document_info_live`, `get_document_summary`, `get_document_metadata`, `get_page_info` (margins + header/footer + columns + автодетект `PageDescName` первого параграфа — для Word-import master-pages типа `MP0`) |
| Стили — чтение | `list_paragraph_styles`, `list_character_styles`, `list_numbering_styles`, `get_paragraph_style_def` (полный snapshot: font/bold/italic/underline/color/char_word_mode/alignment/indents/spacing/line_spacing/tab_stops/context_margin/outline_level/parent/follow) |
| Стили — запись | `apply_paragraph_style`, `set_paragraph_style_props` (симметрично get_paragraph_style_def), `set_page_style_props` (симметрично get_page_info), `set_page_margins`, `apply_numbering` |
| Стили — клонирование | `clone_paragraph_style`, `clone_page_style`, `clone_numbering_rule` (быстрая массовая репликация из открытого source-документа) |
| Поиск | `find_all`, `find_and_replace` |
| Запись текста | `insert_text_live` (с поддержкой `\n` как paragraph break), `delete_range`, `select_range` |
| Форматирование текста | `format_text_live` (bold/italic/underline/font_size/font_name), `set_text_color`, `set_background_color` |
| Параграфы | `set_paragraph_alignment`, `set_paragraph_indent`, `set_paragraph_spacing` (с context_margin), `set_paragraph_tabs`, `set_line_spacing` |
| Headers/Footers | `enable_header`/`enable_footer`, `set_header`/`set_footer`, `get_header`/`get_footer` (плюс через `set_page_style_props`) |
| Закладки | `list_bookmarks`, `add_bookmark`, `remove_bookmark` |
| Гиперссылки | `list_hyperlinks`, `add_hyperlink` |
| Комментарии | `list_comments`, `add_comment` |
| Изображения / shapes | `list_images`, `insert_image` |
| Таблицы | `get_tables_info`, `read_table_cells`, `write_table_cell`, `insert_table`, `remove_table` |
| Секции | `list_sections` |
| Метаданные | `set_document_metadata` |
| Undo / Redo | `undo`, `redo`, `get_undo_history` |
| Конвертация форматов | `clone_document` (file-on-disk, headless: ODT/DOCX/PDF/HTML/RTF/XLSX/...) |
| Батчинг / view | `execute_batch` (массив операций → один HTTP-вызов; auto lock/unlock view), `lock_view`, `unlock_view`, `show_window` |
| Низкоуровневое | `dispatch_uno_command` (whitelist из ~70 безопасных команд, включая `.uno:ControlCodes`/`.uno:SpellOnline` view toggles) |

## Известные ограничения (macOS)

- **Save / Export через MCP не работает.** `doc.store()`, `doc.storeToURL()` и `.uno:Save` через dispatch на macOS Sequoia блокируются на UI-thread из background-thread HTTP-сервера, причём блок может полностью повесить worker до перезапуска LibreOffice. **Workflow: агент делает правки → пользователь жмёт Cmd+S** в окне LibreOffice (это работает мгновенно, потому что выполняется в UI-thread). Tools `save_document_live`, `export_document_live` и `export_active_document` сознательно удалены из набора (последний — в v1.0.1, после реального инцидента с зависанием). Для конвертации формата используй `clone_document` — он работает через hidden component и не касается UI-thread.
- **`dispatch_uno_command` ограничен whitelist'ом** ~66 безопасных команд (formatting, navigation, selection, edit). Любая другая команда (включая Save/Export/Print/Open/Close/RunMacro) возвращает `error` без вызова UNO API — гарантия, что сервер не зависнет.
- **`list_recent_documents`** в текущей версии возвращает индексы вместо URL'ов (структура `Histories/PickList` в LO 26 хитрая). `open_document_live` по абсолютному пути работает безупречно — используй его.

На Linux/Windows ограничения по save отсутствуют (это специфика macOS AppKit). При необходимости вернуть save для не-macOS — раскомментировать `_removed_save_document` в `extension/source/pythonpath/uno_bridge.py` и зарегистрировать tool обратно в `mcp_server.py`.

## Логи и отладка

- HTTP API: `curl http://localhost:8765/health` → `{"status":"healthy"}`
- Список инструментов: `curl http://localhost:8765/tools | jq '.count'`
- Логи расширения: `/tmp/lo_mcp.log` (LibreOffice GUI редиректит stderr в /dev/null, поэтому логи в файле)
- Состояние установки: `unopkg list | grep mcp`

## Снятие

```bash
~/.../libreoffice/scripts/uninstall.sh
# в Claude Code:
/plugin uninstall libreoffice
```

## Лицензия

MIT (наследуется от исходного [`patrup/mcp-libre`](https://github.com/patrup/mcp-libre), на котором основано расширение). Этот плагин содержит существенно переработанные версии исходных файлов: исправлено ~15 багов автора (relative imports, hardcoded paths, неправильная упаковка `.oxt`, race condition в HTTP-сервере, UI-thread deadlock в `create_document`, обработка локализованных имён page styles), плюс добавлено ~40 новых tools (read/write inspect, headers/footers, undo/redo, dispatch).
