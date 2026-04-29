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
~/.claude/plugins/cache/aantonovg/antonov-claude-plugins/libreoffice/scripts/install.sh
```

> Точный путь зависит от расположения plugin cache. Если он другой — найди:
> ```bash
> find ~/.claude -name "install.sh" -path "*libreoffice*" 2>/dev/null
> ```

После этого:
1. Открой LibreOffice (расширение автоматически стартует HTTP API на :8765).
2. Перезапусти Claude Code.
3. Проверь: `claude mcp list` должен показать **`libreoffice-live: ✓ Connected`**.

## Установка в Codex CLI

В отличие от Claude Code, Codex не имеет marketplace для плагинов. Подключение делается **руками** через клон репозитория и редактирование конфига.

1. Клонировать репозиторий:
   ```bash
   git clone https://github.com/aantonovg/antonov-claude-plugins ~/.codex-plugins/antonov-claude-plugins
   ```

2. Поставить prerequisites (см. выше) и расширение в LibreOffice:
   ```bash
   ~/.codex-plugins/antonov-claude-plugins/libreoffice/scripts/install.sh
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
   *(Замени `/Users/YOURNAME/...` на реальный абсолютный путь — Codex не подставляет `${CLAUDE_PLUGIN_ROOT}`.)*

4. Открой LibreOffice, перезапусти Codex, проверь `codex mcp list` (или эквивалент в твоей версии — название команды менялось).

## Что делает `install.sh`

1. Проверяет наличие `soffice`/`libreoffice`, `unopkg`, `uv`.
2. **На macOS** — делает `sudo codesign --force --deep --sign - /Applications/LibreOffice.app`. Это снимает Launch Constraints в TDF-подписи, которые на macOS Sequoia убивают `LibreOfficePython.framework` через `SIGKILL CODESIGNING`. Без этого расширение не сможет стартовать HTTP-сервер. Подпись становится ad-hoc; LibreOffice работает идентично. **Повторять после каждого `brew upgrade --cask libreoffice`**.
3. Закрывает LibreOffice (он не должен быть запущен во время `unopkg add`, иначе UNO-pipe не отдаст активацию).
4. `unopkg remove` старой версии (если была) и `unopkg add` свежей.

Скрипт идемпотентен — можно перезапускать.

## Доступные MCP-инструменты (57)

| Категория | Tools |
|---|---|
| Создание / открытие | `create_document_live`, `open_document_live`, `list_open_documents`, `list_recent_documents`, `open_recent_document` |
| Чтение содержимого | `get_text_content_live`, `get_text_at`, `get_paragraphs`, `get_paragraphs_with_runs`, `get_outline`, `get_paragraph_format_at`, `get_character_format`, `get_selection`, `get_document_info_live`, `get_document_summary`, `get_document_metadata`, `get_page_info` |
| Стили | `list_paragraph_styles`, `list_character_styles`, `apply_paragraph_style` |
| Поиск | `find_all`, `find_and_replace` |
| Запись текста | `insert_text_live` (с поддержкой `\n` как paragraph break), `delete_range`, `select_range` |
| Форматирование текста | `format_text_live` (bold/italic/underline/font_size/font_name), `set_text_color`, `set_background_color` |
| Параграфы | `set_paragraph_alignment`, `set_paragraph_indent`, `set_line_spacing` |
| Headers/Footers | `enable_header`/`enable_footer`, `set_header`/`set_footer`, `get_header`/`get_footer` |
| Закладки | `list_bookmarks`, `add_bookmark`, `remove_bookmark` |
| Гиперссылки | `list_hyperlinks`, `add_hyperlink` |
| Комментарии | `list_comments`, `add_comment` |
| Изображения / shapes | `list_images`, `insert_image` |
| Таблицы | `get_tables_info`, `read_table_cells`, `write_table_cell`, `insert_table`, `remove_table` |
| Секции | `list_sections` |
| Метаданные | `set_document_metadata` |
| Undo / Redo | `undo`, `redo`, `get_undo_history` |
| Низкоуровневое | `dispatch_uno_command` (whitelist из ~66 безопасных команд) |

## Известные ограничения (macOS)

- **Save через MCP не работает.** `doc.store()`, `doc.storeToURL()` и `.uno:Save` через dispatch на macOS Sequoia блокируются на UI-thread из background-thread HTTP-сервера. **Workflow: агент делает правки → пользователь жмёт Cmd+S** в окне LibreOffice (это работает мгновенно, потому что выполняется в UI-thread). Tools `save_document_live` и `export_document_live` сознательно удалены из набора.
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
