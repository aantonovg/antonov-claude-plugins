# antonov-claude-plugins

Публичный маркетплейс плагинов для [Claude Code](https://claude.com/claude-code).

## Установка маркетплейса

```
/plugin marketplace add aantonovg/antonov-claude-plugins
```

После добавления плагины из этого репозитория будут доступны в `/plugin` для установки.

## Доступные плагины

| Плагин | Версия | Описание |
|--------|--------|----------|
| [sensortower](./sensortower) | 1.0.1 | Sensor Tower MCP — 85 инструментов app-intelligence: рейтинги, метаданные, выручка, ключевые слова, реклама. |
| [libreoffice](./libreoffice) | 1.0.0 | Live-редактирование LibreOffice Writer из Claude Code: 55+ инструментов (параграфы, стили, find&replace, гиперссылки, комментарии, таблицы, headers/footers, undo/redo). |

## Установка плагина

После того как маркетплейс добавлен:

```
/plugin install sensortower@antonov-claude-plugins
/plugin install libreoffice@antonov-claude-plugins
```

Альтернативно — через интерактивное меню `/plugin` → Discover → выбрать плагин.

Все плагины используют [`uv`](https://docs.astral.sh/uv/) — поставь его сначала:

- **macOS / Linux:** `curl -LsSf https://astral.sh/uv/install.sh | sh`
- **Windows (PowerShell):** `irm https://astral.sh/uv/install.ps1 | iex`
- или через пакетный менеджер: `brew install uv` / `winget install astral-sh.uv` / `scoop install uv`

`libreoffice` дополнительно требует [LibreOffice 24.2+](https://www.libreoffice.org/download/) и однократного запуска инсталлера расширения — см. [libreoffice/README.md](./libreoffice/README.md).

## Настройка API-ключа (sensortower)

Плагин `sensortower` читает `SENSORTOWER_API_KEY` из окружения **процесса Claude Code на момент его старта** — не из shell внутри сессии. Переменную нужно задать **до** запуска Claude Code и затем перезапустить его.

### macOS / Linux

Добавить в `~/.zshrc` (или `~/.bashrc`):

```bash
export SENSORTOWER_API_KEY="..."
```

Для GUI-приложения Claude Code на macOS переменные из `~/.zshrc` могут не подхватиться — тогда:

```bash
launchctl setenv SENSORTOWER_API_KEY "..."
```

(или прописать в `~/Library/LaunchAgents/<name>.plist` для постоянного эффекта).

### Windows

PowerShell, постоянная user-переменная (рекомендуется — без лимита длины):

```powershell
[Environment]::SetEnvironmentVariable("SENSORTOWER_API_KEY", "...", "User")
```

Альтернатива через `setx` (лимит значения 1024 символа):

```cmd
setx SENSORTOWER_API_KEY "..."
```

Через GUI: «Изменение переменных среды текущего пользователя» → New → имя `SENSORTOWER_API_KEY`, значение — ваш ключ.

После любой из команд **перезапустите терминал и Claude Code** — в уже работающем процессе переменная не появится.

### Проверка

Если ключ не задан, `server.py` стартует, но в stderr напишет `⚠️ SENSORTOWER_API_KEY is empty`, и API вернёт 401 на первом запросе.

## Структура

```
.
├── .claude-plugin/
│   └── marketplace.json     # манифест маркетплейса
└── <plugin-name>/
    ├── .claude-plugin/
    │   └── plugin.json      # манифест плагина
    ├── .mcp.json            # (опц.) MCP-серверы плагина
    └── skills/              # (опц.) skills плагина
```

## Лицензия

MIT
