# antonov-claude-plugins

Публичный маркетплейс плагинов для [Claude Code](https://claude.com/claude-code) и Codex.

## Установка маркетплейса в Claude Code

```
/plugin marketplace add aantonovg/antonov-claude-plugins
```

После добавления плагины из этого репозитория будут доступны в `/plugin` для установки.

## Доступные плагины

| Плагин | Версия | Описание |
|--------|--------|----------|
| [sensortower](./sensortower) | 1.0.1 | Sensor Tower MCP — 85 инструментов app-intelligence: рейтинги, метаданные, выручка, ключевые слова, реклама. |
| [libreoffice](./libreoffice) | 1.0.8 | Live-редактирование LibreOffice Writer из Claude Code и Codex: MCP-инструменты для параграфов, стилей, таблиц, headers/footers, layout inspection и конвертации. |

## Установка плагина

После того как маркетплейс добавлен:

```
/plugin install sensortower@antonov-claude-plugins
/plugin install libreoffice@antonov-claude-plugins
```

Альтернативно — через интерактивное меню `/plugin` → Discover → выбрать плагин.

## Установка маркетплейса в Codex

Codex CLI поддерживает кастомные marketplace-источники из GitHub:

```bash
codex plugin marketplace add aantonovg/antonov-claude-plugins
```

После этого marketplace появится в `~/.codex/config.toml` как секция `[marketplaces.antonov-claude-plugins]`, а плагины из `.agents/plugins/marketplace.json` будут доступны в интерфейсе Codex.

Если UI не показывает кастомные плагины сразу, обнови marketplace и перезапусти Codex:

```bash
codex plugin marketplace upgrade antonov-claude-plugins
```

Для `libreoffice` после установки плагина нужно один раз поставить расширение LibreOffice:

```bash
~/.codex/plugins/cache/antonov-claude-plugins/libreoffice/*/scripts/install.sh
```

Если путь отличается, найди installed copy:

```bash
find ~/.codex -name "install.sh" -path "*libreoffice*" 2>/dev/null
```

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
├── .agents/
│   └── plugins/
│       └── marketplace.json # манифест маркетплейса Codex
├── .claude-plugin/
│   └── marketplace.json     # манифест маркетплейса
└── <plugin-name>/
    ├── .codex-plugin/
    │   └── plugin.json      # манифест плагина Codex
    ├── .claude-plugin/
    │   └── plugin.json      # манифест плагина
    ├── .mcp.json            # (опц.) MCP-серверы плагина
    └── skills/              # (опц.) skills плагина
```

## Лицензия

MIT
