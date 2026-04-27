# sensortower

Sensor Tower MCP-сервер на FastMCP — даёт Claude Code доступ к 85 инструментам app-intelligence:

- **Рейтинги и чарты** — позиции в App Store / Google Play, топ free/paid/grossing.
- **Метаданные** — описания, иконки, скриншоты, SDK, IAP.
- **Performance** — оценки загрузок, выручки, активных пользователей (DAU/WAU/MAU).
- **ASO** — ключевые слова, traffic score, поисковые подсказки, тренды.
- **Apple Search Ads** — SOV, конкуренты, история размещений.
- **Реклама** — топ креативов, паблишеров, сети, impressions.
- **Featured** — фичеринг в сторе и его влияние на загрузки.
- **Reviews** — отзывы, рейтинги, история обновлений.
- **Usage / Retention / Churn** — панельные данные, демография, когорты.
- **Custom fields, App Store Connect, Games breakdown, Reference**.

Полный каталог методов с параметрами и примерами — в [`skills/sensortower/SKILL.md`](./skills/sensortower/SKILL.md).

## Требования

- [`uv`](https://docs.astral.sh/uv/) для запуска `server.py`.
- API-ключ Sensor Tower в переменной окружения `SENSORTOWER_API_KEY`.

## Установка

Через маркетплейс `claude-marketplace`:

```
/plugin marketplace add aantonovg/claude-marketplace
/plugin install sensortower@antonov-claude-plugins
```

## Переменная окружения

Плагин читает `SENSORTOWER_API_KEY` из окружения **процесса Claude Code на момент старта плагина**, а не из shell внутри сессии. Поэтому переменную нужно задать **до** запуска Claude Code.

**macOS / Linux** — добавить в `~/.zshrc` или `~/.bashrc`:
```bash
export SENSORTOWER_API_KEY="..."
```
Для GUI-приложения Claude Code на macOS — через `launchctl setenv` или `~/Library/LaunchAgents`.

**Windows (PowerShell)** — постоянная user-переменная:
```powershell
[Environment]::SetEnvironmentVariable("SENSORTOWER_API_KEY", "...", "User")
```
или через `setx SENSORTOWER_API_KEY "..."` (есть лимит 1024 символа). После — **перезапустить терминал и Claude Code**, в текущем процессе переменная не появится.

**Windows (GUI)** — «Изменение переменных среды текущего пользователя» → New → имя `SENSORTOWER_API_KEY`.

Если ключ не задан, `server.py` стартует, но в stderr будет `⚠️ SENSORTOWER_API_KEY is empty`, и API вернёт 401 на первом же запросе.

## Состав плагина

```
sensortower/
├── .claude-plugin/plugin.json   # манифест
├── .mcp.json                    # MCP-сервер (uv run server.py)
├── server.py                    # FastMCP-сервер
├── sensortower_openapi.yaml     # OpenAPI-спецификация (источник методов)
└── skills/sensortower/SKILL.md  # каталог 85 методов с описаниями
```

## Лицензия

MIT
