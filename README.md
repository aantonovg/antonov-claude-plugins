# claude-marketplace

Публичный маркетплейс плагинов для [Claude Code](https://claude.com/claude-code).

## Установка маркетплейса

```bash
/plugin marketplace add aantonovg/claude-marketplace
```

После добавления плагины из этого репозитория будут доступны в `/plugin` для установки.

## Доступные плагины

| Плагин | Версия | Описание |
|--------|--------|----------|
| [sensortower](./sensortower) | 1.0.0 | Sensor Tower MCP — 85 инструментов app-intelligence: рейтинги, метаданные, выручка, ключевые слова, реклама. |

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
