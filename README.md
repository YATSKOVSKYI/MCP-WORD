# MCP-сервер для Microsoft Word (docx)

**MCP Word Server** — это простой CLI-сервер на Node.js, реализующий Model Context Protocol (MCP) для создания и управления документами Microsoft Word (формат `.docx`). С помощью JSON-команд, отправленных через стандартный ввод/вывод, можно:

- Создавать и сохранять документы.
- Добавлять и форматировать текст (параграфы, заголовки, списки).
- Работать с таблицами (создание, строки, стили, объединение ячеек).
- Вставлять изображения и подписи.
- Управлять колонтитулами и нумерацией страниц.
- Настраивать разметку (поля, ориентацию, разрывы).
- Задавать метаданные документа (тема, автор, заголовок).

## 📦 Установка

1. Клонируйте репозиторий:
   ```bash
   git clone https://github.com/<ваш-аккаунт>/mcp-word-server.git
   cd mcp-word-server
   ```
2. Установите зависимости:
   ```bash
   npm install
   ```

## 🚀 Запуск

```bash
node server.js
```

При запуске сервер готов принимать JSON-команды через `stdin` и ответит JSON-объектом в `stdout`.

## 🔧 Конфигурация MCP-клиента

Добавьте в ваш файл `cline_mcp_settings.json`:

```json
{
  "mcpServers": {
    "word-server": {
      "command": "node",
      "args": ["server.js"],
      "env": {},
      "disabled": false,
      "autoApprove": [
        "create_document",
        "add_paragraph",
        "add_heading",
        "add_bullet_list",
        "add_numbered_list",
        "set_alignment",
        "set_font_size",
        "set_font_family",
        "set_spacing",
        "add_table",
        "add_table_row",
        "set_table_style",
        "merge_cells",
        "insert_image",
        "insert_image_with_caption",
        "add_header",
        "add_footer",
        "add_page_numbering",
        "add_page_break",
        "set_margins",
        "set_orientation",
        "set_title",
        "set_author",
        "set_subject",
        "save_document"
      ]
    }
  }
}
```

## 📝 Список команд

Все команды принимают объект `{ command: string, params?: any }` и возвращают `{ command, status, result? | error? }`.

### 1. Создание и сохранение
- **create_document** — создаёт новый документ. Дополнительно можно передать `{ title, author, subject }`.
- **save_document** — сохраняет документ в файл. Параметр `{ path: string }`.

### 2. Работа с текстом
- **add_paragraph** `{ text: string }`
- **add_heading** `{ text: string, level?: number }`
- **add_bullet_list** `{ items: string[] }`
- **add_numbered_list** `{ items: string[] }`
- **set_alignment** `{ alignment: "left"|"center"|"right"|"justify" }`
- **set_font_size** `{ size: number }` (в пунктах)
- **set_font_family** `{ name: string }`
- **set_spacing** `{ before?: number, after?: number, line?: number }` или `number` (межстрочный)

### 3. Таблицы
- **add_table** `{ data?: string[][] }`
- **add_table_row** `{ cells: string[] }`
- **set_table_style** `{ style: string }`
- **merge_cells** `{ startRow, startCol, endRow, endCol }`

### 4. Изображения
- **insert_image** `{ path: string, width?: number, height?: number }`
- **insert_image_with_caption** `{ path: string, caption: string, width?: number, height?: number }`

### 5. Колонтитулы и нумерация
- **add_header** `{ text: string }`
- **add_footer** `{ text: string }`
- **add_page_numbering** `{} // добавляет "Стр. X из Y" в footer`

### 6. Разметка
- **add_page_break** `{}`
- **set_margins** `{ top:number, bottom:number, left:number, right:number }` (см или дюймы)
- **set_orientation** `{ orientation: "portrait"|"landscape" }`

### 7. Метаданные
- **set_title** `{ title: string }`
- **set_author** `{ author: string }`
- **set_subject** `{ subject: string }`

## 📋 Примеры использования

Пример взаимодействия через `stdin`:
```json
{ "command": "create_document", "params": { "title": "Отчет" } }
{ "command": "add_heading", "params": { "text": "Заголовок отчета", "level": 1 } }
{ "command": "add_paragraph", "params": { "text": "Основной текст абзаца." } }
{ "command": "add_table", "params": { "data": [["Параметр","Значение"],["Температура","25°C"]] } }
{ "command": "insert_image_with_caption", "params": { "path": "diagram.jpg", "caption": "Рисунок 1. Схема." } }
{ "command": "add_page_numbering" }
{ "command": "save_document", "params": { "path": "report.docx" } }
```

## 🛠️ Разработка и вклад

- **Node.js** v14+ и **npm**
- Основная зависимость: [docx](https://www.npmjs.com/package/docx)

Пожалуйста, создавайте issues и pull requests для улучшений и исправлений.

## 📄 Лицензия

MIT © 2025

