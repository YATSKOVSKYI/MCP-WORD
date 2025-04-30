#!/usr/bin/env node

// Импорт библиотеки docx и необходимых классов
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
    Table, TableRow, TableCell, ImageRun, Header, Footer,
    PageOrientation, PageBreak, PageNumber } = require("docx");
const fs = require("fs");
const readline = require("readline");

// Глобальные переменные состояния документа
let doc = null;
let docChildren = [];             // Контейнер для всех блоков документа (Paragraph, Table, и т.д.)
let currentTable = null;          // Ссылка на текущую таблицу (если идет заполнение таблицы)
let lastParagraph = null;         // Ссылка на последний добавленный Paragraph (для форматирования)
let headerContent = [];           // Параграфы верхнего колонтитула
let footerContent = [];           // Параграфы нижнего колонтитула
let pageOrientation = PageOrientation.PORTRAIT;  // Ориентация страницы (по умолчанию книжная)
let pageMargins = null;           // Поля страницы (если null, будут использованы по умолчанию)
let docProps = { title: "", subject: "", creator: "" };  // Метаданные документа
let numberingUsed = false;
let bulletListUsed = false;

// Функция для формирования структуры ответа
function makeResponse(success, command, message) {
if (success) {
    return { command: command, status: "success", result: message || null };
} else {
    return { command: command, status: "error", error: message };
}
}

// Основная функция обработки команд
async function handleCommand(msg) {
const cmd = msg.command;
const params = msg.params || {};
switch (cmd) {
    case "create_document": {
        // Инициализация нового документа
        docChildren = [];
        currentTable = null;
        lastParagraph = null;
        headerContent = [];
        footerContent = [];
        pageOrientation = PageOrientation.PORTRAIT;
        pageMargins = null;
        // Обновляем метаданные, если переданы
        if (params.title) docProps.title = params.title;
        if (params.author) docProps.creator = params.author;
        if (params.subject) docProps.subject = params.subject;
        // Создаем новый документ (пока без секций, секция будет добавлена при сохранении)
        doc = new Document();
        return makeResponse(true, cmd, "Document created");
    }

    case "save_document": {
        if (!doc) {
            return makeResponse(false, cmd, "No document to save. Use create_document first.");
        }
        // Путь для сохранения обязателен
        const filePath = params.path || params.file || params.filename;
        if (!filePath) {
            return makeResponse(false, cmd, "Missing file path for save_document");
        }
        try {
            // Формируем секцию документа со всеми накопленными элементами
            const sectionProps = {};
            // Устанавливаем ориентацию, если не по умолчанию
            if (pageOrientation) {
                sectionProps.page = { size: { orientation: pageOrientation } };
            }
            // Устанавливаем поля, если заданы
            if (pageMargins) {
                if (!sectionProps.page) sectionProps.page = {};
                sectionProps.page.margin = pageMargins;
            }
            // Формируем объект секции
            const section = {
                properties: sectionProps,
                children: docChildren
            };
            // Если есть колонтитулы, добавляем их
            if (headerContent.length > 0) {
                section.headers = { default: new Header({ children: headerContent }) };
            }
            if (footerContent.length > 0) {
                section.footers = { default: new Footer({ children: footerContent }) };
            }
            // Если использовались списки, формируем конфигурацию нумерации
            let numberingConfig = undefined;
            if (numberingUsed || bulletListUsed) {
                numberingConfig = { config: [] };
                if (numberingUsed) {
                    numberingConfig.config.push({
                        reference: "default-numbering",
                        levels: [{
                            level: 0,
                            format: "decimal",    // десятичная нумерация
                            text: "%1.",          // "1.", "2.", ...
                            alignment: AlignmentType.START
                        }]
                    });
                }
                if (bulletListUsed) {
                    numberingConfig.config.push({
                        reference: "default-bullet",
                        levels: [{
                            level: 0,
                            format: "bullet",
                            text: "\u2022",       // маркер "•"
                            alignment: AlignmentType.START
                        }]
                    });
                }
            }
            // Создаем документ с указанной секцией и метаданными
            doc = new Document({
                sections: [ section ],
                title: docProps.title || undefined,
                creator: docProps.creator || undefined,
                subject: docProps.subject || undefined,
                numbering: numberingConfig
            });
            // Пакуем документ в буфер и сохраняем в файл
            const buffer = await Packer.toBuffer(doc);
            fs.writeFileSync(filePath, buffer);
            return makeResponse(true, cmd, `Document saved to ${filePath}`);
        } catch (err) {
            return makeResponse(false, cmd, "Failed to save document: " + err.message);
        }
    }

    case "add_paragraph": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        const text = params.text || params;
        if (typeof text !== "string") {
            return makeResponse(false, cmd, "Parameter 'text' is required for add_paragraph");
        }
        // Создаем параграф с заданным текстом
        const para = new Paragraph(text);
        docChildren.push(para);
        lastParagraph = para;
        currentTable = null;  // завершаем добавление в таблицу, если было
        return makeResponse(true, cmd, "Paragraph added");
    }

    case "add_heading": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        const text = params.text || params.heading || params;
        let level = params.level || params.size || params.headingLevel;
        if (typeof text !== "string") {
            return makeResponse(false, cmd, "Parameter 'text' is required for add_heading");
        }
        if (!level || isNaN(level)) level = 1;
        if (level < 1 || level > 5) level = 1;
        // Создаем параграф с текстом заголовка
        let headingPara = new Paragraph(text);
        // Применяем стиль заголовка в зависимости от уровня
        switch (level) {
            case 1: headingPara.heading(HeadingLevel.HEADING_1); break;
            case 2: headingPara.heading(HeadingLevel.HEADING_2); break;
            case 3: headingPara.heading(HeadingLevel.HEADING_3); break;
            case 4: headingPara.heading(HeadingLevel.HEADING_4); break;
            case 5: headingPara.heading(HeadingLevel.HEADING_5); break;
            default: headingPara.heading(HeadingLevel.HEADING_1);
        }
        docChildren.push(headingPara);
        lastParagraph = headingPara;
        currentTable = null;
        return makeResponse(true, cmd, `Heading (level ${level}) added`);
    }

    case "add_bullet_list": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        let items = params.items || params;
        if (!Array.isArray(items)) {
            // Возможно передали одиночный элемент
            if (typeof params === "string") {
                items = [params];
            } else {
                return makeResponse(false, cmd, "Parameter 'items' (array of strings) is required for add_bullet_list");
            }
        }
        items.forEach(item => {
            const para = new Paragraph(item);
            // Отметить как пункт маркерованного списка
            // Способ 1: с использованием встроенного метода bullet()
            // para.bullet();
            // Способ 2: через нумерацию с форматированием bullet (предпочтительно для единообразия)
            para.numbering({ reference: "default-bullet", level: 0 });
            docChildren.push(para);
            // lastParagraph обновляем на последний добавленный пункт
            lastParagraph = para;
        });
        bulletListUsed = true;
        currentTable = null;
        return makeResponse(true, cmd, `Bullet list added (${items.length} items)`);
    }

    case "add_numbered_list": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        let items = params.items || params;
        if (!Array.isArray(items)) {
            if (typeof params === "string") {
                items = [params];
            } else {
                return makeResponse(false, cmd, "Parameter 'items' (array of strings) is required for add_numbered_list");
            }
        }
        items.forEach(item => {
            const para = new Paragraph(item);
            // Назначаем параграфу нумерацию (уровень 0 списка "default-numbering")
            para.numbering({ reference: "default-numbering", level: 0 });
            docChildren.push(para);
            lastParagraph = para;
        });
        numberingUsed = true;
        currentTable = null;
        return makeResponse(true, cmd, `Numbered list added (${items.length} items)`);
    }

    case "set_alignment": {
        if (!lastParagraph) {
            return makeResponse(false, cmd, "No paragraph available to align");
        }
        const alignment = (params.align || params.alignment || params).toString().toLowerCase();
        switch (alignment) {
            case "left":
                lastParagraph.alignment = AlignmentType.LEFT;
                break;
            case "right":
                lastParagraph.alignment = AlignmentType.RIGHT;
                break;
            case "center":
            case "centre":
                lastParagraph.alignment = AlignmentType.CENTER;
                break;
            case "justify":
            case "justified":
                lastParagraph.alignment = AlignmentType.JUSTIFIED;
                break;
            default:
                return makeResponse(false, cmd, "Invalid alignment value");
        }
        return makeResponse(true, cmd, `Alignment set to ${alignment}`);
    }

    case "set_font_size": {
        if (!lastParagraph) {
            return makeResponse(false, cmd, "No paragraph available to set font size");
        }
        let sizePt = params.size || params.font_size || params;
        if (typeof sizePt !== "number") {
            sizePt = parseFloat(sizePt);
        }
        if (!sizePt || isNaN(sizePt)) {
            return makeResponse(false, cmd, "Invalid font size");
        }
        // docx ожидает "half-points"
        const halfPoints = Math.round(sizePt * 2);
        lastParagraph.fontSize(halfPoints);
        return makeResponse(true, cmd, `Font size set to ${sizePt}pt`);
    }

    case "set_font_family": {
        if (!lastParagraph) {
            return makeResponse(false, cmd, "No paragraph available to set font family");
        }
        const fontName = params.name || params.font || params.font_family || params;
        if (typeof fontName !== "string" || fontName.length === 0) {
            return makeResponse(false, cmd, "Invalid font name");
        }
        lastParagraph.font(fontName);
        return makeResponse(true, cmd, `Font family set to '${fontName}'`);
    }

    case "set_spacing": {
        if (!lastParagraph) {
            return makeResponse(false, cmd, "No paragraph available to set spacing");
        }
        let before = 0, after = 0, line = 0;
        if (typeof params === "number") {
            // Если передано одно число, трактуем как межстрочный коэффициент
            line = Math.round(params * 240);
        } else {
            if (params.before) {
                // Переводим пункты в TWIP (1 пт = 20 TWIP)
                before = Math.round(parseFloat(params.before) * 20);
            }
            if (params.after) {
                after = Math.round(parseFloat(params.after) * 20);
            }
            if (params.line) {
                const lineParam = parseFloat(params.line);
                if (!isNaN(lineParam)) {
                    // >0 воспринимаем как множитель
                    if (lineParam > 0 && lineParam <= 10) {
                        line = Math.round(lineParam * 240);
                    } else {
                        // Если указано большое число, возможно уже в TWIP, ограничимся 0 (авто)
                        line = Math.round(lineParam);
                    }
                }
            }
        }
        lastParagraph.spacing({ before: before || 0, after: after || 0, line: line || 0 });
        return makeResponse(true, cmd, "Spacing updated");
    }

    case "add_table": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        // Если есть незавершенная таблица, начинаем новую
        currentTable = new Table({ rows: [] });
        // Если задано количество столбцов, сохраним для использования при добавлении строк (можно задать ширины)
        let numCols = params.columns || params.cols;
        // Если переданы данные таблицы (массив массивов)
        if (Array.isArray(params.data) || Array.isArray(params.rows)) {
            const dataRows = params.data || params.rows;
            // Добавляем каждую строку
            dataRows.forEach(rowData => {
                const cells = [];
                rowData.forEach(cellText => {
                    const cellParagraph = new Paragraph(cellText != null ? String(cellText) : "");
                    cells.push(new TableCell({ children: [cellParagraph] }));
                });
                const tableRow = new TableRow({ children: cells });
                // Добавляем в текущую таблицу
                currentTable.root.push(tableRow);
                // После добавления строки, currentTable.rows.length увеличивается (но мы используем root напрямую)
            });
            // Определяем число столбцов по первой строке данных, если не задано явно
            if (!numCols && dataRows.length > 0) {
                numCols = dataRows[0].length;
            }
        }
        // TODO: В перспективе можно задать столбцы (numCols) и их ширину, но пропустим для простоты
        // Добавляем таблицу в документ
        docChildren.push(currentTable);
        lastParagraph = null;
        return makeResponse(true, cmd, "Table created");
    }

    case "add_table_row": {
        if (!currentTable) {
            return makeResponse(false, cmd, "No active table. Use add_table first.");
        }
        let cellsData = params.cells || params.row || params;
        if (!Array.isArray(cellsData)) {
            return makeResponse(false, cmd, "Parameter 'cells' (array) is required for add_table_row");
        }
        // Создаем ячейки таблицы из переданных данных
        const cellElements = cellsData.map(cell => {
            const text = cell != null ? String(cell) : "";
            return new TableCell({
                children: [ new Paragraph(text) ]
            });
        });
        const newRow = new TableRow({ children: cellElements });
        currentTable.root.push(newRow);
        return makeResponse(true, cmd, "Row added to table");
    }

    case "set_table_style": {
        if (!currentTable) {
            return makeResponse(false, cmd, "No active table to style.");
        }
        const styleName = params.style || params;
        if (typeof styleName !== "string" || styleName.trim() === "") {
            return makeResponse(false, cmd, "Invalid table style name");
        }
        try {
            // Устанавливаем имя стиля таблицы (через TableProperties)
            // В docx это можно сделать добавлением TableProperties с tblStyle
            currentTable.properties = currentTable.properties || {};
            currentTable.properties.style = styleName;
            return makeResponse(true, cmd, `Table style set to '${styleName}'`);
        } catch (err) {
            return makeResponse(false, cmd, "Failed to set table style: " + err.message);
        }
    }

    case "merge_cells": {
        if (!currentTable) {
            return makeResponse(false, cmd, "No active table to merge cells in.");
        }
        // Ожидаем параметры startRow, startCol, endRow, endCol (индексация с 0)
        const sr = params.startRow, sc = params.startCol;
        const er = params.endRow, ec = params.endCol;
        if ([sr, sc, er, ec].some(v => typeof v !== "number")) {
            return makeResponse(false, cmd, "merge_cells requires startRow, startCol, endRow, endCol");
        }
        if (sr > er || sc > ec) {
            return makeResponse(false, cmd, "Invalid cell range for merge");
        }
        // Получаем внутренний представление таблицы (массив строк)
        const tableRows = currentTable.root;  // Table.root содержит массив дочерних элементов (TableRow и TableProperties)
        // Проверка границ
        if (sr < 0 || er >= tableRows.length) {
            return makeResponse(false, cmd, "Row index out of range");
        }
        // Определяем сколько строк и столбцов объединяется
        const rowCount = er - sr + 1;
        let colCount = ec - sc + 1;
        // Обходим указанный диапазон
        for (let r = sr; r <= er; r++) {
            // Каждый элемент tableRows[r] – это TableRow с свойством children (массив ячеек)
            const row = tableRows[r];
            if (!row || !row.children) continue;
            for (let c = sc; c <= ec; c++) {
                if (r === sr && c === sc) {
                    // Верхняя левая ячейка диапазона - сохраняем её, задаём span
                    const cell = row.children[c];
                    if (!cell) continue;
                    if (rowCount > 1) {
                        cell.properties = cell.properties || {};
                        cell.properties.rowSpan = rowCount;
                    }
                    if (colCount > 1) {
                        cell.properties = cell.properties || {};
                        cell.properties.columnSpan = colCount;
                    }
                } else {
                    // Все остальные ячейки диапазона удаляем
                    if (row.children[c]) {
                        row.children.splice(c, 1);
                    }
                }
            }
        }
        return makeResponse(true, cmd, `Cells merged from (${sr},${sc}) to (${er},${ec})`);
    }

    case "insert_image": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        const imagePath = params.path || params.file;
        if (!imagePath) {
            return makeResponse(false, cmd, "Parameter 'path' is required for insert_image");
        }
        let width = params.width, height = params.height;
        try {
            if (!fs.existsSync(imagePath)) {
                return makeResponse(false, cmd, "Image file not found: " + imagePath);
            }
            const imageData = fs.readFileSync(imagePath);
            const imageOpts = { data: imageData };
            // Если указаны размеры, добавляем трансформацию (размеры в пунктах или пикселях)
            if (width || height) {
                // Если задан только один размер, второй подгоняется пропорционально – для простоты опустим вычисление
                imageOpts.transformation = {};
                if (width) imageOpts.transformation.width = parseInt(width);
                if (height) imageOpts.transformation.height = parseInt(height);
            }
            const imageRun = new ImageRun(imageOpts);
            const paragraph = new Paragraph({ children: [imageRun] });
            docChildren.push(paragraph);
            lastParagraph = paragraph;
            currentTable = null;
            return makeResponse(true, cmd, "Image inserted");
        } catch (err) {
            return makeResponse(false, cmd, "Failed to insert image: " + err.message);
        }
    }

    case "insert_image_with_caption": {
        // Вставляем изображение, затем подпись
        const imageRes = await handleCommand({ command: "insert_image", params: params });
        if (imageRes.status === "error") {
            return imageRes;  // вернуть ошибку если не удалось вставить изображение
        }
        const captionText = params.caption || params.text;
        if (captionText && typeof captionText === "string") {
            const captionPara = new Paragraph(captionText);
            // Сделаем подпись курсивом и выравниваем по центру (как пример оформления)
            captionPara.italics();
            captionPara.alignment = AlignmentType.CENTER;
            docChildren.push(captionPara);
            lastParagraph = captionPara;
        }
        return makeResponse(true, cmd, "Image with caption inserted");
    }

    case "add_header": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        const text = params.text || params;
        if (typeof text !== "string") {
            return makeResponse(false, cmd, "Parameter 'text' is required for add_header");
        }
        const headerPara = new Paragraph(text);
        headerContent.push(headerPara);
        return makeResponse(true, cmd, "Header content added");
    }

    case "add_footer": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        const text = params.text || params;
        if (typeof text !== "string") {
            return makeResponse(false, cmd, "Parameter 'text' is required for add_footer");
        }
        const footerPara = new Paragraph(text);
        footerContent.push(footerPara);
        return makeResponse(true, cmd, "Footer content added");
    }

    case "add_page_numbering": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        // Создаем параграф с номером страницы / количеством страниц
        const pageNumberPara = new Paragraph({
            children: [
                new TextRun("Стр. "), 
                PageNumber.CURRENT,
                new TextRun(" из "),
                PageNumber.TOTAL_PAGES
            ],
            alignment: AlignmentType.CENTER
        });
        footerContent.push(pageNumberPara);
        return makeResponse(true, cmd, "Page numbering added (footer)");
    }

    case "add_page_break": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        // Вставляем разрыв страницы
        const breakPara = new Paragraph({ children: [ new PageBreak() ] });
        docChildren.push(breakPara);
        lastParagraph = null;
        currentTable = null;
        return makeResponse(true, cmd, "Page break added");
    }

    case "set_margins": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        // Принимаем поля в сантиметрах либо дюймах (если значение < 5, допустим, считаем в см; если >5 — возможно в дюймах)
        const top = parseFloat(params.top ?? params.t);
        const bottom = parseFloat(params.bottom ?? params.b);
        const left = parseFloat(params.left ?? params.l);
        const right = parseFloat(params.right ?? params.r);
        if ([top, bottom, left, right].some(v => isNaN(v))) {
            return makeResponse(false, cmd, "Invalid margins (top, bottom, left, right required)");
        }
        // Определяем, в чем задано – если все значения <= 5, предположим, что в сантиметрах, иначе в дюймах
        let inInches = false;
        const vals = [top, bottom, left, right];
        if (vals.some(v => v > 5)) {
            inInches = true;
        }
        // Переводим в TWIP
        const factor = inInches ? 1440 : Math.round(1440 / 2.54);  // 1 inch = 1440 twips; 1 cm ~ 567 twips
        pageMargins = {
            top: Math.round(top * factor),
            bottom: Math.round(bottom * factor),
            left: Math.round(left * factor),
            right: Math.round(right * factor)
        };
        return makeResponse(true, cmd, `Page margins set (${top}${inInches?"in":"cm"} top, etc.)`);
    }

    case "set_orientation": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        const orientation = (params.orientation || params).toString().toLowerCase();
        if (orientation === "landscape") {
            pageOrientation = PageOrientation.LANDSCAPE;
        } else if (orientation === "portrait") {
            pageOrientation = PageOrientation.PORTRAIT;
        } else {
            return makeResponse(false, cmd, "Invalid orientation value");
        }
        return makeResponse(true, cmd, `Orientation set to ${orientation}`);
    }

    case "set_title": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        const title = params.title || params;
        if (typeof title !== "string") {
            return makeResponse(false, cmd, "Invalid title");
        }
        docProps.title = title;
        return makeResponse(true, cmd, `Document title set to '${title}'`);
    }

    case "set_author": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        const author = params.author || params.name || params;
        if (typeof author !== "string") {
            return makeResponse(false, cmd, "Invalid author name");
        }
        docProps.creator = author;
        return makeResponse(true, cmd, `Document author set to '${author}'`);
    }

    case "set_subject": {
        if (!doc) {
            return makeResponse(false, cmd, "No document. Use create_document first.");
        }
        const subject = params.subject || params;
        if (typeof subject !== "string") {
            return makeResponse(false, cmd, "Invalid subject");
        }
        docProps.subject = subject;
        return makeResponse(true, cmd, `Document subject set to '${subject}'`);
    }

    default:
        return makeResponse(false, cmd, "Unknown command: " + cmd);
}
}

// Запуск сервера: читаем построчно STDIN и обрабатываем как JSON команды
const rl = readline.createInterface({
input: process.stdin,
output: process.stdout,
terminal: false
});
rl.on("line", async (input) => {
let msg;
try {
    msg = JSON.parse(input);
} catch (err) {
    const errorResp = { status: "error", error: "Invalid JSON input" };
    process.stdout.write(JSON.stringify(errorResp) + "\n");
    return;
}
const response = await handleCommand(msg);
// Печатаем ответ в STDOUT
process.stdout.write(JSON.stringify(response) + "\n");
});
