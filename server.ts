import express from 'express';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';
import vm from 'node:vm';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import {
  AlignmentType,
  Document,
  Packer,
  PageBreak,
  Paragraph,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  TextRun,
  UnderlineType,
  WidthType,
} from 'docx';
import PptxGenJS from 'pptxgenjs';
import ExcelJS from 'exceljs';
import PDFDocument from 'pdfkit';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ASSETS_DIR = path.join(__dirname, 'assets');
const FONT_PATH = path.join(ASSETS_DIR, 'DejaVuSans.ttf');
const PY_RUNNER_PATH = path.join(__dirname, 'server', 'py_runner.py');

interface StoredFile {
  id: string;
  filename: string;
  mimeType: string;
  buffer: Buffer;
  createdAt: number;
}

// In-memory store for generated files (ephemeral container runtime)
const generatedFiles = new Map<string, StoredFile>();
const MAX_STORED_FILES = 100;

function storeGeneratedFile(filename: string, mimeType: string, buffer: Buffer): StoredFile {
  const safeName = path.basename(filename.trim()) || 'document';
  const id = crypto.randomUUID();
  const entry: StoredFile = {
    id,
    filename: safeName,
    mimeType,
    buffer,
    createdAt: Date.now(),
  };
  generatedFiles.set(id, entry);
  if (generatedFiles.size > MAX_STORED_FILES) {
    const oldestKey = generatedFiles.keys().next().value;
    if (oldestKey) {
      generatedFiles.delete(oldestKey);
    }
  }
  return entry;
}

function parseCsv(csvString: string): string[][] {
  if (!csvString || !csvString.trim()) {
    return [];
  }
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < csvString.length; i++) {
    const ch = csvString[i];
    const next = csvString[i + 1];

    if (inQuotes) {
      if (ch === '"' && next === '"') {
        field += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        field += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        row.push(field);
        field = '';
      } else if (ch === '\r' && next === '\n') {
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
        i++;
      } else if (ch === '\n' || ch === '\r') {
        row.push(field);
        rows.push(row);
        row = [];
        field = '';
      } else {
        field += ch;
      }
    }
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows;
}

async function createDocx(
  title: string,
  text: string,
  listItems: string,
  tableCsv: string,
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  const children: Array<Paragraph | Table> = [];

  if (title) {
    children.push(
      new Paragraph({
        text: String(title),
        heading: HeadingLevel.HEADING_1,
      })
    );
  }

  if (text) {
    for (const line of String(text).split('\n')) {
      children.push(new Paragraph({ text: line }));
    }
  }

  if (listItems) {
    for (const raw of String(listItems).split('\n')) {
      const item = raw.trim();
      if (item) {
        children.push(
          new Paragraph({
            text: item,
            bullet: { level: 0 },
          })
        );
      }
    }
  }

  const tableData = parseCsv(tableCsv || '');
  if (tableData.length > 0) {
    const cols = tableData[0]?.length ?? 0;
    if (cols > 0) {
      const table = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: tableData.map(
          (row) =>
            new TableRow({
              children: Array.from({ length: cols }, (_, j) => {
                const value = row[j] ?? '';
                return new TableCell({
                  children: [new Paragraph({ text: String(value) })],
                });
              }),
            })
        ),
      });
      children.push(table);
    }
  }

  if (children.length === 0) {
    children.push(new Paragraph({ text: '' }));
  }

  const doc = new Document({
    sections: [{ children }],
  });

  const buffer = await Packer.toBuffer(doc);
  const stored = storeGeneratedFile(
    filename,
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer
  );
  return { message: `Файл ${stored.filename} сохранен в Downloads`, stored };
}

async function createPptx(
  title: string,
  text: string,
  listItems: string,
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  const PptxCtor =
    (PptxGenJS as unknown as { default?: typeof PptxGenJS }).default ?? PptxGenJS;
  const prs = new PptxCtor();
  const slide = prs.addSlide();

  slide.addText(String(title || ''), {
    x: 0.6,
    y: 0.4,
    w: 8.8,
    h: 0.9,
    fontSize: 26,
    bold: true,
    color: '1F2937',
  });

  const bodyRuns: Array<{ text: string; options?: Record<string, unknown> }> = [];
  if (text) {
    bodyRuns.push({
      text: String(text),
      options: { breakLine: true, fontSize: 16, color: '374151' },
    });
  }
  if (listItems) {
    for (const raw of String(listItems).split('\n')) {
      const item = raw.trim();
      if (item) {
        bodyRuns.push({
          text: item,
          options: { bullet: true, indentLevel: 1, fontSize: 15, color: '374151' },
        });
      }
    }
  }

  if (bodyRuns.length > 0) {
    slide.addText(bodyRuns, {
      x: 0.6,
      y: 1.5,
      w: 8.8,
      h: 4.5,
      valign: 'top',
    });
  }

  const output = await prs.write({ outputType: 'nodebuffer' });
  const buffer = Buffer.isBuffer(output)
    ? output
    : Buffer.from(output as ArrayBuffer);

  const stored = storeGeneratedFile(
    filename,
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer
  );
  return { message: `Файл ${stored.filename} сохранен в Downloads`, stored };
}

async function createXlsx(
  tableCsv: string,
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  const tableData = parseCsv(tableCsv || '');
  if (tableData.length === 0) {
    throw new Error('Нет данных для таблицы Excel.');
  }

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');

  for (let i = 0; i < tableData.length; i++) {
    const row = tableData[i];
    for (let j = 0; j < row.length; j++) {
      const cell = ws.getCell(i + 1, j + 1);
      cell.value = row[j];
      cell.alignment = { wrapText: true };
      if (i === 0) {
        cell.font = { bold: true };
      }
    }
  }

  const arrayBuffer = await wb.xlsx.writeBuffer();
  const buffer = Buffer.from(arrayBuffer);
  const stored = storeGeneratedFile(
    filename,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer
  );
  return { message: `Файл ${stored.filename} сохранен в Downloads`, stored };
}

function resolveValidFontPath(): string | null {
  const candidates = [
    FONT_PATH,
    '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
    '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
  ];
  for (const candidate of candidates) {
    if (!fs.existsSync(candidate)) continue;
    try {
      const fd = fs.openSync(candidate, 'r');
      const header = Buffer.alloc(4);
      fs.readSync(fd, header, 0, 4, 0);
      fs.closeSync(fd);
      const hex = header.toString('hex');
      if (hex === '00010000' || hex === '4f54544f' || hex === '74727565') {
        return candidate;
      }
    } catch {
      // Ignore unreadable candidate
    }
  }
  return null;
}

async function createPdf(
  title: string,
  text: string,
  listItems: string,
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  const validFontPath = resolveValidFontPath();

  const buffer = await new Promise<Buffer>((resolve, reject) => {
    const doc = new PDFDocument({ margin: 50 });
    const chunks: Buffer[] = [];
    doc.on('data', (chunk: Buffer) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    if (validFontPath) {
      doc.registerFont('DejaVu', validFontPath);
      doc.font('DejaVu');
    } else {
      doc.font('Helvetica');
    }

    doc.fontSize(16).text(String(title || ''), { align: 'center' });
    doc.moveDown(1);
    doc.fontSize(12).text(String(text || ''));

    if (listItems) {
      for (const raw of String(listItems).split('\n')) {
        const item = raw.trim();
        if (item) {
          doc.text(`• ${item}`);
        }
      }
    }

    doc.end();
  });

  const stored = storeGeneratedFile(filename, 'application/pdf', buffer);
  return { message: `Файл ${stored.filename} сохранен в Downloads`, stored };
}

interface ParsedLayout {
  title: string;
  text: string;
  listItems: string;
  tableCsv: string;
}

function escapeCsvCell(value: string): string {
  const s = String(value ?? '');
  if (s.includes(',') || s.includes('"') || s.includes('\n')) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

function parseLayoutFromText(rawText: string, sourceFilename?: string): ParsedLayout {
  const trimmed = String(rawText || '').trim();
  if (!trimmed) {
    return { title: '', text: '', listItems: '', tableCsv: '' };
  }

  // 1. JSON format
  if (
    sourceFilename?.toLowerCase().endsWith('.json') ||
    (trimmed.startsWith('{') && trimmed.endsWith('}'))
  ) {
    try {
      const parsed = JSON.parse(trimmed) as Record<string, unknown>;
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
        const listVal = parsed.listItems ?? parsed.list ?? parsed.items;
        const listStr = Array.isArray(listVal)
          ? listVal.map(String).join('\n')
          : String(listVal ?? '');
        const tableVal = parsed.tableCsv ?? parsed.table;
        let tableStr = '';
        if (Array.isArray(tableVal)) {
          tableStr = tableVal
            .map((r) =>
              Array.isArray(r) ? r.map((c) => escapeCsvCell(String(c))).join(',') : String(r)
            )
            .join('\n');
        } else if (typeof tableVal === 'string') {
          tableStr = tableVal;
        }
        return {
          title: String(parsed.title ?? ''),
          text: String(parsed.text ?? parsed.body ?? ''),
          listItems: listStr,
          tableCsv: tableStr,
        };
      }
    } catch {
      // Fall through to structured text parsing
    }
  }

  // 2. Pure CSV file
  if (sourceFilename?.toLowerCase().endsWith('.csv')) {
    const baseTitle = path.basename(sourceFilename, path.extname(sourceFilename));
    return {
      title: baseTitle,
      text: '',
      listItems: '',
      tableCsv: trimmed,
    };
  }

  // 3. Structured text / Markdown / Script-style markup
  const lines = rawText.split(/\r?\n/);
  let title = '';
  const textLines: string[] = [];
  const listLines: string[] = [];
  const tableRows: string[] = [];
  let inTableBlock = false;

  for (const rawLine of lines) {
    const line = rawLine.trimEnd();
    const stripped = line.trim();

    if (/^TABLE:\s*$/i.test(stripped)) {
      inTableBlock = true;
      continue;
    }

    if (inTableBlock) {
      if (!stripped) {
        inTableBlock = false;
        continue;
      }
      if (stripped.startsWith('|') && stripped.endsWith('|')) {
        if (/^\|[\s:-]+\|/.test(stripped) && !/[^\s|:-]/.test(stripped)) {
          continue;
        }
        const cells = stripped
          .slice(1, -1)
          .split('|')
          .map((c) => escapeCsvCell(c.trim()));
        tableRows.push(cells.join(','));
      } else {
        tableRows.push(stripped);
      }
      continue;
    }

    // TITLE: directive
    const titleMatch = stripped.match(/^(?:TITLE|Заголовок)\s*:\s*(.+)$/i);
    if (titleMatch && !title) {
      title = titleMatch[1].trim();
      continue;
    }

    // Markdown H1 (# Heading)
    const h1Match = stripped.match(/^#\s+(.+)$/);
    if (h1Match && !title) {
      title = h1Match[1].trim();
      continue;
    }

    // Markdown H2/H3
    const subHeadingMatch = stripped.match(/^#{2,6}\s+(.+)$/);
    if (subHeadingMatch) {
      textLines.push(subHeadingMatch[1].trim());
      continue;
    }

    // Markdown table row | a | b |
    if (stripped.startsWith('|') && stripped.endsWith('|') && stripped.length > 2) {
      // Skip markdown separator line like |---|---|
      if (!/[^\s|:-]/.test(stripped)) {
        continue;
      }
      const cells = stripped
        .slice(1, -1)
        .split('|')
        .map((c) => escapeCsvCell(c.trim()));
      tableRows.push(cells.join(','));
      continue;
    }

    // Bullet / numbered list items
    const bulletMatch = stripped.match(/^(?:[-*•]|\d+[.)])\s+(.+)$/);
    if (bulletMatch) {
      listLines.push(bulletMatch[1].trim());
      continue;
    }

    textLines.push(line);
  }

  if (!title && sourceFilename) {
    title = path.basename(sourceFilename, path.extname(sourceFilename));
  }

  return {
    title,
    text: textLines.join('\n').trim(),
    listItems: listLines.join('\n'),
    tableCsv: tableRows.join('\n'),
  };
}

async function createTxt(
  title: string,
  text: string,
  listItems: string,
  tableCsv: string,
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  const parts: string[] = [];
  if (title) {
    parts.push(title.toUpperCase());
    parts.push('='.repeat(Math.min(Math.max(title.length, 3), 60)));
    parts.push('');
  }
  if (text) {
    parts.push(text);
    parts.push('');
  }
  if (listItems) {
    const items = listItems
      .split('\n')
      .map((s) => s.trim())
      .filter(Boolean);
    if (items.length > 0) {
      for (const item of items) {
        parts.push(`• ${item}`);
      }
      parts.push('');
    }
  }
  const tableData = parseCsv(tableCsv || '');
  if (tableData.length > 0) {
    for (const row of tableData) {
      parts.push(row.join('\t|\t'));
    }
  }

  const buffer = Buffer.from(parts.join('\n').trim() + '\n', 'utf-8');
  const stored = storeGeneratedFile(filename, 'text/plain; charset=utf-8', buffer);
  return { message: `Файл ${stored.filename} сохранен в Downloads`, stored };
}

async function createMd(
  title: string,
  text: string,
  listItems: string,
  tableCsv: string,
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  const parts: string[] = [];
  if (title) {
    parts.push(`# ${title}`);
    parts.push('');
  }
  if (text) {
    parts.push(text);
    parts.push('');
  }
  if (listItems) {
    const items = listItems
      .split('\n')
      .map((s) => s.trim())
      .filter(Boolean);
    if (items.length > 0) {
      for (const item of items) {
        parts.push(`- ${item}`);
      }
      parts.push('');
    }
  }
  const tableData = parseCsv(tableCsv || '');
  if (tableData.length > 0) {
    const cols = tableData[0]?.length ?? 0;
    if (cols > 0) {
      parts.push(`| ${tableData[0].map((c) => c || ' ').join(' | ')} |`);
      parts.push(`| ${Array.from({ length: cols }, () => '---').join(' | ')} |`);
      for (let i = 1; i < tableData.length; i++) {
        const padded = Array.from({ length: cols }, (_, j) => tableData[i][j] ?? '');
        parts.push(`| ${padded.join(' | ')} |`);
      }
    }
  }

  const buffer = Buffer.from(parts.join('\n').trim() + '\n', 'utf-8');
  const stored = storeGeneratedFile(filename, 'text/markdown; charset=utf-8', buffer);
  return { message: `Файл ${stored.filename} сохранен в Downloads`, stored };
}

function escapeHtml(value: string): string {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

async function createHtml(
  title: string,
  text: string,
  listItems: string,
  tableCsv: string,
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  const sections: string[] = [];
  if (title) {
    sections.push(`<h1>${escapeHtml(title)}</h1>`);
  }
  if (text) {
    for (const para of text.split('\n')) {
      if (para.trim()) {
        sections.push(`<p>${escapeHtml(para)}</p>`);
      }
    }
  }
  if (listItems) {
    const items = listItems
      .split('\n')
      .map((s) => s.trim())
      .filter(Boolean);
    if (items.length > 0) {
      sections.push(`<ul>${items.map((i) => `<li>${escapeHtml(i)}</li>`).join('')}</ul>`);
    }
  }
  const tableData = parseCsv(tableCsv || '');
  if (tableData.length > 0) {
    const rowsHtml = tableData
      .map((row, idx) => {
        const tag = idx === 0 ? 'th' : 'td';
        return `<tr>${row.map((cell) => `<${tag}>${escapeHtml(cell)}</${tag}>`).join('')}</tr>`;
      })
      .join('');
    sections.push(`<table>${rowsHtml}</table>`);
  }

  const html = `<!doctype html>
<html lang="ru">
<head>
  <meta charset="UTF-8" />
  <title>${escapeHtml(title || filename)}</title>
  <style>
    body { font-family: system-ui, -apple-system, sans-serif; max-width: 800px; margin: 40px auto; padding: 0 20px; color: #1f2937; line-height: 1.6; }
    h1 { margin-bottom: 16px; }
    table { border-collapse: collapse; width: 100%; margin-top: 16px; }
    th, td { border: 1px solid #d1d5db; padding: 8px 10px; text-align: left; }
    th { background: #f3f4f6; }
  </style>
</head>
<body>
${sections.join('\n')}
</body>
</html>
`;

  const buffer = Buffer.from(html, 'utf-8');
  const stored = storeGeneratedFile(filename, 'text/html; charset=utf-8', buffer);
  return { message: `Файл ${stored.filename} сохранен в Downloads`, stored };
}

async function generateByFormat(
  format: string,
  layout: ParsedLayout,
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  let targetFilename = String(filename || '').trim();
  if (!targetFilename) {
    throw new Error('Укажите имя файла.');
  }

  if (format === 'docx') {
    if (!targetFilename.endsWith('.docx')) targetFilename += '.docx';
    return createDocx(layout.title, layout.text, layout.listItems, layout.tableCsv, targetFilename);
  }
  if (format === 'pptx') {
    if (!targetFilename.endsWith('.pptx')) targetFilename += '.pptx';
    return createPptx(layout.title, layout.text, layout.listItems, targetFilename);
  }
  if (format === 'xlsx') {
    if (!targetFilename.endsWith('.xlsx')) targetFilename += '.xlsx';
    return createXlsx(layout.tableCsv, targetFilename);
  }
  if (format === 'pdf') {
    if (!targetFilename.endsWith('.pdf')) targetFilename += '.pdf';
    return createPdf(layout.title, layout.text, layout.listItems, targetFilename);
  }
  if (format === 'txt') {
    if (!targetFilename.endsWith('.txt')) targetFilename += '.txt';
    return createTxt(layout.title, layout.text, layout.listItems, layout.tableCsv, targetFilename);
  }
  if (format === 'md') {
    if (!targetFilename.endsWith('.md')) targetFilename += '.md';
    return createMd(layout.title, layout.text, layout.listItems, layout.tableCsv, targetFilename);
  }
  if (format === 'html') {
    if (!targetFilename.endsWith('.html')) targetFilename += '.html';
    return createHtml(layout.title, layout.text, layout.listItems, layout.tableCsv, targetFilename);
  }
  throw new Error(`Неизвестный формат: ${String(format)}`);
}

function extractFencedCodeIfMarkdown(code: string): string {
  const trimmed = code.trim();
  const blocks: string[] = [];
  const regex = /```(?:python|py|javascript|js)?\s*\n([\s\S]*?)```/gi;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(trimmed)) !== null) {
    if (match[1]?.trim()) {
      blocks.push(match[1].trim());
    }
  }
  if (blocks.length > 0 && /^(?:#|\s*```)/.test(trimmed)) {
    return blocks.join('\n\n');
  }
  return code;
}

function preprocessScript(code: string): string {
  const normalizedCode = extractFencedCodeIfMarkdown(code);
  // Convert Python triple-quoted strings ("""...""" and '''...''') into JS template literals
  const withTripleQuotesConverted = normalizedCode.replace(
    /"""([\s\S]*?)"""|'''([\s\S]*?)'''/g,
    (_m, d, s) => {
      const inner = String(d ?? s ?? '')
        .replace(/\\/g, '\\\\')
        .replace(/`/g, '\\`')
        .replace(/\$\{/g, '\\${');
      return '`' + inner + '`';
    }
  );

  const lines = withTripleQuotesConverted.split('\n');
  const transformedLines: string[] = [];

  for (const line of lines) {
    // Replace # comments outside string quotes with //
    let inSingle = false;
    let inDouble = false;
    let inBacktick = false;
    let escaped = false;
    let out = '';

    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (escaped) {
        out += ch;
        escaped = false;
        continue;
      }
      if (ch === '\\') {
        out += ch;
        escaped = true;
        continue;
      }
      if (!inDouble && !inBacktick && ch === "'") {
        inSingle = !inSingle;
        out += ch;
      } else if (!inSingle && !inBacktick && ch === '"') {
        inDouble = !inDouble;
        out += ch;
      } else if (!inSingle && !inDouble && ch === '`') {
        inBacktick = !inBacktick;
        out += ch;
      } else if (!inSingle && !inDouble && !inBacktick && ch === '#') {
        out += '//' + line.slice(i + 1);
        break;
      } else {
        out += ch;
      }
    }

    // Replace Python booleans/None outside quotes in a simple token-safe way
    out = out.replace(/\bTrue\b/g, 'true').replace(/\bFalse\b/g, 'false').replace(/\bNone\b/g, 'null');

    // Await async document helpers if not already awaited
    out = out.replace(
      /(?<!await\s)\b(create_docx|create_pptx|create_xlsx|create_pdf|create_txt|create_md|create_html|create_from_text|save_file_to_download)\s*\(/g,
      'await $1('
    );

    transformedLines.push(out);
  }

  return transformedLines.join('\n');
}

interface DocxIrRun {
  text?: string;
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: boolean;
  fontName?: string | null;
  fontSizePt?: number | null;
  colorHex?: string | null;
}

interface DocxIrBlock {
  type: 'paragraph' | 'table' | 'page_break';
  headingLevel?: number | null;
  alignment?: string | null;
  spaceBeforeTwips?: number | null;
  spaceAfterTwips?: number | null;
  firstLineIndentTwips?: number | null;
  leftIndentTwips?: number | null;
  rightIndentTwips?: number | null;
  lineSpacing?: number | null;
  lineSpacingPt?: number | null;
  pageBreakBefore?: boolean;
  runs?: DocxIrRun[];
  rows?: string[][];
}

interface DocxIrSpec {
  defaultFont?: string;
  defaultSizePt?: number;
  section?: {
    pageWidthTwips?: number | null;
    pageHeightTwips?: number | null;
    topMarginTwips?: number | null;
    bottomMarginTwips?: number | null;
    leftMarginTwips?: number | null;
    rightMarginTwips?: number | null;
  };
  blocks: DocxIrBlock[];
}

function mapDocxAlignment(
  val?: string | null
): (typeof AlignmentType)[keyof typeof AlignmentType] | undefined {
  if (!val) return undefined;
  const upper = val.toUpperCase();
  if (upper === 'CENTER') return AlignmentType.CENTER;
  if (upper === 'RIGHT' || upper === 'END') return AlignmentType.END;
  if (upper === 'LEFT' || upper === 'START') return AlignmentType.START;
  if (upper === 'JUSTIFY' || upper === 'BOTH' || upper === 'DISTRIBUTE') return AlignmentType.BOTH;
  return undefined;
}

function mapHeadingLevel(
  level?: number | null
): (typeof HeadingLevel)[keyof typeof HeadingLevel] | undefined {
  if (level === 1) return HeadingLevel.HEADING_1;
  if (level === 2) return HeadingLevel.HEADING_2;
  if (level === 3) return HeadingLevel.HEADING_3;
  if (level === 4) return HeadingLevel.HEADING_4;
  return undefined;
}

async function renderDocxFromIr(
  spec: DocxIrSpec,
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  const defaultFont = spec.defaultFont || 'Times New Roman';
  const defaultSizePt = spec.defaultSizePt || 12;
  const children: Array<Paragraph | Table> = [];

  for (const block of spec.blocks || []) {
    if (block.type === 'page_break') {
      children.push(new Paragraph({ children: [new PageBreak()] }));
      continue;
    }

    if (block.type === 'table') {
      const tableRows = block.rows || [];
      if (tableRows.length > 0) {
        const cols = Math.max(...tableRows.map((r) => r.length), 1);
        children.push(
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: tableRows.map(
              (row) =>
                new TableRow({
                  children: Array.from({ length: cols }, (_, j) => {
                    return new TableCell({
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: String(row[j] ?? ''),
                              font: defaultFont,
                              size: Math.round(defaultSizePt * 2),
                            }),
                          ],
                        }),
                      ],
                    });
                  }),
                })
            ),
          })
        );
      }
      continue;
    }

    const runs: Array<TextRun> = [];
    for (const r of block.runs || []) {
      const text = String(r.text ?? '');
      const parts = text.split('\n');
      parts.forEach((part, idx) => {
        const fontSizePt = r.fontSizePt ?? defaultSizePt;
        runs.push(
          new TextRun({
            text: part,
            break: idx > 0 ? 1 : undefined,
            bold: r.bold ?? undefined,
            italics: r.italic ?? undefined,
            underline: r.underline ? { type: UnderlineType.SINGLE } : undefined,
            font: r.fontName || defaultFont,
            size: Math.round(fontSizePt * 2),
            color: r.colorHex || undefined,
          })
        );
      });
    }

    if (runs.length === 0) {
      runs.push(
        new TextRun({
          text: '',
          font: defaultFont,
          size: Math.round(defaultSizePt * 2),
        })
      );
    }

    const spacing: { before?: number; after?: number; line?: number } = {};
    if (typeof block.spaceBeforeTwips === 'number') spacing.before = block.spaceBeforeTwips;
    if (typeof block.spaceAfterTwips === 'number') spacing.after = block.spaceAfterTwips;
    if (typeof block.lineSpacing === 'number') {
      spacing.line = Math.round(block.lineSpacing * 240);
    } else if (typeof block.lineSpacingPt === 'number') {
      spacing.line = Math.round(block.lineSpacingPt * 20);
    }

    const indent: { firstLine?: number; left?: number; right?: number } = {};
    if (typeof block.firstLineIndentTwips === 'number') indent.firstLine = block.firstLineIndentTwips;
    if (typeof block.leftIndentTwips === 'number') indent.left = block.leftIndentTwips;
    if (typeof block.rightIndentTwips === 'number') indent.right = block.rightIndentTwips;

    children.push(
      new Paragraph({
        children: runs,
        heading: mapHeadingLevel(block.headingLevel),
        alignment: mapDocxAlignment(block.alignment),
        pageBreakBefore: block.pageBreakBefore || undefined,
        spacing: Object.keys(spacing).length > 0 ? spacing : undefined,
        indent: Object.keys(indent).length > 0 ? indent : undefined,
      })
    );
  }

  if (children.length === 0) {
    children.push(new Paragraph({ text: '' }));
  }

  const sec = spec.section || {};
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: defaultFont,
            size: Math.round(defaultSizePt * 2),
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: sec.pageWidthTwips ?? 11906,
              height: sec.pageHeightTwips ?? 16838,
            },
            margin: {
              top: sec.topMarginTwips ?? 1134,
              bottom: sec.bottomMarginTwips ?? 1134,
              left: sec.leftMarginTwips ?? 1417,
              right: sec.rightMarginTwips ?? 850,
            },
          },
        },
        children,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  const stored = storeGeneratedFile(
    filename,
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer
  );
  return { message: `Файл ${stored.filename} сохранен в Downloads`, stored };
}

async function renderPptxFromIr(
  slides: Array<{ title: string; bodyLines: string[] }>,
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  const PptxCtor =
    (PptxGenJS as unknown as { default?: typeof PptxGenJS }).default ?? PptxGenJS;
  const prs = new PptxCtor();
  for (const s of slides.length > 0 ? slides : [{ title: '', bodyLines: [] }]) {
    const slide = prs.addSlide();
    slide.addText(String(s.title || ''), {
      x: 0.6,
      y: 0.4,
      w: 8.8,
      h: 0.9,
      fontSize: 26,
      bold: true,
      color: '1F2937',
    });
    if (s.bodyLines.length > 0) {
      slide.addText(
        s.bodyLines.map((line, i) => ({
          text: line,
          options: i === 0 ? { breakLine: true, fontSize: 16 } : { bullet: true, fontSize: 15 },
        })),
        { x: 0.6, y: 1.5, w: 8.8, h: 4.5, valign: 'top' }
      );
    }
  }
  const output = await prs.write({ outputType: 'nodebuffer' });
  const buffer = Buffer.isBuffer(output) ? output : Buffer.from(output as ArrayBuffer);
  const stored = storeGeneratedFile(
    filename,
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer
  );
  return { message: `Файл ${stored.filename} сохранен в Downloads`, stored };
}

async function renderXlsxFromIr(
  rows: string[][],
  filename: string
): Promise<{ message: string; stored: StoredFile }> {
  const csvStr = rows.map((r) => r.map(escapeCsvCell).join(',')).join('\n');
  return createXlsx(csvStr || ' ', filename);
}

function looksLikeExecutableScript(rawText: string, sourceFilename = ''): boolean {
  const code = String(rawText || '').trim();
  if (!code) return false;
  const lower = String(sourceFilename || '').toLowerCase();
  if (lower.endsWith('.py')) return true;
  const markers = [
    'from docx import',
    'import docx',
    'from pptx import',
    'import openpyxl',
    'from fpdf import',
    'create_docx(',
    'create_pptx(',
    'create_xlsx(',
    'create_pdf(',
    'create_from_text(',
    'doc.save(',
    'if __name__ == "__main__":',
    "if __name__ == '__main__':",
  ];
  return markers.some((m) => code.includes(m));
}

function runPythonRunner(
  code: string,
  workDir: string
): Promise<{
  ok: boolean;
  syntaxError?: boolean;
  error?: string;
  logs: string[];
  artifacts: Array<Record<string, unknown>>;
}> {
  return new Promise((resolve, reject) => {
    const child = spawn('python3', ['-B', PY_RUNNER_PATH], {
      cwd: __dirname,
      stdio: ['pipe', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error('Превышено время выполнения скрипта (10 сек).'));
    }, 10000);

    child.stdout.on('data', (chunk) => {
      stdout += chunk.toString('utf-8');
    });
    child.stderr.on('data', (chunk) => {
      stderr += chunk.toString('utf-8');
    });
    child.on('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });
    child.on('close', (exitCode) => {
      clearTimeout(timer);
      if (!stdout.trim()) {
        reject(new Error(stderr.trim() || `Python завершился с кодом ${exitCode}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch {
        reject(new Error(`Некорректный вывод Python-раннера: ${stdout.slice(0, 300)}`));
      }
    });
    child.stdin.write(JSON.stringify({ code, workDir }));
    child.stdin.end();
  });
}

async function runJsVmFallback(code: string) {
  const logs: string[] = [];
  const createdFiles: Array<{ filename: string; downloadUrl: string; message: string }> = [];
  const appendLog = (msg: string) => logs.push(msg);

  const context = vm.createContext({
    create_docx: async (
      title: string,
      text: string,
      list_items: string,
      table_csv: string,
      filename: string
    ) => {
      const r = await createDocx(title, text, list_items, table_csv, filename);
      createdFiles.push({
        filename: r.stored.filename,
        downloadUrl: `/api/download/${r.stored.id}`,
        message: r.message,
      });
      return r.message;
    },
    create_pptx: async (title: string, text: string, list_items: string, filename: string) => {
      const r = await createPptx(title, text, list_items, filename);
      createdFiles.push({
        filename: r.stored.filename,
        downloadUrl: `/api/download/${r.stored.id}`,
        message: r.message,
      });
      return r.message;
    },
    create_xlsx: async (table_csv: string, filename: string) => {
      const r = await createXlsx(table_csv, filename);
      createdFiles.push({
        filename: r.stored.filename,
        downloadUrl: `/api/download/${r.stored.id}`,
        message: r.message,
      });
      return r.message;
    },
    create_pdf: async (title: string, text: string, list_items: string, filename: string) => {
      const r = await createPdf(title, text, list_items, filename);
      createdFiles.push({
        filename: r.stored.filename,
        downloadUrl: `/api/download/${r.stored.id}`,
        message: r.message,
      });
      return r.message;
    },
    create_txt: async (
      title: string,
      text: string,
      list_items: string,
      table_csv: string,
      filename: string
    ) => {
      const r = await createTxt(title, text, list_items, table_csv, filename);
      createdFiles.push({
        filename: r.stored.filename,
        downloadUrl: `/api/download/${r.stored.id}`,
        message: r.message,
      });
      return r.message;
    },
    create_md: async (
      title: string,
      text: string,
      list_items: string,
      table_csv: string,
      filename: string
    ) => {
      const r = await createMd(title, text, list_items, table_csv, filename);
      createdFiles.push({
        filename: r.stored.filename,
        downloadUrl: `/api/download/${r.stored.id}`,
        message: r.message,
      });
      return r.message;
    },
    create_html: async (
      title: string,
      text: string,
      list_items: string,
      table_csv: string,
      filename: string
    ) => {
      const r = await createHtml(title, text, list_items, table_csv, filename);
      createdFiles.push({
        filename: r.stored.filename,
        downloadUrl: `/api/download/${r.stored.id}`,
        message: r.message,
      });
      return r.message;
    },
    create_from_text: async (
      raw_text: string,
      format: string = 'docx',
      filename: string = 'document'
    ) => {
      const layout = parseLayoutFromText(String(raw_text ?? ''));
      const r = await generateByFormat(format, layout, filename);
      createdFiles.push({
        filename: r.stored.filename,
        downloadUrl: `/api/download/${r.stored.id}`,
        message: r.message,
      });
      return r.message;
    },
    parse_layout: (raw_text: string) => parseLayoutFromText(String(raw_text ?? '')),
    save_file_to_download: async (filename: string, content: string | Buffer | Uint8Array) => {
      const buf = Buffer.isBuffer(content)
        ? content
        : content instanceof Uint8Array
        ? Buffer.from(content)
        : Buffer.from(String(content), 'utf-8');
      const stored = storeGeneratedFile(filename, 'application/octet-stream', buf);
      const msg = `Файл ${stored.filename} сохранен в Downloads`;
      createdFiles.push({
        filename: stored.filename,
        downloadUrl: `/api/download/${stored.id}`,
        message: msg,
      });
      return `/Downloads/${stored.filename}`;
    },
    print: (...args: unknown[]) => appendLog(`[PRINT] ${args.map((a) => String(a)).join(' ')}`),
  });

  const transformed = preprocessScript(code);
  const wrapped = `(async () => {\n${transformed}\n})()`;
  const execution = vm.runInContext(wrapped, context, { timeout: 5000 });
  await execution;

  for (const file of createdFiles) {
    appendLog(file.message);
  }
  appendLog('✅ Скрипт выполнен успешно.');
  return { ok: true, logs, createdFiles };
}

async function executeScriptUnified(code: string): Promise<{
  ok: boolean;
  logs: string[];
  createdFiles: Array<{ filename: string; downloadUrl: string; message: string }>;
}> {
  const logs: string[] = [];
  const createdFiles: Array<{ filename: string; downloadUrl: string; message: string }> = [];
  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pydocx_web_'));

  try {
    if (fs.existsSync(PY_RUNNER_PATH)) {
      const pyRes = await runPythonRunner(code, workDir);
      if (pyRes.syntaxError) {
        // Fallback to JS VM only when input looks like JS instead of broken Python
        const looksLikeJs =
          /(?:^|\n)\s*(?:const|let|var|function|async\s+function)\b|\/\/|;\s*$/m.test(code);
        if (looksLikeJs) {
          return await runJsVmFallback(code);
        }
        logs.push(...(pyRes.logs || []));
        logs.push(`❌ ${pyRes.error || 'Синтаксическая ошибка в скрипте.'}`);
        return { ok: false, logs, createdFiles };
      }

      logs.push(...(pyRes.logs || []));
      for (const art of pyRes.artifacts || []) {
        const kind = String(art.kind || '');
        if (kind === 'docx_ir') {
          const r = await renderDocxFromIr(
            art.spec as DocxIrSpec,
            String(art.filename || 'document.docx')
          );
          createdFiles.push({
            filename: r.stored.filename,
            downloadUrl: `/api/download/${r.stored.id}`,
            message: r.message,
          });
        } else if (kind === 'pptx_ir') {
          const r = await renderPptxFromIr(
            (art.slides as Array<{ title: string; bodyLines: string[] }>) || [],
            String(art.filename || 'presentation.pptx')
          );
          createdFiles.push({
            filename: r.stored.filename,
            downloadUrl: `/api/download/${r.stored.id}`,
            message: r.message,
          });
        } else if (kind === 'xlsx_ir') {
          const r = await renderXlsxFromIr(
            (art.rows as string[][]) || [],
            String(art.filename || 'table.xlsx')
          );
          createdFiles.push({
            filename: r.stored.filename,
            downloadUrl: `/api/download/${r.stored.id}`,
            message: r.message,
          });
        } else if (kind === 'pdf_ir') {
          const lines = (art.lines as string[]) || [];
          const r = await createPdf(
            lines[0] || '',
            lines.slice(1).join('\n'),
            '',
            String(art.filename || 'document.pdf')
          );
          createdFiles.push({
            filename: r.stored.filename,
            downloadUrl: `/api/download/${r.stored.id}`,
            message: r.message,
          });
        } else if (kind === 'helper') {
          const r = await generateByFormat(
            String(art.format || 'docx'),
            {
              title: String(art.title || ''),
              text: String(art.text || ''),
              listItems: String(art.listItems || ''),
              tableCsv: String(art.tableCsv || ''),
            },
            String(art.filename || 'document.docx')
          );
          createdFiles.push({
            filename: r.stored.filename,
            downloadUrl: `/api/download/${r.stored.id}`,
            message: r.message,
          });
        }
      }

      // Collect any direct files written inside workDir (e.g. via open() or save_file())
      const recordedNames = new Set(createdFiles.map((f) => f.filename));
      for (const name of fs.readdirSync(workDir)) {
        const full = path.join(workDir, name);
        if (fs.statSync(full).isFile() && !recordedNames.has(name)) {
          const buf = fs.readFileSync(full);
          const stored = storeGeneratedFile(name, 'application/octet-stream', buf);
          const msg = `Файл ${stored.filename} сохранен в Downloads`;
          createdFiles.push({
            filename: stored.filename,
            downloadUrl: `/api/download/${stored.id}`,
            message: msg,
          });
          recordedNames.add(stored.filename);
        }
      }

      for (const file of createdFiles) {
        logs.push(file.message);
      }
      if (pyRes.ok) {
        logs.push('✅ Скрипт выполнен успешно.');
      }
      return { ok: Boolean(pyRes.ok), logs, createdFiles };
    }

    return await runJsVmFallback(code);
  } catch (err) {
    const errorText = err instanceof Error ? `${err.message}\n${err.stack ?? ''}` : String(err);
    logs.push(`❌ Ошибка: ${errorText}`);
    return { ok: false, logs, createdFiles };
  } finally {
    fs.rmSync(workDir, { recursive: true, force: true });
  }
}

async function startServer() {
  const app = express();
  app.use(express.json({ limit: '5mb' }));

  app.get('/api/health', (_req, res) => {
    res.json({ status: 'ok', app: 'PyDocx' });
  });

  app.post('/api/generate', async (req, res) => {
    try {
      const { format, title = '', text = '', listItems = '', tableCsv = '', filename = '' } =
        req.body ?? {};
      const result = await generateByFormat(
        String(format || ''),
        {
          title: String(title ?? ''),
          text: String(text ?? ''),
          listItems: String(listItems ?? ''),
          tableCsv: String(tableCsv ?? ''),
        },
        String(filename || '')
      );

      res.json({
        ok: true,
        message: result.message,
        filename: result.stored.filename,
        downloadUrl: `/api/download/${result.stored.id}`,
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      res.status(400).json({ ok: false, error: message });
    }
  });

  app.post('/api/typeset-text', async (req, res) => {
    try {
      const { content = '', sourceFilename = '', format = 'docx', filename = '' } = req.body ?? {};
      const rawText = String(content ?? '');
      if (!rawText.trim()) {
        res.status(400).json({ ok: false, error: 'Текстовый файл пуст.' });
        return;
      }

      if (looksLikeExecutableScript(rawText, String(sourceFilename || ''))) {
        const scriptRes = await executeScriptUnified(rawText);
        if (!scriptRes.ok) {
          res.status(400).json({
            ok: false,
            error: scriptRes.logs.join('\n') || 'Ошибка выполнения скрипта.',
          });
          return;
        }
        const firstFile = scriptRes.createdFiles[0];
        res.json({
          ok: true,
          executedAsScript: true,
          logs: [
            'ℹ️ Обнаружен исполняемый Python-скрипт — выполнена генерация документа.',
            ...scriptRes.logs,
          ],
          message: firstFile?.message || 'Скрипт выполнен.',
          filename: firstFile?.filename,
          downloadUrl: firstFile?.downloadUrl,
          createdFiles: scriptRes.createdFiles,
        });
        return;
      }

      const parsed = parseLayoutFromText(rawText, String(sourceFilename || ''));
      const baseFromSource = sourceFilename
        ? path.basename(String(sourceFilename), path.extname(String(sourceFilename)))
        : 'layout';
      const targetFilename = String(filename || '').trim() || `${baseFromSource}.${format}`;
      const result = await generateByFormat(String(format || 'docx'), parsed, targetFilename);

      res.json({
        ok: true,
        message: result.message,
        filename: result.stored.filename,
        downloadUrl: `/api/download/${result.stored.id}`,
        parsed: {
          title: parsed.title,
          listCount: parsed.listItems ? parsed.listItems.split('\n').filter(Boolean).length : 0,
          tableRows: parsed.tableCsv ? parseCsv(parsed.tableCsv).length : 0,
        },
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      res.status(400).json({ ok: false, error: message });
    }
  });

  app.post('/api/run-script', async (req, res) => {
    const code = String(req.body?.code ?? '');
    const result = await executeScriptUnified(code);
    res.status(200).json(result);
  });

  app.get('/api/download/:id', (req, res) => {
    const entry = generatedFiles.get(req.params.id);
    if (!entry) {
      res.status(404).json({ error: 'Файл не найден или срок хранения истек.' });
      return;
    }
    const encodedName = encodeURIComponent(entry.filename);
    res.setHeader('Content-Type', entry.mimeType);
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="document"; filename*=UTF-8''${encodedName}`
    );
    res.send(entry.buffer);
  });

  const isProd = process.env.NODE_ENV === 'production';
  if (!isProd) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.join(__dirname, 'dist');
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distDir, 'index.html'));
    });
  }

  const port = Number(process.env.PORT) || 3000;
  app.listen(port, '0.0.0.0', () => {
    console.log(`DocBuilder Mobile listening on http://0.0.0.0:${port}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
