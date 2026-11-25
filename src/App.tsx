import { useState, useRef, useEffect, type ChangeEvent, type DragEvent } from 'react';

type DocFormat = 'docx' | 'pptx' | 'xlsx' | 'pdf' | 'txt' | 'md' | 'html';
type TabKey = 'script' | 'layout' | 'settings';

interface GeneratedFileLink {
  filename: string;
  downloadUrl: string;
}

interface PyDocxSettings {
  defaultDir: string;
  saveNextToSource: boolean;
}

const SETTINGS_STORAGE_KEY = 'pydocx_settings_v1';
const DEFAULT_SAVE_DIR = 'Documents/PyDocx';

const DEFAULT_SCRIPT = `# Пример:
# create_docx('Заголовок', 'Текст', 'Пункт 1\\nПункт 2', 'A,B\\n1,2', 'my.docx')
# create_from_text('# Заголовок\\n- Пункт 1\\n- Пункт 2', 'pdf', 'layout.pdf')
print('Привет из PyDocx!')`;

const SAMPLE_LAYOUT_TEXT = `# Отчет по проекту PyDocx
TITLE: Отчет по проекту PyDocx

Ниже приведены ключевые результаты работы и сводная таблица показателей.

## Основные задачи
- Поддержка верстки из текстовых файлов (.txt, .md, .json, .csv)
- Экспорт в Word, PowerPoint, Excel, PDF, Markdown и HTML
- Выполнение пользовательских скриптов Python

## Сводная таблица
| Модуль | Статус | Формат |
| --- | --- | --- |
| Свой скрипт | Готово | PY/TXT/MD/JSON |
| Текстовая верстка | Готово | DOCX/PPTX/XLSX/PDF |
| Настройки | Готово | Android/Windows |`;

const ACCEPTED_SCRIPT_EXTENSIONS =
  '.py,.txt,.md,.json,.csv,.yaml,.yml,.js,.log,text/*';

const ACCEPTED_LAYOUT_EXTENSIONS =
  '.txt,.md,.py,.js,.ts,.json,.csv,.yaml,.yml,.ini,.log,.rst,text/*';

function triggerDownload(url: string, filename: string) {
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

async function saveFileWithPicker(url: string, filename: string): Promise<string> {
  const win = window as unknown as {
    showSaveFilePicker?: (opts?: { suggestedName?: string }) => Promise<{
      createWritable: () => Promise<{
        write: (data: Blob) => Promise<void>;
        close: () => Promise<void>;
      }>;
    }>;
  };
  if (typeof win.showSaveFilePicker === 'function') {
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error('Не удалось получить файл для сохранения.');
    }
    const blob = await response.blob();
    const handle = await win.showSaveFilePicker({ suggestedName: filename });
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return `Файл ${filename} сохранён через проводник.`;
  }
  triggerDownload(url, filename);
  return `Файл ${filename} отправлен в загрузки браузера.`;
}

function stripExt(name: string): string {
  const idx = name.lastIndexOf('.');
  return idx > 0 ? name.slice(0, idx) : name;
}

function loadSettings(): PyDocxSettings {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw) as Partial<PyDocxSettings>;
      return {
        defaultDir:
          typeof parsed.defaultDir === 'string' && parsed.defaultDir.trim()
            ? parsed.defaultDir.trim()
            : DEFAULT_SAVE_DIR,
        saveNextToSource: parsed.saveNextToSource ?? true,
      };
    }
  } catch {
    // Ignore storage errors
  }
  return {
    defaultDir: DEFAULT_SAVE_DIR,
    saveNextToSource: true,
  };
}

export default function App() {
  const [activeTab, setActiveTab] = useState<TabKey>('script');
  const [settings, setSettings] = useState<PyDocxSettings>(() => loadSettings());
  const [defaultDirInput, setDefaultDirInput] = useState(settings.defaultDir);
  const [settingsLog, setSettingsLog] = useState('');

  // Script tab state (First / Main tab)
  const [scriptCode, setScriptCode] = useState(DEFAULT_SCRIPT);
  const [scriptSourceName, setScriptSourceName] = useState('');
  const [scriptLog, setScriptLog] = useState('');
  const [scriptDownloads, setScriptDownloads] = useState<GeneratedFileLink[]>([]);
  const [isRunningScript, setIsRunningScript] = useState(false);
  const scriptFileInputRef = useRef<HTMLInputElement | null>(null);

  // Text-file layout tab state (Second tab)
  const [sourceFilename, setSourceFilename] = useState('');
  const [layoutContent, setLayoutContent] = useState('');
  const [layoutOutFilename, setLayoutOutFilename] = useState('layout');
  const [layoutLog, setLayoutLog] = useState('');
  const [layoutDownloads, setLayoutDownloads] = useState<GeneratedFileLink[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isTypesetting, setIsTypesetting] = useState(false);
  const layoutFileInputRef = useRef<HTMLInputElement | null>(null);

  // Snackbar state
  const [snackbar, setSnackbar] = useState<{
    message: string;
    downloadUrl?: string;
    filename?: string;
  } | null>(null);
  const snackTimerRef = useRef<number | null>(null);

  useEffect(() => {
    try {
      localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings));
    } catch {
      // Ignore storage errors
    }
  }, [settings]);

  const showSnackbar = (message: string, downloadUrl?: string, dlFilename?: string) => {
    setSnackbar({ message, downloadUrl, filename: dlFilename });
    if (snackTimerRef.current !== null) {
      window.clearTimeout(snackTimerRef.current);
    }
    snackTimerRef.current = window.setTimeout(() => {
      setSnackbar(null);
    }, 5000);
  };

  const handleSaveAsClick = async (item: GeneratedFileLink) => {
    try {
      const msg = await saveFileWithPicker(item.downloadUrl, item.filename);
      showSnackbar(msg);
    } catch (err) {
      if (err instanceof Error && err.name === 'AbortError') {
        return;
      }
      const message = err instanceof Error ? err.message : String(err);
      showSnackbar(`Ошибка сохранения: ${message}`);
    }
  };

  const handleScriptFileUpload = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const content = await file.text();
      setScriptCode(content);
      setScriptSourceName(file.name);
      setScriptLog(`Загружен файл: ${file.name}\n`);
      showSnackbar(`Файл ${file.name} загружен в редактор`);
    }
    e.target.value = '';
  };

  const handleRunScript = async () => {
    setIsRunningScript(true);
    const targetLabel =
      scriptSourceName && settings.saveNextToSource
        ? `рядом с файлом ${scriptSourceName} (на Android/Windows) или ${settings.defaultDir}`
        : settings.defaultDir;
    setScriptLog(`Выполнение... (папка назначения: ${targetLabel})\n`);
    setScriptDownloads([]);

    try {
      const response = await fetch('/api/run-script', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ code: scriptCode }),
      });
      const data = await response.json();
      const output = Array.isArray(data.logs) ? data.logs.join('\n') + '\n' : '';
      setScriptLog((prev) => `${prev}${output}`);

      if (Array.isArray(data.createdFiles) && data.createdFiles.length > 0) {
        setScriptDownloads(data.createdFiles);
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setScriptLog((prev) => `${prev}❌ Ошибка: ${message}\n`);
    } finally {
      setIsRunningScript(false);
    }
  };

  const loadFileIntoLayout = async (file: File) => {
    const content = await file.text();
    setSourceFilename(file.name);
    setLayoutContent(content);
    setLayoutOutFilename(stripExt(file.name) || 'layout');
    setLayoutLog(`Загружен файл: ${file.name} (${content.length} симв.)\n`);
    showSnackbar(`Файл ${file.name} загружен`);
  };

  const handleLayoutFileChange = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      await loadFileIntoLayout(file);
    }
    e.target.value = '';
  };

  const handleLayoutDrop = async (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      await loadFileIntoLayout(file);
    }
  };

  const handleTypesetFromText = async (format: DocFormat) => {
    setIsTypesetting(true);
    setLayoutDownloads([]);
    try {
      const baseName = stripExt(layoutOutFilename.trim() || 'layout');
      const targetName = `${baseName}.${format}`;
      const response = await fetch('/api/typeset-text', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          content: layoutContent,
          sourceFilename,
          format,
          filename: targetName,
        }),
      });
      const data = await response.json();
      if (!response.ok || !data.ok) {
        throw new Error(data.error || 'Не удалось сверстать документ.');
      }

      if (data.executedAsScript) {
        const logsText = Array.isArray(data.logs) && data.logs.length > 0
          ? data.logs.join('\n')
          : data.message || 'Скрипт выполнен.';
        setLayoutLog(`${logsText}\n`);
        const files: GeneratedFileLink[] =
          Array.isArray(data.createdFiles) && data.createdFiles.length > 0
            ? data.createdFiles
            : data.filename && data.downloadUrl
              ? [{ filename: data.filename, downloadUrl: data.downloadUrl }]
              : [];
        setLayoutDownloads(files);
        if (files.length > 0) {
          showSnackbar(`Готово: ${files[0].filename}`, files[0].downloadUrl, files[0].filename);
        } else {
          showSnackbar(data.message || 'Скрипт выполнен');
        }
      } else {
        const summary = data.parsed
          ? `Заголовок: "${data.parsed.title || '—'}", пунктов списка: ${data.parsed.listCount}, строк таблицы: ${data.parsed.tableRows}`
          : '';
        setLayoutLog(`${data.message}\n${summary}\n`);
        const fileItem = { filename: data.filename, downloadUrl: data.downloadUrl };
        setLayoutDownloads([fileItem]);
        showSnackbar(`Готово: ${data.filename}`, data.downloadUrl, data.filename);
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setLayoutLog((prev) => `${prev}❌ Ошибка: ${message}\n`);
      showSnackbar(`Ошибка: ${message}`);
    } finally {
      setIsTypesetting(false);
    }
  };

  const handleRunLayoutAsScript = async () => {
    if (!layoutContent.trim()) {
      showSnackbar('Загрузите или введите текст скрипта.');
      return;
    }
    setIsTypesetting(true);
    setLayoutLog('Выполнение как скрипт...\n');
    setLayoutDownloads([]);
    try {
      const response = await fetch('/api/run-script', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ code: layoutContent }),
      });
      const data = await response.json();
      const output = Array.isArray(data.logs) ? data.logs.join('\n') + '\n' : '';
      setLayoutLog(output);

      if (Array.isArray(data.createdFiles) && data.createdFiles.length > 0) {
        setLayoutDownloads(data.createdFiles);
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setLayoutLog((prev) => `${prev}❌ Ошибка: ${message}\n`);
    } finally {
      setIsTypesetting(false);
    }
  };

  const handleSaveDefaultDir = () => {
    const trimmed = defaultDirInput.trim() || DEFAULT_SAVE_DIR;
    setDefaultDirInput(trimmed);
    setSettings((prev) => ({ ...prev, defaultDir: trimmed }));
    setSettingsLog(`Папка по умолчанию сохранена: ${trimmed}`);
    showSnackbar('Настройки сохранены');
  };

  const handleResetDefaultDir = () => {
    setDefaultDirInput(DEFAULT_SAVE_DIR);
    setSettings((prev) => ({ ...prev, defaultDir: DEFAULT_SAVE_DIR }));
    setSettingsLog(`Восстановлен путь по умолчанию: ${DEFAULT_SAVE_DIR}`);
    showSnackbar('Восстановлена папка Documents/PyDocx');
  };

  const handleRequestStoragePermission = () => {
    const msg =
      'В веб-версии сохранение выполняется через проводник браузера; на Android разрешение запрашивается системным диалогом.';
    setSettingsLog(msg);
    showSnackbar('Готово к сохранению файлов.');
  };

  return (
    <div className="app-shell">
      {/* Desktop Windows-style NavigationRail + Mobile Sticky TabBar */}
      <aside className="desktop-rail" aria-label="Навигация PyDocx">
        <div className="rail-brand">
          <img src="/assets/aistudio/icon_5e6b4422.png" alt="PyDocx" className="brand-icon" />
          <span className="brand-name">PyDocx</span>
        </div>
        <button
          type="button"
          className={`rail-item ${activeTab === 'script' ? 'active' : ''}`}
          onClick={() => setActiveTab('script')}
        >
          Свой скрипт
        </button>
        <button
          type="button"
          className={`rail-item ${activeTab === 'layout' ? 'active' : ''}`}
          onClick={() => setActiveTab('layout')}
        >
          Верстка из файла
        </button>
        <button
          type="button"
          className={`rail-item ${activeTab === 'settings' ? 'active' : ''}`}
          onClick={() => setActiveTab('settings')}
        >
          Настройки
        </button>
      </aside>

      <div className="workspace-column">
        <nav className="tabs-bar" role="tablist">
          <div className="mobile-brand">
            <img src="/assets/aistudio/icon_5e6b4422.png" alt="PyDocx" className="brand-icon-sm" />
            <span>PyDocx</span>
          </div>
          <div className="tabs-strip">
            <button
              type="button"
              role="tab"
              aria-selected={activeTab === 'script'}
              className={`tab-btn ${activeTab === 'script' ? 'active' : ''}`}
              onClick={() => setActiveTab('script')}
            >
              Свой скрипт
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={activeTab === 'layout'}
              className={`tab-btn ${activeTab === 'layout' ? 'active' : ''}`}
              onClick={() => setActiveTab('layout')}
            >
              Верстка из файла
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={activeTab === 'settings'}
              className={`tab-btn ${activeTab === 'settings' ? 'active' : ''}`}
              onClick={() => setActiveTab('settings')}
            >
              Настройки
            </button>
          </div>
        </nav>

        {activeTab === 'script' && (
          <main className="tab-content split-layout">
            <section className="pane-card pane-editor">
              <h1 className="section-title">Свой скрипт</h1>

              <div className="source-banner">
                <span>
                  {scriptSourceName
                    ? `Файл: ${scriptSourceName} • ${
                        settings.saveNextToSource
                          ? 'сохранение рядом с файлом'
                          : `папка: ${settings.defaultDir}`
                      }`
                    : `Режим: ручной ввод / вставка • папка по умолчанию: ${settings.defaultDir}`}
                </span>
                {scriptSourceName && (
                  <button
                    type="button"
                    className="inline-reset-btn"
                    onClick={() => setScriptSourceName('')}
                  >
                    Сбросить файл
                  </button>
                )}
              </div>

              <input
                ref={scriptFileInputRef}
                type="file"
                accept={ACCEPTED_SCRIPT_EXTENSIONS}
                className="hidden-file-input"
                onChange={handleScriptFileUpload}
              />

              <div className="buttons-row">
                <button
                  type="button"
                  className="elevated-btn"
                  disabled={isRunningScript}
                  onClick={handleRunScript}
                >
                  Выполнить
                </button>
                <button
                  type="button"
                  className="outlined-btn"
                  onClick={() => scriptFileInputRef.current?.click()}
                >
                  Загрузить файл (.py, .txt, .md, .json, .csv, .yaml, .js)
                </button>
                <button
                  type="button"
                  className="outlined-btn"
                  onClick={() => {
                    setScriptCode('');
                    setScriptSourceName('');
                  }}
                >
                  Очистить
                </button>
              </div>

              <div className="field-group grow-field">
                <label className="field-label" htmlFor="script-code-field">
                  Код скрипта (Python / JS / Markdown-блоки)
                </label>
                <textarea
                  id="script-code-field"
                  className="field-textarea code"
                  rows={14}
                  placeholder="Доступны: create_docx, create_pptx, create_xlsx, create_pdf, create_from_text, create_txt, create_md, create_html"
                  value={scriptCode}
                  onChange={(e) => setScriptCode(e.target.value)}
                />
              </div>
            </section>

            <section className="pane-card pane-output">
              <p className="log-label">Лог выполнения и экспорт</p>
              <pre className="log-output grow-log">{scriptLog}</pre>

              {scriptDownloads.length > 0 && (
                <div className="downloads-list">
                  {scriptDownloads.map((item) => (
                    <button
                      key={item.downloadUrl}
                      type="button"
                      className="save-as-btn"
                      onClick={() => handleSaveAsClick(item)}
                    >
                      Сохранить {item.filename} в...
                    </button>
                  ))}
                </div>
              )}
            </section>
          </main>
        )}

        {activeTab === 'layout' && (
          <main className="tab-content split-layout">
            <section className="pane-card pane-editor">
              <h1 className="section-title">Верстка из текстовых файлов и скриптов</h1>

              <input
                ref={layoutFileInputRef}
                type="file"
                accept={ACCEPTED_LAYOUT_EXTENSIONS}
                className="hidden-file-input"
                onChange={handleLayoutFileChange}
              />

              <div
                className={`dropzone ${isDragging ? 'drag-over' : ''}`}
                onClick={() => layoutFileInputRef.current?.click()}
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsDragging(true);
                }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={handleLayoutDrop}
                role="button"
                tabIndex={0}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    layoutFileInputRef.current?.click();
                  }
                }}
              >
                <p className="dropzone-title">
                  {sourceFilename
                    ? `Загружен файл: ${sourceFilename}`
                    : 'Нажмите или перетащите текстовый файл / скрипт'}
                </p>
                <p className="dropzone-sub">
                  Поддерживаются: .txt, .md, .py, .js, .json, .csv, .yaml
                </p>
              </div>

              <div className="buttons-row">
                <button
                  type="button"
                  className="outlined-btn"
                  onClick={() => layoutFileInputRef.current?.click()}
                >
                  Выбрать файл (.txt, .md, .py, .json, .csv, .yaml)
                </button>
                <button
                  type="button"
                  className="outlined-btn"
                  onClick={() => {
                    setSourceFilename('example.md');
                    setLayoutContent(SAMPLE_LAYOUT_TEXT);
                    setLayoutOutFilename('report');
                    setLayoutLog('Загружен пример текстовой разметки.\n');
                  }}
                >
                  Загрузить пример разметки
                </button>
              </div>

              <div className="field-group grow-field">
                <label className="field-label" htmlFor="layout-content-field">
                  Содержимое файла (разметка: # заголовок, - список, | таблица |, JSON, CSV или код скрипта)
                </label>
                <textarea
                  id="layout-content-field"
                  className="field-textarea code"
                  rows={10}
                  placeholder="Вставьте разметку, CSV, JSON или загрузите файл выше..."
                  value={layoutContent}
                  onChange={(e) => setLayoutContent(e.target.value)}
                />
              </div>
            </section>

            <section className="pane-card pane-output">
              <div className="field-group">
                <label className="field-label" htmlFor="layout-filename-field">
                  Имя итогового файла (без расширения)
                </label>
                <input
                  id="layout-filename-field"
                  className="field-input"
                  type="text"
                  value={layoutOutFilename}
                  onChange={(e) => setLayoutOutFilename(e.target.value)}
                />
              </div>

              <div className="buttons-row">
                <button
                  type="button"
                  className="elevated-btn"
                  disabled={isTypesetting}
                  onClick={() => handleTypesetFromText('docx')}
                >
                  Word (.docx)
                </button>
                <button
                  type="button"
                  className="elevated-btn"
                  disabled={isTypesetting}
                  onClick={() => handleTypesetFromText('pdf')}
                >
                  PDF (.pdf)
                </button>
                <button
                  type="button"
                  className="elevated-btn"
                  disabled={isTypesetting}
                  onClick={() => handleTypesetFromText('pptx')}
                >
                  PowerPoint (.pptx)
                </button>
                <button
                  type="button"
                  className="elevated-btn"
                  disabled={isTypesetting}
                  onClick={() => handleTypesetFromText('xlsx')}
                >
                  Excel (.xlsx)
                </button>
                <button
                  type="button"
                  className="elevated-btn"
                  disabled={isTypesetting}
                  onClick={() => handleTypesetFromText('md')}
                >
                  Markdown (.md)
                </button>
                <button
                  type="button"
                  className="elevated-btn"
                  disabled={isTypesetting}
                  onClick={() => handleTypesetFromText('html')}
                >
                  HTML (.html)
                </button>
                <button
                  type="button"
                  className="elevated-btn"
                  disabled={isTypesetting}
                  onClick={() => handleTypesetFromText('txt')}
                >
                  Текст (.txt)
                </button>
                <button
                  type="button"
                  className="outlined-btn"
                  disabled={isTypesetting}
                  onClick={handleRunLayoutAsScript}
                >
                  Выполнить как скрипт
                </button>
              </div>

              <hr className="divider" />

              <p className="log-label">Результат верстки:</p>
              <pre className="log-output grow-log">{layoutLog}</pre>

              {layoutDownloads.length > 0 && (
                <div className="downloads-list">
                  {layoutDownloads.map((item) => (
                    <button
                      key={item.downloadUrl}
                      type="button"
                      className="save-as-btn"
                      onClick={() => handleSaveAsClick(item)}
                    >
                      Сохранить {item.filename} в...
                    </button>
                  ))}
                </div>
              )}
            </section>
          </main>
        )}

        {activeTab === 'settings' && (
          <main className="tab-content settings-layout">
            <section className="pane-card settings-card">
              <h1 className="section-title">Настройки PyDocx</h1>

              <div className="field-group">
                <label className="field-label" htmlFor="default-dir-field">
                  Папка сохранения по умолчанию
                </label>
                <input
                  id="default-dir-field"
                  className="field-input"
                  type="text"
                  placeholder="Documents/PyDocx"
                  value={defaultDirInput}
                  onChange={(e) => setDefaultDirInput(e.target.value)}
                />
                <p className="hint-text">
                  По умолчанию из коробки: <code>Documents/PyDocx</code>
                </p>
              </div>

              <div className="buttons-row">
                <button type="button" className="elevated-btn" onClick={handleSaveDefaultDir}>
                  Сохранить путь
                </button>
                <button type="button" className="outlined-btn" onClick={handleResetDefaultDir}>
                  Сбросить (Documents/PyDocx)
                </button>
              </div>

              <hr className="divider" />

              <label className="switch-row">
                <input
                  type="checkbox"
                  checked={settings.saveNextToSource}
                  onChange={(e) =>
                    setSettings((prev) => ({ ...prev, saveNextToSource: e.target.checked }))
                  }
                />
                <span>Сохранять рядом с исходным файлом (если открыт файл скрипта)</span>
              </label>
              <p className="hint-text">
                Если открыт файл с диска — результат сохраняется в его исходную папку. Если код
                вставлен вручную — сохраняется в папку по умолчанию ({settings.defaultDir}).
              </p>

              <hr className="divider" />

              <p className="log-label">Разрешения на запись (Android)</p>
              <p className="hint-text">
                Для прямого сохранения в Documents/PyDocx на Android требуется разрешение на доступ
                к памяти. Кнопка «Сохранить ... в...» работает через системный проводник на любой
                платформе.
              </p>
              <div className="buttons-row">
                <button
                  type="button"
                  className="elevated-btn"
                  onClick={handleRequestStoragePermission}
                >
                  Запросить разрешение на запись
                </button>
              </div>

              {settingsLog && <p className="hint-text">{settingsLog}</p>}
            </section>
          </main>
        )}
      </div>

      {snackbar && (
        <div className="snackbar" role="status">
          <span>{snackbar.message}</span>
          {snackbar.downloadUrl && snackbar.filename && (
            <button
              type="button"
              className="snackbar-action"
              onClick={() =>
                handleSaveAsClick({
                  downloadUrl: snackbar.downloadUrl!,
                  filename: snackbar.filename!,
                })
              }
            >
              Сохранить в...
            </button>
          )}
        </div>
      )}
    </div>
  );
}
