const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const progressWrapper = document.getElementById('progressWrapper');
const progressFill = document.getElementById('progressFill');
const progressValue = document.getElementById('progressValue');
const progressLabel = document.getElementById('progressLabel');
const errorBox = document.getElementById('errorBox');
const searchInput = document.getElementById('searchInput');
const searchButton = document.getElementById('searchButton');
const resetButton = document.getElementById('resetButton');
const caseSensitiveCheckbox = document.getElementById('caseSensitive');
const exactMatchCheckbox = document.getElementById('exactMatch');
const resultsSummary = document.getElementById('resultsSummary');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');
const pageSizeSelect = document.getElementById('pageSize');
const prevPageButton = document.getElementById('prevPage');
const nextPageButton = document.getElementById('nextPage');
const pageInfo = document.getElementById('pageInfo');
const copyButton = document.getElementById('copyButton');
const exportCsvButton = document.getElementById('exportCsvButton');
const exportXlsxButton = document.getElementById('exportXlsxButton');

const worker = new Worker('searchWorker.js', { type: 'module' });
const textDecoder = new TextDecoder();
const textEncoder = new TextEncoder();

let dataset = [];
let headers = [];
let filteredIndexes = [];
let currentPage = 1;
let pageSize = Number(pageSizeSelect.value);
let sortState = { columnIndex: null, direction: 'asc' };
let lastSearchDuration = 0;
let activeFileName = '';

worker.onmessage = (event) => {
    const { type } = event.data;
    if (type === 'ready') {
        setControlsEnabled(true);
        updateSummary();
    } else if (type === 'searchResults') {
        filteredIndexes = event.data.matches;
        lastSearchDuration = event.data.duration;
        currentPage = 1;
        sortState = { columnIndex: null, direction: 'asc' };
        updateSummary();
        renderTable();
        toggleExports(filteredIndexes.length > 0);
        searchButton.disabled = false;
        searchButton.textContent = 'Rechercher';
    } else if (type === 'error') {
        showError(event.data.message || 'Une erreur est survenue dans le moteur de recherche.');
        searchButton.disabled = false;
        searchButton.textContent = 'Rechercher';
    }
};

worker.onerror = (err) => {
    console.error('Erreur worker', err);
    showError('Le moteur de recherche a rencontré une erreur. Consultez la console pour plus de détails.');
};

function setControlsEnabled(enabled) {
    searchButton.disabled = !enabled;
    resetButton.disabled = !enabled;
    copyButton.disabled = !enabled;
    exportCsvButton.disabled = !enabled;
    exportXlsxButton.disabled = !enabled;
    prevPageButton.disabled = !enabled;
    nextPageButton.disabled = !enabled;
    pageSizeSelect.disabled = !enabled;
}

function showError(message) {
    errorBox.textContent = message;
    errorBox.hidden = !message;
}

function resetProgress() {
    progressWrapper.hidden = true;
    progressFill.style.width = '0%';
    progressValue.textContent = '0%';
    progressLabel.textContent = 'Analyse du fichier…';
}

function updateProgress(percent, label) {
    progressWrapper.hidden = false;
    const clamped = Math.max(0, Math.min(100, percent || 0));
    progressFill.style.width = `${clamped.toFixed(1)}%`;
    progressValue.textContent = `${clamped.toFixed(1)}%`;
    if (label) {
        progressLabel.textContent = label;
    }
}

dropZone.addEventListener('dragover', (event) => {
    event.preventDefault();
    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (event) => {
    event.preventDefault();
    dropZone.classList.remove('dragover');
    if (event.dataTransfer?.files?.length) {
        handleFile(event.dataTransfer.files[0]);
    }
});

fileInput.addEventListener('change', (event) => {
    const file = event.target.files?.[0];
    if (file) {
        handleFile(file);
    }
});

searchButton.addEventListener('click', () => {
    executeSearch();
});

searchInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
        event.preventDefault();
        executeSearch();
    }
});

resetButton.addEventListener('click', () => {
    searchInput.value = '';
    caseSensitiveCheckbox.checked = false;
    exactMatchCheckbox.checked = false;
    filteredIndexes = dataset.map((_, idx) => idx);
    currentPage = 1;
    sortState = { columnIndex: null, direction: 'asc' };
    lastSearchDuration = 0;
    updateSummary();
    renderTable();
});

pageSizeSelect.addEventListener('change', () => {
    pageSize = Number(pageSizeSelect.value);
    currentPage = 1;
    renderTable();
});

prevPageButton.addEventListener('click', () => {
    if (currentPage > 1) {
        currentPage -= 1;
        renderTable();
    }
});

nextPageButton.addEventListener('click', () => {
    const totalPages = Math.ceil(filteredIndexes.length / pageSize) || 1;
    if (currentPage < totalPages) {
        currentPage += 1;
        renderTable();
    }
});

copyButton.addEventListener('click', async () => {
    try {
        const exportRows = getFilteredRows();
        if (!exportRows.length) {
            return;
        }
        const csv = window.Papa.unparse({ fields: headers, data: exportRows });
        await navigator.clipboard.writeText(csv);
        displayMessage('Résultats copiés dans le presse-papier.');
    } catch (error) {
        showError(`Impossible de copier : ${error.message}`);
    }
});

exportCsvButton.addEventListener('click', () => {
    const exportRows = getFilteredRows();
    if (!exportRows.length) {
        return;
    }
    const csv = window.Papa.unparse({ fields: headers, data: exportRows });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    triggerDownload(url, deriveFileName('csv'));
});

exportXlsxButton.addEventListener('click', async () => {
    const exportRows = getFilteredRows();
    if (!exportRows.length) {
        return;
    }
    try {
        const blob = await buildXlsxBlob(headers, exportRows, deriveSheetName());
        const url = URL.createObjectURL(blob);
        triggerDownload(url, deriveFileName('xlsx'));
    } catch (error) {
        showError(`Export XLSX impossible : ${error.message}`);
    }
});

function triggerDownload(url, filename) {
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
}

function deriveFileName(extension) {
    const base = activeFileName ? activeFileName.replace(/\.[^.]+$/, '') : 'export';
    return `${base}-filtre.${extension}`;
}

function deriveSheetName() {
    const base = activeFileName ? activeFileName.replace(/\.[^.]+$/, '') : 'Résultats';
    return sanitizeSheetName(base);
}

function displayMessage(message) {
    resultsSummary.textContent = `${message}${lastSearchDuration ? ` · Recherche ${lastSearchDuration.toFixed(1)} ms` : ''}`;
    resultsSummary.hidden = false;
}

function getFilteredRows() {
    return filteredIndexes.map((idx) => dataset[idx]);
}

async function handleFile(file) {
    resetProgress();
    showError('');
    setControlsEnabled(false);
    searchButton.textContent = 'Rechercher';

    try {
        validateFile(file);
    } catch (error) {
        showError(error.message);
        return;
    }

    activeFileName = file.name;
    updateProgress(0, 'Préparation du fichier…');

    try {
        const result = await parseFile(file);
        headers = result.headers;
        dataset = result.rows;
        filteredIndexes = dataset.map((_, idx) => idx);
        currentPage = 1;
        renderHeaders();
        renderTable();
        updateSummary();
        toggleExports(dataset.length > 0);
        worker.postMessage({ type: 'initData', headers, rows: dataset });
    } catch (error) {
        console.error(error);
        showError(`Échec de l'import : ${error.message}`);
    } finally {
        resetProgress();
        fileInput.value = '';
    }
}

function validateFile(file) {
    const maxSize = 1024 * 1024 * 600; // 600 MB
    if (file.size > maxSize) {
        throw new Error('Le fichier dépasse la taille maximale autorisée (600 Mo).');
    }
    const extension = file.name.split('.').pop()?.toLowerCase();
    if (!['csv', 'xlsx'].includes(extension || '')) {
        throw new Error('Format non pris en charge. Veuillez fournir un fichier CSV ou XLSX.');
    }
}

async function parseFile(file) {
    const extension = file.name.split('.').pop()?.toLowerCase();
    if (extension === 'csv') {
        return parseCsv(file);
    }
    return parseXlsx(file);
}

function parseCsv(file) {
    return new Promise((resolve, reject) => {
        const rows = [];
        let headers = [];
        let processedBytes = 0;
        window.Papa.parse(file, {
            header: true,
            worker: true,
            skipEmptyLines: true,
            chunkSize: 1024 * 1024 * 2,
            chunk: (results) => {
                if (!headers.length) {
                    headers = results.meta.fields || [];
                }
                const chunkRows = results.data.map((row) => headers.map((header) => sanitizeValue(row[header])));
                rows.push(...chunkRows);
                processedBytes = results.meta.cursor || processedBytes + chunkRows.length;
                const percent = (processedBytes / file.size) * 100;
                updateProgress(percent, `Lecture CSV (${rows.length.toLocaleString()} lignes)`);
            },
            error: (error) => {
                reject(error);
            },
            complete: () => {
                resolve({ headers, rows });
            }
        });
    });
}

async function parseXlsx(file) {
    updateProgress(0, 'Lecture XLSX…');
    const buffer = await file.arrayBuffer();
    const zipNavigator = createZipNavigator(buffer);
    const workbookXml = await zipNavigator.getText('xl/workbook.xml');
    if (!workbookXml) {
        throw new Error('Classeur Excel invalide (workbook manquant).');
    }
    const sheetDescriptor = extractFirstSheet(workbookXml);
    const relationshipsXml = await zipNavigator.getText('xl/_rels/workbook.xml.rels');
    const sheetPath = resolveSheetPath(sheetDescriptor.relId, relationshipsXml) || 'xl/worksheets/sheet1.xml';
    const sheetXml = await zipNavigator.getText(sheetPath);
    if (!sheetXml) {
        throw new Error('Feuille de calcul introuvable ou illisible.');
    }
    const sharedStringsXml = await zipNavigator.getText('xl/sharedStrings.xml');
    const sharedStrings = sharedStringsXml ? parseSharedStrings(sharedStringsXml) : [];
    const { headers, rows } = await extractSheetRows(sheetXml, sharedStrings);
    if (!headers.length && !rows.length) {
        throw new Error('La feuille est vide.');
    }
    return { headers, rows };
}

function sanitizeValue(value) {
    if (value === undefined || value === null) {
        return '';
    }
    if (typeof value === 'string') {
        return value.trim();
    }
    if (value instanceof Date) {
        return value.toISOString();
    }
    return String(value);
}

function renderHeaders() {
    tableHead.innerHTML = '';
    const headerRow = document.createElement('tr');
    headers.forEach((header, index) => {
        const th = document.createElement('th');
        th.textContent = header || `Colonne ${index + 1}`;
        th.dataset.index = index;
        th.addEventListener('click', () => handleSort(index));
        if (sortState.columnIndex === index) {
            th.classList.add('sorted');
            th.setAttribute('data-direction', sortState.direction);
            th.textContent = `${th.textContent} ${sortState.direction === 'asc' ? '▲' : '▼'}`;
        }
        headerRow.appendChild(th);
    });
    tableHead.appendChild(headerRow);
}

function handleSort(columnIndex) {
    if (sortState.columnIndex === columnIndex) {
        sortState.direction = sortState.direction === 'asc' ? 'desc' : 'asc';
    } else {
        sortState = { columnIndex, direction: 'asc' };
    }
    renderTable();
}

function renderTable() {
    if (!headers.length) {
        tableBody.innerHTML = '';
        tableHead.innerHTML = '';
        pageInfo.textContent = 'Page 0 / 0';
        prevPageButton.disabled = true;
        nextPageButton.disabled = true;
        toggleExports(false);
        return;
    }

    renderHeaders();
    const workingIndexes = sortState.columnIndex !== null
        ? [...filteredIndexes].sort((a, b) => compareCells(dataset[a], dataset[b], sortState))
        : [...filteredIndexes];

    const totalRows = workingIndexes.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));
    if (currentPage > totalPages) {
        currentPage = totalPages;
    }

    const start = (currentPage - 1) * pageSize;
    const end = Math.min(start + pageSize, totalRows);
    const visibleIndexes = workingIndexes.slice(start, end);
    const rowsToRender = visibleIndexes.map((idx) => dataset[idx]);

    tableBody.innerHTML = '';
    const fragment = document.createDocumentFragment();
    rowsToRender.forEach((row) => {
        const tr = document.createElement('tr');
        row.forEach((cell) => {
            const td = document.createElement('td');
            td.textContent = cell;
            tr.appendChild(td);
        });
        fragment.appendChild(tr);
    });
    tableBody.appendChild(fragment);

    pageInfo.textContent = `Page ${totalRows ? currentPage : 0} / ${totalPages}`;
    prevPageButton.disabled = currentPage <= 1;
    nextPageButton.disabled = currentPage >= totalPages;
    toggleExports(totalRows > 0);
}

function updateSummary() {
    if (!dataset.length) {
        resultsSummary.hidden = true;
        return;
    }
    const total = dataset.length;
    const filtered = filteredIndexes.length;
    const duration = lastSearchDuration ? ` · ${lastSearchDuration.toFixed(1)} ms` : '';
    resultsSummary.textContent = `${filtered.toLocaleString()} ligne${filtered > 1 ? 's' : ''} affichée${filtered > 1 ? 's' : ''} sur ${total.toLocaleString()}${duration}`;
    resultsSummary.hidden = false;
}

function toggleExports(enabled) {
    copyButton.disabled = !enabled;
    exportCsvButton.disabled = !enabled;
    exportXlsxButton.disabled = !enabled;
}

function compareCells(rowA, rowB, state) {
    const { columnIndex, direction } = state;
    const left = rowA[columnIndex] ?? '';
    const right = rowB[columnIndex] ?? '';
    const comparison = String(left).localeCompare(String(right), undefined, { sensitivity: 'base', numeric: true });
    return direction === 'asc' ? comparison : -comparison;
}

function executeSearch() {
    if (!dataset.length) {
        return;
    }
    const query = searchInput.value.trim();
    const options = {
        caseSensitive: caseSensitiveCheckbox.checked,
        exactMatch: exactMatchCheckbox.checked
    };
    searchButton.disabled = true;
    searchButton.textContent = 'Recherche…';
    worker.postMessage({ type: 'search', query, options });
}

function createZipNavigator(buffer) {
    const view = new DataView(buffer);
    const entries = new Map();
    const endSignature = 0x06054b50;
    let offset = buffer.byteLength - 22;
    while (offset >= 0 && view.getUint32(offset, true) !== endSignature) {
        offset -= 1;
    }
    if (offset < 0) {
        throw new Error('Archive XLSX invalide (répertoire central introuvable).');
    }
    const centralOffset = view.getUint32(offset + 16, true);
    const totalEntries = view.getUint16(offset + 10, true);
    let cursor = centralOffset;
    const centralSignature = 0x02014b50;
    for (let index = 0; index < totalEntries; index += 1) {
        if (view.getUint32(cursor, true) !== centralSignature) {
            throw new Error('Entrée du répertoire central corrompue.');
        }
        const compression = view.getUint16(cursor + 10, true);
        const compressedSize = view.getUint32(cursor + 20, true);
        const nameLength = view.getUint16(cursor + 28, true);
        const extraLength = view.getUint16(cursor + 30, true);
        const commentLength = view.getUint16(cursor + 32, true);
        const localHeaderOffset = view.getUint32(cursor + 42, true);
        const nameBytes = new Uint8Array(buffer, cursor + 46, nameLength);
        const fileName = textDecoder.decode(nameBytes);
        entries.set(fileName, {
            compression,
            compressedSize,
            localHeaderOffset
        });
        cursor += 46 + nameLength + extraLength + commentLength;
    }

    return {
        async getBytes(path) {
            const entry = entries.get(path);
            if (!entry) {
                return null;
            }
            const localSignature = 0x04034b50;
            if (view.getUint32(entry.localHeaderOffset, true) !== localSignature) {
                throw new Error(`Entrée ZIP invalide pour ${path}.`);
            }
            const nameLength = view.getUint16(entry.localHeaderOffset + 26, true);
            const extraLength = view.getUint16(entry.localHeaderOffset + 28, true);
            const dataStart = entry.localHeaderOffset + 30 + nameLength + extraLength;
            const compressed = new Uint8Array(buffer, dataStart, entry.compressedSize);
            if (entry.compression === 0) {
                return new Uint8Array(compressed);
            }
            if (entry.compression === 8) {
                return inflateBytes(compressed);
            }
            throw new Error(`Compression non supportée (${entry.compression}).`);
        },
        async getText(path) {
            const bytes = await this.getBytes(path);
            if (!bytes) {
                return null;
            }
            return textDecoder.decode(bytes);
        }
    };
}

async function inflateBytes(data) {
    if (typeof DecompressionStream === 'undefined') {
        throw new Error('Décompression DEFLATE indisponible sur ce navigateur.');
    }
    const attempt = async (format) => {
        const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream(format));
        const buffer = await new Response(stream).arrayBuffer();
        return new Uint8Array(buffer);
    };
    try {
        return await attempt('deflate-raw');
    } catch (error) {
        return attempt('deflate');
    }
}

function extractFirstSheet(workbookXml) {
    const parser = new DOMParser();
    const document = parser.parseFromString(workbookXml, 'application/xml');
    const sheet = document.querySelector('sheets sheet');
    if (!sheet) {
        throw new Error('Aucune feuille détectée dans le classeur.');
    }
    const relId = sheet.getAttribute('r:id') || sheet.getAttribute('id') || '';
    const name = sheet.getAttribute('name') || 'Feuille1';
    return { relId, name };
}

function resolveSheetPath(relId, relationshipsXml) {
    if (!relationshipsXml || !relId) {
        return null;
    }
    const parser = new DOMParser();
    const document = parser.parseFromString(relationshipsXml, 'application/xml');
    const relationships = document.getElementsByTagName('Relationship');
    for (let i = 0; i < relationships.length; i += 1) {
        const rel = relationships[i];
        if (rel.getAttribute('Id') === relId) {
            const target = rel.getAttribute('Target') || '';
            return normalizePath('xl/', target);
        }
    }
    return null;
}

function normalizePath(baseDir, target) {
    if (!target) {
        return '';
    }
    if (target.startsWith('/')) {
        return target.replace(/^\/+/, '');
    }
    const baseParts = baseDir.split('/').filter(Boolean);
    const segments = target.split('/');
    for (const segment of segments) {
        if (!segment || segment === '.') {
            continue;
        }
        if (segment === '..') {
            baseParts.pop();
        } else {
            baseParts.push(segment);
        }
    }
    return baseParts.join('/');
}

function parseSharedStrings(xml) {
    if (!xml) {
        return [];
    }
    const parser = new DOMParser();
    const document = parser.parseFromString(xml, 'application/xml');
    const items = document.getElementsByTagName('si');
    const values = [];
    for (let i = 0; i < items.length; i += 1) {
        const item = items[i];
        const texts = item.getElementsByTagName('t');
        let composed = '';
        for (let j = 0; j < texts.length; j += 1) {
            composed += texts[j].textContent || '';
        }
        values.push(sanitizeValue(composed));
    }
    return values;
}

async function extractSheetRows(sheetXml, sharedStrings) {
    const parser = new DOMParser();
    const document = parser.parseFromString(sheetXml, 'application/xml');
    const rowsNodes = Array.from(document.getElementsByTagName('row'));
    if (!rowsNodes.length) {
        return { headers: [], rows: [] };
    }
    const headers = [];
    const rows = [];
    const dataRowTotal = Math.max(rowsNodes.length - 1, 0);
    for (let index = 0; index < rowsNodes.length; index += 1) {
        const values = parseRow(rowsNodes[index], sharedStrings);
        if (index === 0) {
            const headerCount = values.length;
            for (let col = 0; col < headerCount; col += 1) {
                const candidate = sanitizeValue(values[col]);
                headers[col] = candidate || `Colonne ${col + 1}`;
            }
            if (!headers.length && headerCount) {
                for (let col = 0; col < headerCount; col += 1) {
                    headers[col] = `Colonne ${col + 1}`;
                }
            }
            continue;
        }
        if (values.length > headers.length) {
            for (let col = headers.length; col < values.length; col += 1) {
                headers[col] = `Colonne ${col + 1}`;
            }
        }
        if (!headers.length) {
            continue;
        }
        const normalized = new Array(headers.length).fill('');
        for (let col = 0; col < values.length; col += 1) {
            if (values[col] !== undefined) {
                normalized[col] = sanitizeValue(values[col]);
            }
        }
        rows.push(normalized);
        if (dataRowTotal && (rows.length % 250 === 0 || rows.length === dataRowTotal)) {
            const percent = Math.min(100, (rows.length / dataRowTotal) * 100);
            updateProgress(percent, `Lecture XLSX (${rows.length.toLocaleString()} lignes)`);
            await delayFrame();
        }
    }
    updateProgress(100, `Lecture XLSX (${rows.length.toLocaleString()} lignes)`);
    return { headers, rows };
}

function parseRow(rowNode, sharedStrings) {
    const cells = Array.from(rowNode.getElementsByTagName('c'));
    if (!cells.length) {
        return [];
    }
    const values = [];
    let lastIndex = -1;
    cells.forEach((cell, position) => {
        const reference = cell.getAttribute('r');
        let columnIndex;
        if (reference) {
            columnIndex = columnRefToIndex(reference);
        } else {
            columnIndex = lastIndex + 1;
        }
        lastIndex = columnIndex;
        values[columnIndex] = extractCellValue(cell, sharedStrings);
    });
    for (let i = 0; i < values.length; i += 1) {
        if (values[i] === undefined) {
            values[i] = '';
        }
    }
    return values;
}

function extractCellValue(cell, sharedStrings) {
    const type = cell.getAttribute('t');
    if (type === 'inlineStr') {
        const texts = cell.getElementsByTagName('t');
        let inline = '';
        for (let i = 0; i < texts.length; i += 1) {
            inline += texts[i].textContent || '';
        }
        return inline;
    }
    const valueNode = cell.getElementsByTagName('v')[0];
    const rawValue = valueNode ? valueNode.textContent || '' : '';
    if (type === 's') {
        const index = Number(rawValue);
        return Number.isFinite(index) ? (sharedStrings[index] ?? '') : '';
    }
    if (type === 'b') {
        return rawValue === '1' ? 'TRUE' : 'FALSE';
    }
    return rawValue;
}

function columnRefToIndex(reference) {
    const match = reference.match(/[A-Za-z]+/);
    if (!match) {
        return 0;
    }
    const letters = match[0].toUpperCase();
    let index = 0;
    for (let i = 0; i < letters.length; i += 1) {
        index = index * 26 + (letters.charCodeAt(i) - 64);
    }
    return index - 1;
}

function columnIndexToName(index) {
    let dividend = index + 1;
    let columnName = '';
    while (dividend > 0) {
        const modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(65 + modulo) + columnName;
        dividend = Math.floor((dividend - modulo) / 26);
    }
    return columnName;
}

function sanitizeSheetName(name) {
    const cleaned = (name || '').replace(/[\\/*?:\[\]]/g, '').trim();
    if (!cleaned) {
        return 'Feuille1';
    }
    return cleaned.length > 31 ? cleaned.slice(0, 31) : cleaned;
}

function escapeXml(value) {
    return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;')
        .replace(/\r?\n/g, '&#10;');
}

async function buildXlsxBlob(headers, rows, sheetName) {
    if (!headers.length) {
        throw new Error('Aucune donnée à exporter.');
    }
    const safeName = sanitizeSheetName(sheetName);
    const worksheetXml = buildWorksheetXml(headers, rows);
    const workbookXml = `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${escapeXml(safeName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`;
    const workbookRels = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';
    const relationshipsRoot = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>';
    const now = new Date().toISOString();
    const coreXml = `<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(safeName)}</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
    const appXml = '<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Explorateur CSV/XLSX</Application></Properties>';
    const contentTypes = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
    const entries = [
        { path: '[Content_Types].xml', data: textEncoder.encode(contentTypes) },
        { path: '_rels/.rels', data: textEncoder.encode(relationshipsRoot) },
        { path: 'docProps/core.xml', data: textEncoder.encode(coreXml) },
        { path: 'docProps/app.xml', data: textEncoder.encode(appXml) },
        { path: 'xl/workbook.xml', data: textEncoder.encode(workbookXml) },
        { path: 'xl/_rels/workbook.xml.rels', data: textEncoder.encode(workbookRels) },
        { path: 'xl/worksheets/sheet1.xml', data: textEncoder.encode(worksheetXml) }
    ];
    return buildZip(entries);
}

function buildWorksheetXml(headers, rows) {
    const columnCount = headers.length;
    const lines = ['<?xml version="1.0" encoding="UTF-8"?>', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">', '  <sheetData>'];
    const composeRow = (cells, rowNumber) => {
        const parts = [];
        for (let col = 0; col < columnCount; col += 1) {
            const value = col < cells.length ? String(cells[col] ?? '') : '';
            const preserve = /^\s|\s$/.test(value) ? ' xml:space="preserve"' : '';
            const cellRef = `${columnIndexToName(col)}${rowNumber}`;
            parts.push(`<c r="${cellRef}" t="inlineStr"><is><t${preserve}>${escapeXml(value)}</t></is></c>`);
        }
        return `    <row r="${rowNumber}">${parts.join('')}</row>`;
    };
    lines.push(composeRow(headers, 1));
    rows.forEach((row, index) => {
        const normalized = row.slice(0, columnCount);
        while (normalized.length < columnCount) {
            normalized.push('');
        }
        lines.push(composeRow(normalized, index + 2));
    });
    lines.push('  </sheetData>', '</worksheet>');
    return lines.join('\n');
}

function buildZip(entries) {
    const fileParts = [];
    const centralParts = [];
    let offset = 0;
    entries.forEach((entry) => {
        const nameBytes = textEncoder.encode(entry.path);
        const dataBytes = entry.data instanceof Uint8Array ? entry.data : new Uint8Array(entry.data);
        const size = dataBytes.length;
        const crc = computeCrc32(dataBytes);
        const localHeader = new Uint8Array(30 + nameBytes.length);
        const localView = new DataView(localHeader.buffer);
        localView.setUint32(0, 0x04034b50, true);
        localView.setUint16(4, 20, true);
        localView.setUint16(6, 0, true);
        localView.setUint16(8, 0, true);
        localView.setUint16(10, 0, true);
        localView.setUint16(12, 0, true);
        localView.setUint32(14, crc, true);
        localView.setUint32(18, size, true);
        localView.setUint32(22, size, true);
        localView.setUint16(26, nameBytes.length, true);
        localView.setUint16(28, 0, true);
        localHeader.set(nameBytes, 30);
        fileParts.push(localHeader, dataBytes);
        const centralHeader = new Uint8Array(46 + nameBytes.length);
        const centralView = new DataView(centralHeader.buffer);
        centralView.setUint32(0, 0x02014b50, true);
        centralView.setUint16(4, 20, true);
        centralView.setUint16(6, 20, true);
        centralView.setUint16(8, 0, true);
        centralView.setUint16(10, 0, true);
        centralView.setUint16(12, 0, true);
        centralView.setUint16(14, 0, true);
        centralView.setUint32(16, crc, true);
        centralView.setUint32(20, size, true);
        centralView.setUint32(24, size, true);
        centralView.setUint16(28, nameBytes.length, true);
        centralView.setUint16(30, 0, true);
        centralView.setUint16(32, 0, true);
        centralView.setUint16(34, 0, true);
        centralView.setUint16(36, 0, true);
        centralView.setUint32(38, 0, true);
        centralView.setUint32(42, offset, true);
        centralHeader.set(nameBytes, 46);
        centralParts.push(centralHeader);
        offset += localHeader.length + dataBytes.length;
    });
    const centralOffset = offset;
    centralParts.forEach((part) => {
        offset += part.length;
    });
    const centralSize = offset - centralOffset;
    const eocd = new Uint8Array(22);
    const view = new DataView(eocd.buffer);
    view.setUint32(0, 0x06054b50, true);
    view.setUint16(4, 0, true);
    view.setUint16(6, 0, true);
    view.setUint16(8, entries.length, true);
    view.setUint16(10, entries.length, true);
    view.setUint32(12, centralSize, true);
    view.setUint32(16, centralOffset, true);
    view.setUint16(20, 0, true);
    return new Blob([...fileParts, ...centralParts, eocd], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

const CRC32_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
        let value = i;
        for (let j = 0; j < 8; j += 1) {
            value = (value & 1) ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
        }
        table[i] = value >>> 0;
    }
    return table;
})();

function computeCrc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) {
        const index = (crc ^ bytes[i]) & 0xff;
        crc = (CRC32_TABLE[index] ^ (crc >>> 8)) >>> 0;
    }
    return (crc ^ 0xffffffff) >>> 0;
}

function delayFrame() {
    return new Promise((resolve) => setTimeout(resolve));
}
