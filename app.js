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

exportXlsxButton.addEventListener('click', () => {
    const exportRows = getFilteredRows();
    if (!exportRows.length) {
        return;
    }
    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...exportRows]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Résultats');
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    triggerDownload(url, deriveFileName('xlsx'));
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
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
        throw new Error('Le classeur ne contient aucune feuille lisible.');
    }
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const headerRowIndex = range.s.r;
    const headers = [];
    for (let C = range.s.c; C <= range.e.c; C += 1) {
        const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: C });
        const cell = worksheet[cellAddress];
        const headerLabel = cell ? sanitizeValue(cell.v) : `Colonne ${C - range.s.c + 1}`;
        headers.push(headerLabel);
    }

    const rows = [];
    const totalRows = range.e.r - headerRowIndex;
    const chunkSize = 400;

    for (let startRow = headerRowIndex + 1; startRow <= range.e.r; startRow += chunkSize) {
        const endRow = Math.min(startRow + chunkSize - 1, range.e.r);
        const subsetRange = {
            s: { r: startRow, c: range.s.c },
            e: { r: endRow, c: range.e.c }
        };
        const slice = XLSX.utils.sheet_to_json(worksheet, {
            header: headers,
            range: subsetRange,
            defval: '',
            blankrows: false
        });
        slice.forEach((record) => {
            const row = headers.map((header) => sanitizeValue(record[header]));
            rows.push(row);
        });
        const processedRows = rows.length;
        const percent = totalRows ? (processedRows / totalRows) * 100 : 100;
        updateProgress(percent, `Lecture XLSX (${rows.length.toLocaleString()} lignes)`);
        await new Promise((resolve) => setTimeout(resolve));
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
