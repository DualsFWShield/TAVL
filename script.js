/**
 * TAVL Survey Tool - Main Script
 * Handles Excel parsing, UI rendering, interaction logic, and local persistence.
 */

document.addEventListener('DOMContentLoaded', () => {

    /* ==========================================================================
       CONFIGURATION
       Modify this section to adapt to changes in the Excel file structure.
       ========================================================================== */
    const CONFIG = {
        // Excel Matrix Structure (Row Indices are 1-based)
        EXCEL: {
            ROW_CATEGORY: 3,    // Main Headers (e.g. "Sécurité")
            ROW_QUESTION: 4,    // Questions (e.g. "Extincteur ?")
            ROW_TYPE: 5,        // Data Types (e.g. "v/f", "nombre")
            ROW_DATA_START: 6,  // Where actual data begins
            DEFAULT_AUDITOIRE_COL: 3 // Fallback column index if not found by name
        },

        // Storage Keys for IndexedDB
        DB: {
            NAME: 'TavlDB',
            STORE_FILE: 'file',
            STORE_EDITS: 'edits'
        },

        // Keywords used to identify special columns or behaviors (Case Insensitive)
        KEYWORDS: {
            IDENTITY_COL: ['auditoires'], // To find the sidebar name column
            READ_ONLY: [
                'bâtiment', 'batiment',
                'auditoires', // Strict Check logic is applied in code
                'capacité annoncée',
                'gradin' // Part of GMF logic
            ],
            OPTIONAL_BADGE_TEXT: 'Facultatif'
        },

        // Magic Fill Logic Configuration
        AUTO_FILL: {
            CAPACITY_SOURCE: 'capacité annoncée',   // Source column
            CAPACITY_TARGET: ['capacité réelle', 'réellement fonctionnelles'], // Target columns
            DATE_TARGET: 'date de passage',
            NEGATIVE_KEYWORDS: ['humidit', 'infiltration'], // defaulting to "Non" instead of "Oui"
            DEFAULT_DATE_FORMAT: 'dd/mm/yyyy'
        },

        // Supported Data Types in Excel Row 5
        TYPES: {
            TRUE_FALSE: 'v/f',
            YES_NO: 'o/n',
            DATE: 'date',
            NUMBER: 'nombre',
            GMF: 'gmf', // Gradin/Mobile/Fixe
            TEXT: 'text'
        }
    };

    /* ==========================================================================
       DOM ELEMENTS
       ========================================================================== */
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const uploadSection = document.getElementById('upload-section');
    const splitView = document.getElementById('split-view');
    const auditoireList = document.getElementById('auditoire-list');
    const formContainer = document.getElementById('form-container');
    const currentAuditoireTitle = document.getElementById('current-auditoire-title');
    const exportXlsxBtn = document.getElementById('export-xlsx-btn');
    const clearBtn = document.getElementById('clear-btn');
    const nextFieldBtn = document.getElementById('next-field-btn');
    const unlockBtn = document.getElementById('unlock-btn');
    const fillDefaultsBtn = document.getElementById('fill-defaults-btn');

    /* ==========================================================================
       STATE & VARIABLES
       ========================================================================== */
    let currentWorkbook = null;
    let mainWorksheet = null;
    let schema = []; // Array of { colIndex, category, question, type }
    let dataRows = []; // Array of { rowIndex, auditoireName, ... }
    let currentRowIndex = null;
    let db = null;
    let lastFocusedInput = null; // Track focus for smart navigation

    /* ==========================================================================
       INDEXED DB PREISTENCE
       ========================================================================== */
    const initDB = () => {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(CONFIG.DB.NAME, 2);
            request.onupgradeneeded = (e) => {
                db = e.target.result;
                if (!db.objectStoreNames.contains(CONFIG.DB.STORE_FILE)) db.createObjectStore(CONFIG.DB.STORE_FILE);
                if (!db.objectStoreNames.contains(CONFIG.DB.STORE_EDITS)) db.createObjectStore(CONFIG.DB.STORE_EDITS, { keyPath: 'id' });
            };
            request.onsuccess = (e) => {
                db = e.target.result;
                resolve(db);
                checkSavedSession();
            };
            request.onerror = (e) => reject(e);
        });
    };

    async function saveFileToDB(fileBuffer, fileName) {
        if (!db) return;
        const tx = db.transaction([CONFIG.DB.STORE_FILE], 'readwrite');
        tx.objectStore(CONFIG.DB.STORE_FILE).put({ buffer: fileBuffer, name: fileName }, 'current');
    }

    async function saveEditToDB(row, col, val) {
        if (!db) return;
        const tx = db.transaction([CONFIG.DB.STORE_EDITS], 'readwrite');
        const id = `${row}-${col}`;
        tx.objectStore(CONFIG.DB.STORE_EDITS).put({ id, row, col, val });
    }

    async function clearDB() {
        if (!db) return;
        const tx = db.transaction([CONFIG.DB.STORE_FILE, CONFIG.DB.STORE_EDITS], 'readwrite');
        tx.objectStore(CONFIG.DB.STORE_FILE).clear();
        tx.objectStore(CONFIG.DB.STORE_EDITS).clear();
    }

    async function checkSavedSession() {
        if (!db) return;
        const tx = db.transaction([CONFIG.DB.STORE_FILE], 'readonly');
        const req = tx.objectStore(CONFIG.DB.STORE_FILE).get('current');

        req.onsuccess = async () => {
            if (req.result && req.result.buffer) {
                if (confirm('Une session précédente (' + req.result.name + ') a été trouvée. Voulez-vous la restaurer ?')) {
                    const blob = new Blob([req.result.buffer]);
                    const file = new File([blob], req.result.name);
                    await handleFile(file, false); // Do not overwrite DB yet

                    // Restore Edits
                    const tx2 = db.transaction([CONFIG.DB.STORE_EDITS], 'readonly');
                    const cursorReq = tx2.objectStore(CONFIG.DB.STORE_EDITS).openCursor();
                    cursorReq.onsuccess = (e) => {
                        const cursor = e.target.result;
                        if (cursor) {
                            const { row, col, val } = cursor.value;
                            const r = mainWorksheet.getRow(row);
                            const c = r.getCell(col);
                            c.value = val;
                            cursor.continue();
                        } else {
                            console.log('Session restored');
                        }
                    };
                } else {
                    clearDB();
                }
            }
        };
    }

    /* ==========================================================================
       FILE HANDLING & PARSING
       ========================================================================== */

    // Helper to safely get cell value
    const getVal = (row, colIndex) => {
        const cell = row.getCell(colIndex);
        if (!cell || cell.value === null) return '';
        if (typeof cell.value === 'object') {
            if (cell.value.richText) return cell.value.richText.map(t => t.text).join('');
            if (cell.value.text) return cell.value.text;
            if (cell.value.result !== undefined) return cell.value.result.toString();
        }
        return cell.value.toString();
    };

    async function handleFile(file, isNewUpload = false) {
        try {
            const arrayBuffer = await file.arrayBuffer();

            if (isNewUpload) {
                await clearDB();
                saveFileToDB(arrayBuffer, file.name);
            }

            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            currentWorkbook = workbook;
            mainWorksheet = workbook.worksheets[0];

            parseMatrixStructure(mainWorksheet);
            renderSidebar();

            uploadSection.classList.add('hidden');
            splitView.classList.remove('hidden');

        } catch (error) {
            console.error('Error parsing file:', error);
            alert('Erreur: Structure du fichier non reconnue ou fichier invalide.');
        }
    }

    /**
     * Parses the Excel Header Rows to Determine Schema
     */
    function parseMatrixStructure(sheet) {
        schema = [];
        dataRows = [];

        let auditoireColIndex = CONFIG.EXCEL.DEFAULT_AUDITOIRE_COL;
        let lastCategory = '';

        const row3 = sheet.getRow(CONFIG.EXCEL.ROW_CATEGORY);
        const row4 = sheet.getRow(CONFIG.EXCEL.ROW_QUESTION);
        const row5 = sheet.getRow(CONFIG.EXCEL.ROW_TYPE);

        const colCount = sheet.columnCount;

        for (let c = 1; c <= colCount; c++) {
            let cat = getVal(row3, c);
            if (cat) lastCategory = cat;

            // Detect Identity Column (Name of Auditorium)
            if (cat && CONFIG.KEYWORDS.IDENTITY_COL.some(k => cat.toLowerCase().includes(k))) {
                auditoireColIndex = c;
            }

            let question = getVal(row4, c);
            let typeRaw = getVal(row5, c);
            let type = typeRaw.toLowerCase().trim();

            // Normalization of Type Strings
            if (type.includes('..') || type.includes('date')) type = CONFIG.TYPES.DATE;

            // Specific Header Overrides (GMF for Gradins)
            if ((cat && cat.includes('Gradin')) || (question && question.includes('Gradin'))) {
                type = CONFIG.TYPES.GMF;
            }

            if (question || type) {
                schema.push({
                    colIndex: c,
                    category: lastCategory,
                    question: question,
                    type: type || CONFIG.TYPES.TEXT
                });
            }
        }

        // Scan Data Rows for Auditoriums
        sheet.eachRow((row, rowIndex) => {
            if (rowIndex < CONFIG.EXCEL.ROW_DATA_START) return;

            const name = getVal(row, auditoireColIndex);
            if (name) {
                dataRows.push({ rowIndex: rowIndex, name: name });
            }
        });
    }

    /* ==========================================================================
       UI RENDERING
       ========================================================================== */

    function renderSidebar() {
        auditoireList.innerHTML = '';
        dataRows.forEach((item, index) => {
            const li = document.createElement('li');
            li.className = 'sidebar-item';
            li.textContent = item.name;
            li.onclick = () => selectAuditoire(index);
            auditoireList.appendChild(li);
        });
    }

    function selectAuditoire(index) {
        document.querySelectorAll('.sidebar-item').forEach(el => el.classList.remove('active'));
        auditoireList.children[index].classList.add('active');

        const item = dataRows[index];
        currentRowIndex = item.rowIndex;
        currentAuditoireTitle.textContent = item.name;

        renderForm(item.rowIndex);
    }

    function renderForm(rowIndex) {
        formContainer.innerHTML = '';
        const row = mainWorksheet.getRow(rowIndex);

        // Group fields by Category
        const byCategory = {};
        schema.forEach(field => {
            if (!byCategory[field.category]) byCategory[field.category] = [];
            byCategory[field.category].push(field);
        });

        // Iterate Categories
        for (const [category, fields] of Object.entries(byCategory)) {
            const catDiv = document.createElement('div');
            catDiv.className = 'form-category';
            const h3 = document.createElement('h3');
            h3.textContent = category || 'Général';
            catDiv.appendChild(h3);

            let hasVisibleFields = false;

            fields.forEach(field => {
                if (!field.question && !field.type) return;
                hasVisibleFields = true;

                const cell = row.getCell(field.colIndex);
                const val = getVal(row, field.colIndex);

                // Check for Pattern Fill (Optional Detection)
                const isFacultatif = (c) => {
                    if (!c || !c.style || !c.style.fill) return false;
                    const f = c.style.fill;
                    return (f.type === 'pattern' && f.pattern && f.pattern !== 'none' && f.pattern !== 'solid');
                };

                const isOptional = isFacultatif(cell);

                const group = document.createElement('div');
                group.className = 'field-group';

                const label = document.createElement('label');
                label.className = 'field-question';
                label.textContent = field.question || field.category;

                // Badges
                if (isOptional) {
                    const badge = document.createElement('span');
                    badge.className = 'badge-optional';
                    badge.textContent = CONFIG.KEYWORDS.OPTIONAL_BADGE_TEXT;
                    label.appendChild(badge);
                }

                // Replace "testé" text with Badge
                const labelTextLower = label.textContent.toLowerCase();
                if (labelTextLower.includes('(testé)') || labelTextLower.includes('(testée)')) {
                    label.childNodes[0].textContent = label.childNodes[0].textContent.replace(/\(testé\)/i, '').replace(/\(testée\)/i, '');
                    const badgeTest = document.createElement('span');
                    badgeTest.className = 'badge-test-required';
                    badgeTest.textContent = 'Test manuel requis';
                    label.appendChild(badgeTest);
                }

                group.appendChild(label);

                // --- Read-Only Logic ---
                const catNorm = (field.category || '').toLowerCase().trim();
                const questNorm = (field.question || '').toLowerCase().trim();
                let isReadOnly = false;

                const forceEditMode = document.body.classList.contains('force-edit-mode');

                if (!forceEditMode) {
                    // Strict checks for structural fields
                    const isBatiment = catNorm === 'bâtiments' || catNorm === 'batiments' || questNorm === 'bâtiments' || questNorm === 'batiments';
                    const isAuditoire = catNorm === 'auditoires' || questNorm === 'auditoires';
                    const isCapacite = catNorm.includes(CONFIG.AUTO_FILL.CAPACITY_SOURCE) || questNorm.includes(CONFIG.AUTO_FILL.CAPACITY_SOURCE);
                    const isGradin = catNorm.includes('gradin') && catNorm.includes('mobile'); // Complex header usually for GMF

                    if (isBatiment || isAuditoire || isCapacite || isGradin ||
                        (field.type === CONFIG.TYPES.GMF && (catNorm.includes('gradin') || questNorm.includes('gradin')))) {
                        isReadOnly = true;
                    }
                }

                // --- Render Input ---
                const type = field.type;

                if (type === CONFIG.TYPES.TRUE_FALSE || type === CONFIG.TYPES.YES_NO || type === CONFIG.TYPES.GMF) {
                    // Radio Groups
                    let options = [];
                    let labels = {};

                    if (type === CONFIG.TYPES.TRUE_FALSE) { options = ['v', 'f']; labels = { v: 'Vrai', f: 'Faux' }; }
                    else if (type === CONFIG.TYPES.YES_NO) { options = ['o', 'n']; labels = { o: 'Oui', n: 'Non' }; }
                    else if (type === CONFIG.TYPES.GMF) { options = ['G', 'M', 'F']; labels = { G: 'Gradin', M: 'Mobile', F: 'Fixe' }; }

                    const radioContainer = document.createElement('div');
                    radioContainer.className = 'radio-group';
                    if (isReadOnly) radioContainer.classList.add('disabled-group');

                    options.forEach(opt => {
                        const wrapper = document.createElement('label');
                        wrapper.className = 'radio-option';

                        const input = document.createElement('input');
                        input.type = 'radio';
                        input.name = `field-${field.colIndex}`;
                        input.value = opt;
                        if (isReadOnly) input.disabled = true;

                        if (val.toString().toLowerCase() === opt.toLowerCase()) input.checked = true;

                        input.addEventListener('change', () => updateCell(field.colIndex, opt));

                        const textSpan = document.createElement('span');
                        textSpan.textContent = labels[opt] || opt;

                        wrapper.appendChild(input);
                        wrapper.appendChild(textSpan);
                        radioContainer.appendChild(wrapper);
                    });
                    group.appendChild(radioContainer);

                } else if (type === CONFIG.TYPES.DATE) {
                    const input = document.createElement('input');
                    input.type = 'date';
                    if (isReadOnly) {
                        input.disabled = true;
                        input.classList.add('input-disabled');
                    }

                    // Parse Date Value
                    if (val && !isNaN(Date.parse(val))) {
                        const d = new Date(val);
                        if (!isNaN(d)) input.value = d.toISOString().split('T')[0];
                        else input.value = val;
                    } else if (typeof val === 'number') {
                        const d = new Date(Math.round((val - 25569) * 86400 * 1000));
                        if (!isNaN(d)) input.value = d.toISOString().split('T')[0];
                    } else {
                        input.value = val;
                    }

                    input.addEventListener('change', (e) => {
                        const d = new Date(e.target.value);
                        updateCell(field.colIndex, !isNaN(d) ? d : e.target.value);
                    });
                    group.appendChild(input);

                } else if (type === CONFIG.TYPES.NUMBER) {
                    const input = document.createElement('input');
                    input.type = 'number';
                    input.value = val;
                    if (isReadOnly) {
                        input.disabled = true;
                        input.classList.add('input-disabled');
                    }
                    input.addEventListener('input', (e) => updateCell(field.colIndex, e.target.value));
                    group.appendChild(input);
                } else {
                    // Default Text Array
                    const input = document.createElement('textarea');
                    input.rows = 2;
                    input.value = val;
                    if (isReadOnly) {
                        input.disabled = true;
                        input.classList.add('input-disabled');
                    }
                    input.addEventListener('input', (e) => updateCell(field.colIndex, e.target.value));
                    group.appendChild(input);
                }

                catDiv.appendChild(group);
            });

            if (hasVisibleFields) formContainer.appendChild(catDiv);
        }
    }

    function updateCell(colIndex, value) {
        if (!currentRowIndex) return;
        const row = mainWorksheet.getRow(currentRowIndex);
        const cell = row.getCell(colIndex);

        cell.value = value;
        if (value instanceof Date) {
            cell.numFmt = CONFIG.AUTO_FILL.DEFAULT_DATE_FORMAT;
        }

        saveEditToDB(currentRowIndex, colIndex, value);
    }

    /* ==========================================================================
       INTERACTIVE FEATURES
       ========================================================================== */

    // Export Logic
    exportXlsxBtn.addEventListener('click', async () => {
        if (!currentWorkbook) return;
        try {
            const buffer = await currentWorkbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Releve_TAVL_Completed.xlsx';
            a.click();
            window.URL.revokeObjectURL(url);
        } catch (error) {
            console.error(error);
            alert('Export failed');
        }
    });

    clearBtn.addEventListener('click', () => {
        splitView.classList.add('hidden');
        uploadSection.classList.remove('hidden');
        currentWorkbook = null;
        fileInput.value = '';
    });

    // 1. Force Edit Mode
    unlockBtn.addEventListener('click', () => {
        if (confirm("Voulez-vous activer le mode 'Édition Forcée' ? Cela déverrouillera tous les champs structurels (Bâtiments, Capacité, etc.).")) {
            document.body.classList.toggle('force-edit-mode');
            if (currentRowIndex) renderForm(currentRowIndex);
        }
    });

    // 2. Fill Defaults
    fillDefaultsBtn.addEventListener('click', () => {
        if (!currentRowIndex || !mainWorksheet) return;

        const row = mainWorksheet.getRow(currentRowIndex);
        let editsMade = 0;

        // Find Value of "Announced Capacity"
        let capaAnnonceVal = '';
        const capaField = schema.find(f => {
            const c = (f.category || '').toLowerCase();
            const q = (f.question || '').toLowerCase();
            return c === CONFIG.AUTO_FILL.CAPACITY_SOURCE || q === CONFIG.AUTO_FILL.CAPACITY_SOURCE;
        });
        if (capaField) {
            capaAnnonceVal = getVal(row, capaField.colIndex);
        }

        const today = new Date().toISOString().split('T')[0];

        schema.forEach(field => {
            // Optional Check
            const cell = row.getCell(field.colIndex);
            let isOptional = false;
            if (cell && cell.style && cell.style.fill) {
                const f = cell.style.fill;
                if (f.type === 'pattern' && f.pattern && f.pattern !== 'none' && f.pattern !== 'solid') isOptional = true;
            }
            if (isOptional) return;

            // Skip if value exists
            const currentVal = getVal(row, field.colIndex);
            if (currentVal && currentVal.toString().trim() !== '') return;

            const qNorm = (field.question || '').toLowerCase();
            const catNorm = (field.category || '').toLowerCase();

            // Logic: Real Capacity
            if ((catNorm.includes('capacité réelle') || CONFIG.AUTO_FILL.CAPACITY_TARGET.some(t => qNorm.includes(t)))) {
                if (capaAnnonceVal) {
                    updateCell(field.colIndex, capaAnnonceVal);
                    editsMade++;
                }
            }
            // Logic: Date
            else if (field.type === CONFIG.TYPES.DATE || qNorm.includes(CONFIG.AUTO_FILL.DATE_TARGET)) {
                const today = new Date();
                today.setHours(0, 0, 0, 0);
                updateCell(field.colIndex, today);
                editsMade++;
            }
            // Logic: O/N (Default Oui, except negatives)
            else if (field.type === CONFIG.TYPES.YES_NO) {
                if (CONFIG.AUTO_FILL.NEGATIVE_KEYWORDS.some(k => qNorm.includes(k))) {
                    updateCell(field.colIndex, 'n');
                } else {
                    updateCell(field.colIndex, 'o');
                }
                editsMade++;
            }
            // Logic: V/F (Default Vrai)
            else if (field.type === CONFIG.TYPES.TRUE_FALSE) {
                updateCell(field.colIndex, 'v');
                editsMade++;
            }
        });

        if (editsMade > 0) {
            renderForm(currentRowIndex);
        }
    });

    // 3. Smart Navigation

    // Drag/Drop Listeners
    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0], true);
    });
    fileInput.addEventListener('change', (e) => { if (e.target.files.length > 0) handleFile(e.target.files[0], true); });


    // Show/Hide FAB
    const observer = new MutationObserver(() => {
        if (!splitView.classList.contains('hidden')) {
            nextFieldBtn.classList.remove('hidden');
        } else {
            nextFieldBtn.classList.add('hidden');
        }
    });
    observer.observe(splitView, { attributes: true });

    formContainer.addEventListener('focusin', (e) => {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
            lastFocusedInput = e.target;
        }
    });

    nextFieldBtn.addEventListener('click', () => {
        const inputs = Array.from(formContainer.querySelectorAll('input:not(:disabled), textarea:not(:disabled)'));

        let startIdx = 0;
        const current = lastFocusedInput;
        const currentIdx = inputs.indexOf(current);

        if (currentIdx !== -1) {
            startIdx = currentIdx + 1;
            // Skip siblings if starting from radio
            if (current.type === 'radio') {
                const name = current.name;
                while (startIdx < inputs.length && inputs[startIdx].name === name) {
                    startIdx++;
                }
            }
        }

        let found = null;

        const isTarget = (input) => {
            const group = input.closest('.field-group');
            if (group) {
                const badge = group.querySelector('.badge-optional');
                if (badge) return false;
            }

            if (input.type === 'radio') {
                const name = input.name;
                const groupRadios = formContainer.querySelectorAll(`input[name="${name}"]`);
                let isChecked = false;
                groupRadios.forEach(r => { if (r.checked) isChecked = true; });
                return !isChecked;
            } else {
                return !input.value;
            }
        };

        // Forward Pass
        for (let i = startIdx; i < inputs.length; i++) {
            if (isTarget(inputs[i])) {
                found = inputs[i];
                break;
            }
            // Skip Siblings Logic
            if (inputs[i].type === 'radio') {
                const name = inputs[i].name;
                while (i + 1 < inputs.length && inputs[i + 1].name === name) {
                    i++;
                }
            }
        }

        // Wrap Around Pass
        if (!found && startIdx > 0) {
            for (let i = 0; i < startIdx; i++) {
                if (isTarget(inputs[i])) {
                    found = inputs[i];
                    break;
                }
                if (inputs[i].type === 'radio') {
                    const name = inputs[i].name;
                    while (i + 1 < startIdx && inputs[i + 1].name === name) {
                        i++;
                    }
                }
            }
        }

        if (found) {
            found.scrollIntoView({ behavior: 'smooth', block: 'center' });
            found.focus();
        } else {
            alert("Tous les champs obligatoires semblent remplis !");
        }
    });

    initDB();
});
