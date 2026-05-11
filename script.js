/**
 * TAVL Survey Tool - Main Script
 * Handles Excel parsing, UI rendering, interaction logic, and local persistence.
 */

// Register Service Worker for PWA
if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
        navigator.serviceWorker.register('./sw.js')
            .then(registration => {
                console.log('ServiceWorker registration successful with scope: ', registration.scope);
            })
            .catch(err => {
                console.log('ServiceWorker registration failed: ', err);
            });
    });
}

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
            OPTIONAL_BADGE_TEXT: 'Non-Applicable'
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
        },

        // Excel Styles
        STYLES: {
            GREY_PATTERN: {
                fill: {
                    type: 'pattern',
                    pattern: 'darkGray',
                    fgColor: { argb: 'FF000000' },
                    bgColor: { argb: 'FFFFFFFF' }
                }
            }
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
    const searchInput = document.getElementById('search-input');
    const searchContainer = document.getElementById('search-container');
    const themeToggleBtn = document.getElementById('theme-toggle');
    const progressBarContainer = document.getElementById('progress-bar-container');
    const progressFill = document.getElementById('progress-fill');
    const progressCount = document.getElementById('progress-count');
    const scrollTopBtn = document.getElementById('scroll-top-btn');

    // Category icon mapping
    const CATEGORY_ICONS = {
        'bâtiment': '🏢', 'batiment': '🏢', 'auditoire': '🎓',
        'capacité annoncée': '👥', 'capacite annoncee': '👥',
        'capacité réelle': '👥', 'capacite reelle': '👥',
        'gradin': '🏟️', 'mobilier': '🪑',
        'objet': '📦', 'perdu': '📦',
        'horloge': '🕰️', 'téléphone': '📞', 'telephone': '📞',
        'audiovisuel': '🎥', 'lavabo': '🚰', 'radiateur': '🌡️',
        'occultation': '🪟', 'store': '🪟', 'lamelle': '🪟',
        'poignée': '🚪', 'poignee': '🚪', 'fenêtre': '🪟', 'fenetre': '🪟', 'porte-manteau': '🧥',
        'revêtement': '🏗️', 'revetement': '🏗️', 'sol': '🏗️',
        'prise': '🔌', 'courant': '🔌', '220': '🔌',
        'éclairage': '💡', 'eclairage': '💡', 'lumineux': '💡',
        'porte': '🚪', 'accès': '🚪', 'acces': '🚪',
        'sticker': '🏷️', 'affichage': '🏷️', 'colle': '🏷️',
        'pictogramme': '🚫', 'fumer': '🚫',
        'humidité': '💧', 'humidite': '💧', 'infiltration': '💧',
        'enlèvement': '🚛', 'enlevement': '🚛', 'gpex': '🚛',
        'constat': '📝', 'remarque': '📝',
        'date': '📅', 'passage': '📅',
        'tavl': '👤', 'relevé': '👤', 'releve': '👤',
        'sécurité': '🔒', 'securite': '🔒',
    };
    function getCategoryIcon(catName) {
        const lower = (catName || '').toLowerCase();
        for (const [key, icon] of Object.entries(CATEGORY_ICONS)) {
            if (lower.includes(key)) return icon;
        }
        return '📋';
    }
    const burgerBtn = document.getElementById('burger-btn');
    const sidebarOverlay = document.getElementById('sidebar-overlay');
    const sidebar = document.querySelector('.sidebar');
    const modalOverlay = document.getElementById('modal-overlay');
    const modalIcon = document.getElementById('modal-icon');
    const modalTitle = document.getElementById('modal-title');
    const modalMessage = document.getElementById('modal-message');
    const modalActions = document.getElementById('modal-actions');

    /* ==========================================================================
       MODAL SYSTEM (replaces browser confirm/alert)
       ========================================================================== */
    function showModal({ icon = 'ℹ️', title = '', message = '', buttons = [] }) {
        return new Promise(resolve => {
            modalIcon.textContent = icon;
            modalTitle.textContent = title;
            modalMessage.textContent = message;
            modalActions.innerHTML = '';

            buttons.forEach(btn => {
                const b = document.createElement('button');
                b.className = 'modal-btn' + (btn.primary ? ' modal-btn-primary' : '');
                b.textContent = btn.label;
                b.addEventListener('click', () => {
                    modalOverlay.classList.add('hidden');
                    resolve(btn.value);
                });
                modalActions.appendChild(b);
            });

            modalOverlay.classList.remove('hidden');
        });
    }

    function showAlert(message, icon = 'ℹ️', title = 'Information') {
        return showModal({
            icon, title, message,
            buttons: [{ label: 'OK', value: true, primary: true }]
        });
    }

    function showConfirm(message, icon = '❓', title = 'Confirmation') {
        return showModal({
            icon, title, message,
            buttons: [
                { label: 'Annuler', value: false },
                { label: 'Confirmer', value: true, primary: true }
            ]
        });
    }

    /* ==========================================================================
       BURGER MENU (Mobile Drawer)
       ========================================================================== */
    function openDrawer() {
        if (!sidebar) return;
        sidebar.classList.add('drawer-open');
        sidebarOverlay.classList.remove('hidden');
        sidebarOverlay.classList.add('visible');
    }

    function closeDrawer() {
        if (!sidebar) return;
        sidebar.classList.remove('drawer-open');
        sidebarOverlay.classList.remove('visible');
        setTimeout(() => sidebarOverlay.classList.add('hidden'), 300);
    }

    burgerBtn.addEventListener('click', () => {
        if (sidebar.classList.contains('drawer-open')) closeDrawer();
        else openDrawer();
    });

    sidebarOverlay.addEventListener('click', closeDrawer);

    /* ==========================================================================
       THEME MANAGEMENT
       ========================================================================== */
    function initTheme() {
        const storedTheme = localStorage.getItem('theme');
        const systemPrefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;

        let theme = 'dark';
        if (storedTheme) {
            theme = storedTheme;
        } else if (!systemPrefersDark) {
            theme = 'light';
        }

        applyTheme(theme);
    }

    function applyTheme(theme) {
        if (theme === 'light') {
            document.documentElement.setAttribute('data-theme', 'light');
            themeToggleBtn.textContent = '🌙'; // Moon to switch back to dark
            themeToggleBtn.title = 'Passer en mode sombre';
        } else {
            document.documentElement.removeAttribute('data-theme');
            themeToggleBtn.textContent = '☀'; // Sun to switch to light
            themeToggleBtn.title = 'Passer en mode clair';
        }
        localStorage.setItem('theme', theme);
    }

    themeToggleBtn.addEventListener('click', () => {
        const currentTheme = document.documentElement.getAttribute('data-theme') === 'light' ? 'light' : 'dark';
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';
        applyTheme(newTheme);
    });

    // Initialize immediately
    initTheme();

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

    async function saveEditToDB(row, col, val, isGreyed = false) {
        if (!db) return;
        const tx = db.transaction([CONFIG.DB.STORE_EDITS], 'readwrite');
        const id = `${row}-${col}`;
        tx.objectStore(CONFIG.DB.STORE_EDITS).put({ id, row, col, val, isGreyed });
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
                const shouldRestore = await showConfirm(
                    'Une session précédente (' + req.result.name + ') a été trouvée. Voulez-vous la restaurer ?',
                    '💾', 'Session sauvegardée'
                );
                if (shouldRestore) {
                    const blob = new Blob([req.result.buffer]);
                    const file = new File([blob], req.result.name);
                    await handleFile(file, false); // Do not overwrite DB yet

                    // Restore Edits
                    const tx2 = db.transaction([CONFIG.DB.STORE_EDITS], 'readonly');
                    const cursorReq = tx2.objectStore(CONFIG.DB.STORE_EDITS).openCursor();
                    cursorReq.onsuccess = (e) => {
                        const cursor = e.target.result;
                        if (cursor) {
                            const { row, col, val, isGreyed } = cursor.value;
                            const r = mainWorksheet.getRow(row);
                            const c = r.getCell(col);
                            c.value = val;

                            // Restore Styling
                            if (isGreyed) {
                                c.style = {
                                    ...c.style,
                                    fill: CONFIG.STYLES.GREY_PATTERN.fill
                                };
                            }

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

            // Reset Form View
            formContainer.innerHTML = '<div class="empty-state">Veuillez sélectionner un auditoire dans la liste à gauche.</div>';
            currentAuditoireTitle.textContent = 'Sélectionnez un auditoire';
            currentRowIndex = null;
            if (searchContainer) searchContainer.style.display = 'none';

            uploadSection.classList.add('hidden');
            splitView.classList.remove('hidden');
            burgerBtn.classList.remove('hidden');

        } catch (error) {
            console.error('Error parsing file:', error);
            showAlert('Structure du fichier non reconnue ou fichier invalide.', '❌', 'Erreur');
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
            li.dataset.rowIndex = item.rowIndex; // Store for easy access

            const spanName = document.createElement('span');
            spanName.textContent = item.name;
            li.appendChild(spanName);

            const check = document.createElement('span');
            check.className = 'status-indicator';
            check.innerHTML = '✔';

            const name = item.name.toLowerCase();
            const threshold = name.includes('hall, sanitaire') ? 40 : 60;

            // Initial Calculation
            const percentage = calculateCompletion(item.rowIndex);
            if (percentage >= threshold) check.classList.add('visible');

            li.appendChild(check);

            li.onclick = () => selectAuditoire(index);
            auditoireList.appendChild(li);
        });
    }

    function calculateCompletion(rowIndex) {
        if (!mainWorksheet) return 0;
        const row = mainWorksheet.getRow(rowIndex);
        let total = 0;
        let filled = 0;

        schema.forEach(field => {
            if (!field.question) return;

            const cell = row.getCell(field.colIndex);

            // Check for Pattern Fill (Optional Detection) OR Manual Grey
            // We reuse the logic: if it has a pattern that is NOT 'none' or 'solid', it is exempted.
            // This covers both 'mediumGray' (original optional) and 'darkGray' (manual toggle)
            let isExempt = false;
            if (cell && cell.style && cell.style.fill) {
                const f = cell.style.fill;
                if (f.type === 'pattern' && f.pattern && f.pattern !== 'none' && f.pattern !== 'solid') {
                    isExempt = true;
                }
            }

            if (isExempt) return; // Don't count this question

            total++;
            const val = getVal(row, field.colIndex);
            if (val && val.toString().trim() !== '') {
                filled++;
            }
        });

        return total === 0 ? 0 : (filled / total) * 100;
    }

    function updateSidebarStatus(rowIndex) {
        const li = Array.from(auditoireList.children).find(el => parseInt(el.dataset.rowIndex) === rowIndex);
        if (li) {
            const rowData = dataRows.find(r => r.rowIndex === rowIndex);
            const name = rowData ? rowData.name.toLowerCase() : '';
            const threshold = name.includes('hall, sanitaire') ? 40 : 60;

            const percentage = calculateCompletion(rowIndex);
            const check = li.querySelector('.status-indicator');
            if (percentage >= threshold) check.classList.add('visible');
            else check.classList.remove('visible');
        }
    }

    function selectAuditoire(index) {
        document.querySelectorAll('.sidebar-item').forEach(el => el.classList.remove('active'));
        auditoireList.children[index].classList.add('active');

        const item = dataRows[index];
        currentRowIndex = item.rowIndex;
        currentAuditoireTitle.textContent = item.name;

        // Fade transition
        formContainer.classList.add('fade-out');
        setTimeout(() => {
            renderForm(item.rowIndex);
            updateProgressBar(item.rowIndex);
            formContainer.classList.remove('fade-out');
            formContainer.classList.add('fade-in');
            setTimeout(() => formContainer.classList.remove('fade-in'), 250);
        }, 150);

        searchInput.value = '';
        if (searchContainer) searchContainer.style.display = 'block';
        if (progressBarContainer) progressBarContainer.style.display = 'block';
        closeDrawer();
    }

    function updateProgressBar(rowIndex) {
        if (!mainWorksheet || !progressFill) return;
        const row = mainWorksheet.getRow(rowIndex);
        let total = 0, filled = 0;
        schema.forEach(field => {
            if (!field.question) return;
            const cell = row.getCell(field.colIndex);
            let isExempt = false;
            if (cell && cell.style && cell.style.fill) {
                const f = cell.style.fill;
                if (f.type === 'pattern' && f.pattern && f.pattern !== 'none' && f.pattern !== 'solid') isExempt = true;
            }
            if (isExempt) return;
            total++;
            const val = getVal(row, field.colIndex);
            if (val && val.toString().trim() !== '') filled++;
        });
        const pct = total === 0 ? 0 : Math.round((filled / total) * 100);
        progressFill.style.width = pct + '%';
        progressCount.textContent = filled + '/' + total;
    }

    function renderForm(rowIndex) {
        formContainer.innerHTML = '';
        const row = mainWorksheet.getRow(rowIndex);

        // Extract structurally read-only fields for ID Card
        const forceEditMode = document.body.classList.contains('force-edit-mode');
        const idFields = [];
        const normalFields = [];

        schema.forEach(field => {
            const catNorm = (field.category || '').toLowerCase().trim();
            const questNorm = (field.question || '').toLowerCase().trim();
            let isStructureReadOnly = false;

            if (!forceEditMode) {
                const isBatiment = catNorm === 'bâtiments' || catNorm === 'batiments' || questNorm === 'bâtiments' || questNorm === 'batiments';
                const isAuditoire = catNorm === 'auditoires' || questNorm === 'auditoires';
                const isCapacite = catNorm.includes(CONFIG.AUTO_FILL.CAPACITY_SOURCE) || questNorm.includes(CONFIG.AUTO_FILL.CAPACITY_SOURCE);
                const isGradin = catNorm.includes('gradin') && catNorm.includes('mobile');
                
                if (isBatiment || isAuditoire || isCapacite || isGradin || (field.type === CONFIG.TYPES.GMF && (catNorm.includes('gradin') || questNorm.includes('gradin')))) {
                    isStructureReadOnly = true;
                }
            }

            if (isStructureReadOnly) {
                idFields.push(field);
            } else {
                normalFields.push(field);
            }
        });

        // Render ID Card if not empty
        if (idFields.length > 0) {
            const idCard = document.createElement('div');
            idCard.className = 'id-card';

            // Find Auditoire Name to use as Title
            const auditoireField = idFields.find(f => {
                const c = (f.category || '').toLowerCase();
                const q = (f.question || '').toLowerCase();
                return c === 'auditoires' || q === 'auditoires';
            });
            const auditoireName = auditoireField ? getVal(row, auditoireField.colIndex) : 'Auditoire';

            const header = document.createElement('div');
            header.className = 'id-card-header';
            header.innerHTML = `<span class="id-icon">🎓</span><h2>${auditoireName}</h2>`;
            idCard.appendChild(header);

            const body = document.createElement('div');
            body.className = 'id-card-body';

            idFields.forEach(field => {
                if (field === auditoireField) return; // skip rendering name again
                const val = getVal(row, field.colIndex);
                if (!val && val !== 0 && val !== '0') return; // skip empty

                const item = document.createElement('div');
                item.className = 'id-card-item';
                
                const label = document.createElement('div');
                label.className = 'id-card-label';
                label.textContent = field.question || field.category;
                
                const value = document.createElement('div');
                value.className = 'id-card-value';
                
                // Format value if needed (GMF translation)
                let displayVal = val;
                if (field.type === CONFIG.TYPES.GMF) {
                    if (val === 'G') displayVal = 'Gradin';
                    else if (val === 'M') displayVal = 'Mobile';
                    else if (val === 'F') displayVal = 'Fixe';
                }
                value.textContent = displayVal;

                item.appendChild(label);
                item.appendChild(value);
                body.appendChild(item);
            });

            idCard.appendChild(body);
            formContainer.appendChild(idCard);
        }

        // Group normal fields by Category
        const byCategory = {};
        normalFields.forEach(field => {
            if (!byCategory[field.category]) byCategory[field.category] = [];
            byCategory[field.category].push(field);
        });

        // Iterate Categories
        for (const [category, fields] of Object.entries(byCategory)) {
            const catDiv = document.createElement('div');
            catDiv.className = 'form-category';
            const h3 = document.createElement('h3');
            const iconSpan = document.createElement('span');
            iconSpan.className = 'category-icon';
            iconSpan.textContent = getCategoryIcon(category);
            h3.appendChild(iconSpan);
            h3.appendChild(document.createTextNode(category || 'Général'));
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

                // Check if currently greyed (darkGray pattern = manually marked N/A)
                const isCurrentlyGrey = (c) => {
                    if (!c || !c.style || !c.style.fill) return false;
                    const f = c.style.fill;
                    return (f.type === 'pattern' && f.pattern === 'darkGray');
                };
                const isGreyed = isCurrentlyGrey(cell);

                // Field is disabled if optional (Non-Applicable) OR manually greyed
                const isFieldDisabled = isOptional || isGreyed;

                const group = document.createElement('div');
                group.className = 'field-group';
                if (isFieldDisabled) group.classList.add('disabled-group');

                const label = document.createElement('label');
                label.className = 'field-question';
                if (isFieldDisabled) label.classList.add('question-optional');
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

                // --- Grey Out Toggle (Manual N/A) ---
                const headerWrapper = document.createElement('div');
                headerWrapper.style.display = 'flex';
                headerWrapper.style.justifyContent = 'space-between';
                headerWrapper.style.alignItems = 'center';

                headerWrapper.appendChild(label);

                // Tooltip wrapper for checkbox
                const toggleWrapper = document.createElement('div');
                toggleWrapper.className = 'grey-toggle-wrapper';

                const greyToggle = document.createElement('input');
                greyToggle.type = 'checkbox';
                greyToggle.title = 'Marquer comme non-accessible (grisé)';
                greyToggle.className = 'grey-toggle-checkbox';

                const tooltipSpan = document.createElement('span');
                tooltipSpan.className = 'tooltip-text';
                tooltipSpan.textContent = 'Non-accessible (griser)';

                // Pre-check if optional or manually greyed
                if (isFieldDisabled) {
                    greyToggle.checked = true;
                }

                greyToggle.onchange = (e) => {
                    toggleCellGrey(field.colIndex, e.target.checked);
                };

                toggleWrapper.appendChild(greyToggle);
                toggleWrapper.appendChild(tooltipSpan);
                headerWrapper.appendChild(toggleWrapper);
                group.appendChild(headerWrapper);

                // --- Render Input ---
                const type = field.type;
                const shouldDisableInput = isFieldDisabled; // Since structural read-only are filtered out

                if (type === CONFIG.TYPES.TRUE_FALSE || type === CONFIG.TYPES.YES_NO || type === CONFIG.TYPES.GMF) {
                    // Toggle Pill Groups
                    let options = [];
                    let labels = {};

                    if (type === CONFIG.TYPES.TRUE_FALSE) { options = ['v', 'f']; labels = { v: 'Vrai', f: 'Faux' }; }
                    else if (type === CONFIG.TYPES.YES_NO) { options = ['o', 'n']; labels = { o: 'Oui', n: 'Non' }; }
                    else if (type === CONFIG.TYPES.GMF) { options = ['G', 'M', 'F']; labels = { G: 'Gradin', M: 'Mobile', F: 'Fixe' }; }

                    const pillContainer = document.createElement('div');
                    pillContainer.className = 'pill-group';
                    if (isFieldDisabled) {
                        pillContainer.classList.add('disabled-group');
                    }

                    options.forEach(opt => {
                        const wrapper = document.createElement('label');
                        wrapper.className = 'pill-option';

                        const input = document.createElement('input');
                        input.type = 'radio';
                        input.name = `field-${field.colIndex}`;
                        input.value = opt;
                        if (shouldDisableInput) input.disabled = true;

                        const isChecked = val.toString().toLowerCase() === opt.toLowerCase();
                        if (isChecked) {
                            input.checked = true;
                            // Color coding for pills
                            if (type === CONFIG.TYPES.YES_NO) {
                                wrapper.classList.add(opt === 'o' ? 'selected-success' : 'selected-danger');
                            } else if (type === CONFIG.TYPES.TRUE_FALSE) {
                                wrapper.classList.add(opt === 'v' ? 'selected-success' : 'selected-danger');
                            } else {
                                wrapper.classList.add('selected');
                            }
                        }

                        input.addEventListener('change', () => {
                            // Update all siblings
                            pillContainer.querySelectorAll('.pill-option').forEach(p => {
                                p.classList.remove('selected', 'selected-success', 'selected-danger');
                            });
                            if (type === CONFIG.TYPES.YES_NO) {
                                wrapper.classList.add(opt === 'o' ? 'selected-success' : 'selected-danger');
                            } else if (type === CONFIG.TYPES.TRUE_FALSE) {
                                wrapper.classList.add(opt === 'v' ? 'selected-success' : 'selected-danger');
                            } else {
                                wrapper.classList.add('selected');
                            }
                            updateCell(field.colIndex, opt);
                        });

                        const textSpan = document.createElement('span');
                        textSpan.textContent = labels[opt] || opt;

                        wrapper.appendChild(input);
                        wrapper.appendChild(textSpan);
                        pillContainer.appendChild(wrapper);
                    });
                    group.appendChild(pillContainer);

                } else if (type === CONFIG.TYPES.DATE) {
                    const input = document.createElement('input');
                    input.type = 'date';
                    if (shouldDisableInput) {
                        input.disabled = true;
                        input.classList.add('input-disabled');
                    }

                    // Parse Date Value
                    // Parse Date Value
                    if (typeof val === 'string' && /^\d{2}\/\d{2}\/\d{4}$/.test(val)) {
                        // Handle "dd/mm/yyyy" string format specifically
                        const [d, m, y] = val.split('/');
                        input.value = `${y}-${m}-${d}`;
                    } else if (val && !isNaN(Date.parse(val))) {
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
                        if (!e.target.value) {
                            updateCell(field.colIndex, '');
                            return;
                        }
                        // Construct Local Date strictly from input components to avoid timezone shifts
                        const [y, m, d] = e.target.value.split('-').map(Number);
                        const localDate = new Date(y, m - 1, d);
                        updateCell(field.colIndex, localDate);
                    });
                    group.appendChild(input);

                } else if (type === CONFIG.TYPES.NUMBER) {
                    const input = document.createElement('input');
                    input.type = 'number';
                    input.value = val;
                    if (shouldDisableInput) {
                        input.disabled = true;
                        input.classList.add('input-disabled');
                    }

                    // Auto-fill capacité réelle from capacité annoncée
                    const qLow = (field.question || '').toLowerCase();
                    const cLow = (field.category || '').toLowerCase();
                    const isCapaciteReelle = CONFIG.AUTO_FILL.CAPACITY_TARGET.some(t => qLow.includes(t)) || cLow.includes('capacité réelle');

                    if (isCapaciteReelle && !val) {
                        const capField = schema.find(f => {
                            const fc = (f.category || '').toLowerCase();
                            const fq = (f.question || '').toLowerCase();
                            return fc === CONFIG.AUTO_FILL.CAPACITY_SOURCE || fq === CONFIG.AUTO_FILL.CAPACITY_SOURCE;
                        });
                        if (capField) {
                            const capVal = getVal(row, capField.colIndex);
                            if (capVal) {
                                if (isOptional) {
                                    // Non-applicable: just show as placeholder
                                    input.placeholder = capVal;
                                } else {
                                    // Applicable: fill with real value
                                    input.value = capVal;
                                    updateCell(field.colIndex, capVal);
                                }
                            }
                        }
                    }

                    input.addEventListener('input', (e) => updateCell(field.colIndex, e.target.value));
                    group.appendChild(input);
                } else {
                    // Default Text Array
                    const input = document.createElement('textarea');
                    input.rows = 2;
                    input.value = val;
                    if (shouldDisableInput) {
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

        // Force Date objects to strict "dd/mm/yyyy" string format
        if (value instanceof Date) {
            const day = String(value.getDate()).padStart(2, '0');
            const month = String(value.getMonth() + 1).padStart(2, '0');
            const year = value.getFullYear();
            value = `${day}/${month}/${year}`;
        }

        cell.value = value;
        // Removed numFmt assignment since we are now using explicit string

        saveEditToDB(currentRowIndex, colIndex, value);
        updateSidebarStatus(currentRowIndex);
        updateProgressBar(currentRowIndex);
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
            showAlert('L\'export a échoué. Veuillez réessayer.', '❌', 'Erreur d\'export');
        }
    });

    clearBtn.addEventListener('click', () => {
        splitView.classList.add('hidden');
        uploadSection.classList.remove('hidden');
        burgerBtn.classList.add('hidden');
        currentWorkbook = null;
        fileInput.value = '';
    });

    // 1. Force Edit Mode
    unlockBtn.addEventListener('click', async () => {
        const shouldUnlock = await showConfirm(
            "Cela déverrouillera tous les champs structurels (Bâtiments, Capacité, etc.).",
            '🔓', 'Activer l\'Édition Forcée ?'
        );
        if (shouldUnlock) {
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
            updateSidebarStatus(currentRowIndex);
            updateProgressBar(currentRowIndex);
        }
    });

    // Search Logic
    searchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();

        const categories = document.querySelectorAll('.form-category');
        categories.forEach(cat => {
            let hasVisible = false;
            const fields = cat.querySelectorAll('.field-group');

            fields.forEach(field => {
                const text = field.textContent.toLowerCase();
                if (text.includes(query)) {
                    field.classList.remove('hidden');
                    hasVisible = true;
                } else {
                    field.classList.add('hidden');
                }
            });

            if (hasVisible) cat.classList.remove('hidden');
            else cat.classList.add('hidden');
        });
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
            showAlert('Tous les champs obligatoires semblent remplis !', '✅', 'Terminé');
        }
    });

    function toggleCellGrey(colIndex, isGreyed) {
        if (!mainWorksheet || currentRowIndex === null) return;
        const row = mainWorksheet.getRow(currentRowIndex);
        const cell = row.getCell(colIndex);

        if (isGreyed) {
            cell.style = {
                ...cell.style,
                fill: CONFIG.STYLES.GREY_PATTERN.fill
            };
        } else {
            const newStyle = { ...cell.style };
            delete newStyle.fill;
            cell.style = newStyle;
        }

        saveEditToDB(currentRowIndex, colIndex, cell.value, isGreyed);
        updateSidebarStatus(currentRowIndex);
        updateProgressBar(currentRowIndex);

        // Re-render form to ensure clean state (avoids opacity stacking bugs)
        renderForm(currentRowIndex);
    }

    // Scroll-to-top button
    if (scrollTopBtn) {
        formContainer.addEventListener('scroll', () => {
            if (formContainer.scrollTop > 300) {
                scrollTopBtn.classList.add('visible');
            } else {
                scrollTopBtn.classList.remove('visible');
            }
        });
        scrollTopBtn.addEventListener('click', () => {
            formContainer.scrollTo({ top: 0, behavior: 'smooth' });
        });
    }

    initDB();
});
