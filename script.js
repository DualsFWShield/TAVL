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
            STORE_EDITS: 'edits',
            STORE_REMARKS: 'remarks'
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
    const progressBadge = document.getElementById('progress-badge');
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
    const fileTabs = document.getElementById('file-tabs');
    const fileTabsBar = document.getElementById('file-tabs-bar');
    const addFileBtn = document.getElementById('add-file-btn');
    const addFileInput = document.getElementById('add-file-input');
    const burgerBtn = document.getElementById('burger-btn');
    const mobileFab = document.getElementById('mobile-sidebar-fab');
    const sidebarOverlay = document.getElementById('sidebar-overlay');
    const sidebar = document.querySelector('.sidebar');
    const modalOverlay = document.getElementById('modal-overlay');
    const modalIcon = document.getElementById('modal-icon');
    const modalTitle = document.getElementById('modal-title');
    const modalMessage = document.getElementById('modal-message');
    const modalActions = document.getElementById('modal-actions');

    // Post-Relevé DOM elements
    const modeSwitcher = document.getElementById('mode-switcher');
    const modeReleveBtn = document.getElementById('mode-releve-btn');
    const modePostBtn = document.getElementById('mode-post-btn');
    const modeStatsBtn = document.getElementById('mode-stats-btn');
    const modeIdentityBtn = document.getElementById('mode-identity-btn');
    const postReleveSection = document.getElementById('post-releve-section');
    const statsSection = document.getElementById('stats-section');
    const identitySection = document.getElementById('identity-section');
    const dashboardMetricsRow = document.getElementById('dashboard-metrics-row');
    const dashboardFilesList = document.getElementById('dashboard-files-list');
    const dashboardStatsList = document.getElementById('dashboard-stats-list');
    const anomaliesSearchInput = document.getElementById('anomalies-search-input');
    const anomaliesFilterFile = document.getElementById('anomalies-filter-file');
    const anomaliesFeed = document.getElementById('anomalies-feed');
    const generateSampleBtn = document.getElementById('generate-sample-btn');
    const verificationList = document.getElementById('verification-list');
    const exportFilesCheckboxes = document.getElementById('export-files-checkboxes');
    const exportCategoriesCheckboxes = document.getElementById('export-categories-checkboxes');
    const exportAnomaliesXlsxBtn = document.getElementById('export-anomalies-xlsx-btn');
    const exportFilteredXlsxBtn = document.getElementById('export-filtered-xlsx-btn');

    // Post-Relevé State
    let currentMode = 'releve'; // 'releve' or 'post'
    let verificationSample = []; // Array of { sessionId, rowIndex, verified }


    /* ==========================================================================
       MODAL SYSTEM (replaces browser confirm/alert)
       ========================================================================== */
    function showModal({ icon = 'ℹ️', title = '', message = '', buttons = [] }) {
        return new Promise(resolve => {
            modalIcon.textContent = icon;
            modalTitle.textContent = title;
            if (message instanceof HTMLElement) {
                modalMessage.innerHTML = '';
                modalMessage.appendChild(message);
            } else {
                modalMessage.textContent = message;
            }
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
    
    if (mobileFab) {
        mobileFab.addEventListener('click', () => {
            if (sidebar.classList.contains('drawer-open')) closeDrawer();
            else openDrawer();
        });
    }

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
       STATE & VARIABLES (Session-based for multi-file support)
       ========================================================================== */
    const sessions = new Map(); // Map<sessionId, {id, fileName, workbook, worksheet, schema, dataRows, currentRowIndex}>
    let activeSessionId = null;
    let sessionCounter = 0;
    let db = null;
    let lastFocusedInput = null; // Track focus for smart navigation

    /** Get the active session object */
    function S() { return sessions.get(activeSessionId); }

    function generateSessionId() {
        return 'sess_' + (++sessionCounter) + '_' + Date.now();
    }

    /* ==========================================================================
       INDEXED DB PERSISTENCE (Multi-save support)
       ========================================================================== */
    const initDB = () => {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(CONFIG.DB.NAME, 4);
            request.onupgradeneeded = (e) => {
                db = e.target.result;
                if (!db.objectStoreNames.contains(CONFIG.DB.STORE_FILE)) db.createObjectStore(CONFIG.DB.STORE_FILE);
                if (!db.objectStoreNames.contains(CONFIG.DB.STORE_EDITS)) {
                    const editStore = db.createObjectStore(CONFIG.DB.STORE_EDITS, { keyPath: 'id' });
                    editStore.createIndex('sessionId', 'sessionId', { unique: false });
                }
                if (!db.objectStoreNames.contains(CONFIG.DB.STORE_REMARKS)) {
                    // Key format: sessionId::roomName
                    const remarksStore = db.createObjectStore(CONFIG.DB.STORE_REMARKS, { keyPath: 'id' });
                    remarksStore.createIndex('sessionId', 'sessionId', { unique: false });
                }
            };
            request.onsuccess = (e) => {
                db = e.target.result;
                resolve(db);
                checkSavedSessions();
            };
            request.onerror = (e) => reject(e);
        });
    };

    function saveFileToDB(sessionId, fileBuffer, fileName) {
        if (!db) return;
        const tx = db.transaction([CONFIG.DB.STORE_FILE], 'readwrite');
        tx.objectStore(CONFIG.DB.STORE_FILE).put(
            { buffer: fileBuffer, name: fileName, timestamp: Date.now() },
            sessionId
        );
    }

    function saveEditToDB(sessionId, row, col, val, isGreyed = false) {
        if (!db) return;
        const tx = db.transaction([CONFIG.DB.STORE_EDITS], 'readwrite');
        const id = `${sessionId}::${row}-${col}`;
        tx.objectStore(CONFIG.DB.STORE_EDITS).put({ id, sessionId, row, col, val, isGreyed });
    }

    function clearSessionFromDB(sessionId) {
        return new Promise(resolve => {
            if (!db) return resolve();
            // Remove file entry
            const tx1 = db.transaction([CONFIG.DB.STORE_FILE], 'readwrite');
            tx1.objectStore(CONFIG.DB.STORE_FILE).delete(sessionId);
            // Remove associated edits
            const tx2 = db.transaction([CONFIG.DB.STORE_EDITS], 'readwrite');
            const store = tx2.objectStore(CONFIG.DB.STORE_EDITS);
            const cursorReq = store.openCursor();
            cursorReq.onsuccess = (e) => {
                const cursor = e.target.result;
                if (cursor) {
                    if (cursor.value.sessionId === sessionId || cursor.key.startsWith(sessionId + '::')) {
                        cursor.delete();
                    }
                    cursor.continue();
                }
            };
            tx2.oncomplete = () => resolve();
            tx2.onerror = () => resolve();
        });
    }

    function clearAllDB() {
        if (!db) return;
        const tx = db.transaction([CONFIG.DB.STORE_FILE, CONFIG.DB.STORE_EDITS, CONFIG.DB.STORE_REMARKS], 'readwrite');
        tx.objectStore(CONFIG.DB.STORE_FILE).clear();
        tx.objectStore(CONFIG.DB.STORE_EDITS).clear();
        tx.objectStore(CONFIG.DB.STORE_REMARKS).clear();
    }

    function saveRemarkToDB(sessionId, roomName, remarqueGenerale, debriefing) {
        if (!db) return;
        const tx = db.transaction([CONFIG.DB.STORE_REMARKS], 'readwrite');
        const store = tx.objectStore(CONFIG.DB.STORE_REMARKS);
        const id = `${sessionId}::${roomName}`;
        store.put({ id, sessionId, roomName, remarqueGenerale, debriefing });
    }

    function loadRemarkFromDB(sessionId, roomName) {
        return new Promise(resolve => {
            if (!db) return resolve(null);
            const tx = db.transaction([CONFIG.DB.STORE_REMARKS], 'readonly');
            const store = tx.objectStore(CONFIG.DB.STORE_REMARKS);
            const id = `${sessionId}::${roomName}`;
            const req = store.get(id);
            req.onsuccess = () => resolve(req.result || null);
            req.onerror = () => resolve(null);
        });
    }

    function loadAllRemarksForSession(sessionId) {
        return new Promise(resolve => {
            if (!db) return resolve([]);
            const remarks = [];
            const tx = db.transaction([CONFIG.DB.STORE_REMARKS], 'readonly');
            const index = tx.objectStore(CONFIG.DB.STORE_REMARKS).index('sessionId');
            const req = index.openCursor(IDBKeyRange.only(sessionId));
            req.onsuccess = (e) => {
                const cursor = e.target.result;
                if (cursor) {
                    remarks.push(cursor.value);
                    cursor.continue();
                } else {
                    resolve(remarks);
                }
            };
            req.onerror = () => resolve([]);
        });
    }

    const savedSessionsContainer = document.getElementById('saved-sessions');
    const savedSessionsList = document.getElementById('saved-sessions-list');
    const clearAllSavesBtn = document.getElementById('clear-all-saves-btn');

    async function checkSavedSessions() {
        if (!db) return;
        const tx = db.transaction([CONFIG.DB.STORE_FILE], 'readonly');
        const store = tx.objectStore(CONFIG.DB.STORE_FILE);
        const allKeysReq = store.getAllKeys();

        allKeysReq.onsuccess = async () => {
            let keys = allKeysReq.result || [];
            // Migrate old 'current' key format
            if (keys.includes('current')) {
                keys = keys.filter(k => k !== 'current');
                const oldTx = db.transaction([CONFIG.DB.STORE_FILE], 'readonly');
                const oldReq = oldTx.objectStore(CONFIG.DB.STORE_FILE).get('current');
                oldReq.onsuccess = async () => {
                    if (oldReq.result && oldReq.result.buffer) {
                        const newId = generateSessionId();
                        const migTx = db.transaction([CONFIG.DB.STORE_FILE], 'readwrite');
                        migTx.objectStore(CONFIG.DB.STORE_FILE).put(
                            { buffer: oldReq.result.buffer, name: oldReq.result.name, timestamp: Date.now() },
                            newId
                        );
                        migTx.objectStore(CONFIG.DB.STORE_FILE).delete('current');
                        // Migrate edits
                        const edTx = db.transaction([CONFIG.DB.STORE_EDITS], 'readwrite');
                        const edStore = edTx.objectStore(CONFIG.DB.STORE_EDITS);
                        const edCursor = edStore.openCursor();
                        edCursor.onsuccess = (e) => {
                            const cursor = e.target.result;
                            if (cursor) {
                                const val = cursor.value;
                                if (!val.sessionId) {
                                    const newEdit = { ...val, id: `${newId}::${val.row}-${val.col}`, sessionId: newId };
                                    edStore.put(newEdit);
                                    edStore.delete(cursor.key);
                                }
                                cursor.continue();
                            } else {
                                // Migration done, refresh list
                                renderSavedSessionsList();
                            }
                        };
                    }
                };
                return;
            }

            renderSavedSessionsList();
        };
    }

    /** Renders saved sessions as inline cards on the upload page */
    async function renderSavedSessionsList() {
        if (!db || !savedSessionsContainer || !savedSessionsList) return;

        const tx = db.transaction([CONFIG.DB.STORE_FILE], 'readonly');
        const store = tx.objectStore(CONFIG.DB.STORE_FILE);
        const allKeysReq = store.getAllKeys();

        allKeysReq.onsuccess = async () => {
            const keys = (allKeysReq.result || []).filter(k => k !== 'current');
            if (keys.length === 0) {
                savedSessionsContainer.classList.add('hidden');
                return;
            }

            // Read all session metadata
            const items = [];
            const readTx = db.transaction([CONFIG.DB.STORE_FILE], 'readonly');
            const readStore = readTx.objectStore(CONFIG.DB.STORE_FILE);
            for (const key of keys) {
                const r = readStore.get(key);
                await new Promise(res => {
                    r.onsuccess = () => {
                        if (r.result) {
                            items.push({ id: key, name: r.result.name, timestamp: r.result.timestamp || 0 });
                        }
                        res();
                    };
                });
            }

            // Sort by most recent first
            items.sort((a, b) => b.timestamp - a.timestamp);

            savedSessionsList.innerHTML = '';
            items.forEach(item => {
                const card = document.createElement('div');
                card.className = 'saved-session-card';

                // Check if already open
                const isAlreadyOpen = Array.from(sessions.values()).some(
                    s => s.id === item.id
                );

                const icon = document.createElement('div');
                icon.className = 'saved-session-icon';
                icon.textContent = '📊';

                const info = document.createElement('div');
                info.className = 'saved-session-info';

                const name = document.createElement('div');
                name.className = 'saved-session-name';
                name.textContent = item.name;

                const date = document.createElement('div');
                date.className = 'saved-session-date';
                if (item.timestamp) {
                    const d = new Date(item.timestamp);
                    date.textContent = d.toLocaleDateString('fr-BE', {
                        day: '2-digit', month: 'short', year: 'numeric',
                        hour: '2-digit', minute: '2-digit'
                    });
                } else {
                    date.textContent = 'Date inconnue';
                }

                info.appendChild(name);
                info.appendChild(date);

                const actions = document.createElement('div');
                actions.className = 'saved-session-actions';

                if (isAlreadyOpen) {
                    const openBadge = document.createElement('span');
                    openBadge.style.cssText = 'font-size:0.72rem;color:var(--success-color);font-weight:600;padding:0.3rem 0.6rem;';
                    openBadge.textContent = '✓ Ouvert';
                    actions.appendChild(openBadge);
                } else {
                    const restoreBtn = document.createElement('button');
                    restoreBtn.className = 'saved-session-restore';
                    restoreBtn.textContent = 'Ouvrir';
                    restoreBtn.addEventListener('click', async (e) => {
                        e.stopPropagation();
                        await restoreSessions([item.id]);
                        renderSavedSessionsList(); // Refresh to show "Ouvert" badge
                    });
                    actions.appendChild(restoreBtn);
                }

                // Merge Button
                const mergeBtn = document.createElement('button');
                mergeBtn.className = 'saved-session-merge';
                mergeBtn.textContent = '🔗';
                mergeBtn.title = 'Fusionner avec une autre sauvegarde';
                mergeBtn.addEventListener('click', async (e) => {
                    e.stopPropagation();
                    
                    const otherSaves = items.filter(i => i.id !== item.id && i.name === item.name);
                    if (otherSaves.length === 0) {
                        showAlert('Aucune autre sauvegarde disponible pour la fusion.', 'ℹ️', 'Fusion impossible');
                        return;
                    }

                    const selectHtml = document.createElement('div');
                    selectHtml.style.cssText = 'text-align:left; font-size:0.85rem;';
                    selectHtml.innerHTML = `<p style="margin-bottom:0.5rem;color:var(--text-secondary)">Sélectionnez la sauvegarde dont vous souhaitez importer les données dans <strong>${item.name}</strong> :</p>`;
                    
                    const select = document.createElement('select');
                    select.style.cssText = 'width:100%; padding:0.5rem; border-radius:6px; background:var(--background-color); color:var(--text-primary); border:1px solid var(--border-color);';
                    otherSaves.forEach(other => {
                        const opt = document.createElement('option');
                        opt.value = other.id;
                        const d = other.timestamp ? new Date(other.timestamp).toLocaleDateString('fr-BE', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' }) : '';
                        opt.textContent = `${other.name} (${d})`;
                        select.appendChild(opt);
                    });
                    selectHtml.appendChild(select);
                    selectHtml.innerHTML += `<p style="margin-top:0.75rem; font-size:0.75rem; color:var(--danger-color)">⚠️ En cas de conflit, les données de la sauvegarde sélectionnée écraseront celles de la sauvegarde actuelle.</p>`;
                    // Need to re-attach select because innerHTML destroyed it
                    const wrapper = document.createElement('div');
                    wrapper.appendChild(selectHtml);
                    
                    // Small hack to rebuild the select properly since innerHTML above wiped the object reference
                    const finalSelectHtml = document.createElement('div');
                    finalSelectHtml.style.cssText = 'text-align:left; font-size:0.85rem;';
                    
                    const p1 = document.createElement('p');
                    p1.style.cssText = 'margin-bottom:0.5rem;color:var(--text-secondary)';
                    p1.innerHTML = `Sélectionnez la sauvegarde à importer dans <strong>${item.name}</strong> :`;
                    finalSelectHtml.appendChild(p1);
                    
                    finalSelectHtml.appendChild(select);
                    
                    const p2 = document.createElement('p');
                    p2.style.cssText = 'margin-top:0.75rem; font-size:0.75rem; color:var(--danger-color)';
                    p2.textContent = '⚠️ En cas de conflit sur une même case, les données de la sauvegarde importée remplaceront celles de la sauvegarde cible.';
                    finalSelectHtml.appendChild(p2);

                    const confirm = await showModal({
                        icon: '🔗',
                        title: 'Fusionner des sauvegardes',
                        message: finalSelectHtml,
                        buttons: [
                            { label: 'Annuler', value: false },
                            { label: 'Fusionner', value: true, primary: true }
                        ]
                    });

                    if (confirm) {
                        const sourceId = select.value;
                        // Fetch all edits from source
                        const sourceEdits = await new Promise(resolve => {
                            const edits = [];
                            const tx = db.transaction([CONFIG.DB.STORE_EDITS], 'readonly');
                            const cursorReq = tx.objectStore(CONFIG.DB.STORE_EDITS).openCursor();
                            cursorReq.onsuccess = (ev) => {
                                const cursor = ev.target.result;
                                if (cursor) {
                                    if (cursor.value.sessionId === sourceId || cursor.key.startsWith(sourceId + '::')) {
                                        edits.push(cursor.value);
                                    }
                                    cursor.continue();
                                } else {
                                    resolve(edits);
                                }
                            };
                        });
                        
                        // Push edits to target
                        const txWrite = db.transaction([CONFIG.DB.STORE_EDITS], 'readwrite');
                        const storeWrite = txWrite.objectStore(CONFIG.DB.STORE_EDITS);
                        sourceEdits.forEach(edit => {
                            const newId = `${item.id}::${edit.row}-${edit.col}`;
                            storeWrite.put({ ...edit, id: newId, sessionId: item.id });
                        });

                        txWrite.oncomplete = () => {
                            showAlert('Les sauvegardes ont été fusionnées avec succès.', '✅', 'Fusion terminée');
                            if (isAlreadyOpen) {
                                // If target is currently open, we should technically reload its edits. 
                                // The easiest is to instruct the user or auto-reload it.
                                const sess = sessions.get(item.id);
                                if (sess) {
                                    // Apply edits to memory immediately
                                    sourceEdits.forEach(edit => {
                                        if(sess.worksheet) {
                                            const cell = sess.worksheet.getRow(edit.row).getCell(edit.col);
                                            cell.value = edit.val;
                                            if (edit.isGreyed) cell.style = { ...cell.style, fill: CONFIG.STYLES.GREY_PATTERN.fill };
                                        }
                                    });
                                    if (activeSessionId === item.id) {
                                        renderSidebar();
                                        if (sess.currentRowIndex) renderForm(sess.currentRowIndex);
                                    }
                                }
                            }
                        };
                    }
                });
                actions.appendChild(mergeBtn);

                const deleteBtn = document.createElement('button');
                deleteBtn.className = 'saved-session-delete';
                deleteBtn.textContent = '🗑';
                deleteBtn.title = 'Supprimer cette sauvegarde';
                deleteBtn.addEventListener('click', async (e) => {
                    e.stopPropagation();
                    clearSessionFromDB(item.id);
                    // Small delay for IDB transaction to complete
                    setTimeout(() => renderSavedSessionsList(), 100);
                });
                actions.appendChild(deleteBtn);

                card.appendChild(icon);
                card.appendChild(info);
                card.appendChild(actions);

                // Clicking the card itself also restores (if not already open)
                if (!isAlreadyOpen) {
                    card.addEventListener('click', async () => {
                        await restoreSessions([item.id]);
                        renderSavedSessionsList();
                    });
                }

                savedSessionsList.appendChild(card);
            });

            savedSessionsContainer.classList.remove('hidden');
        };
    }

    // Clear all saves button
    if (clearAllSavesBtn) {
        clearAllSavesBtn.addEventListener('click', async () => {
            const confirmed = await showConfirm(
                'Supprimer toutes les sauvegardes ? Cette action est irréversible.',
                '🗑', 'Tout effacer'
            );
            if (confirmed) {
                clearAllDB();
                setTimeout(() => renderSavedSessionsList(), 100);
            }
        });
    }

    // Open all saves button
    const openAllSavesBtn = document.getElementById('open-all-saves-btn');
    if (openAllSavesBtn) {
        openAllSavesBtn.addEventListener('click', async () => {
            if (!db) return;
            const tx = db.transaction([CONFIG.DB.STORE_FILE], 'readonly');
            const store = tx.objectStore(CONFIG.DB.STORE_FILE);
            const allKeysReq = store.getAllKeys();
            
            allKeysReq.onsuccess = async () => {
                const keys = (allKeysReq.result || []).filter(k => k !== 'current');
                // Filter out keys that are already open
                const keysToOpen = keys.filter(k => !Array.from(sessions.values()).some(s => s.id === k));
                if (keysToOpen.length > 0) {
                    await restoreSessions(keysToOpen);
                    renderSavedSessionsList();
                } else {
                    showAlert('Toutes les sauvegardes sont déjà ouvertes.', 'ℹ️', 'Information');
                }
            };
        });
    }

    async function restoreSessions(sessionIds) {
        for (const sessId of sessionIds) {
            const fileTx = db.transaction([CONFIG.DB.STORE_FILE], 'readonly');
            const fileReq = fileTx.objectStore(CONFIG.DB.STORE_FILE).get(sessId);

            await new Promise((resolve) => {
                fileReq.onsuccess = async () => {
                    if (!fileReq.result || !fileReq.result.buffer) { resolve(); return; }

                    const blob = new Blob([fileReq.result.buffer]);
                    const file = new File([blob], fileReq.result.name);
                    await handleFile(file, false, sessId); // Restore with original sessionId

                    // Restore Edits for this session
                    const session = sessions.get(sessId);
                    if (!session) { resolve(); return; }

                    const edTx = db.transaction([CONFIG.DB.STORE_EDITS], 'readonly');
                    const cursorReq = edTx.objectStore(CONFIG.DB.STORE_EDITS).openCursor();
                    cursorReq.onsuccess = (e) => {
                        const cursor = e.target.result;
                        if (cursor) {
                            const edit = cursor.value;
                            if (edit.sessionId === sessId || cursor.key.startsWith(sessId + '::')) {
                                const r = session.worksheet.getRow(edit.row);
                                const c = r.getCell(edit.col);
                                c.value = edit.val;
                                if (edit.isGreyed) {
                                    c.style = { ...c.style, fill: CONFIG.STYLES.GREY_PATTERN.fill };
                                }
                            }
                            cursor.continue();
                        } else {
                            // All edits restored for this session, recalculate sidebar
                            if (sessId === activeSessionId) renderSidebar();
                            resolve();
                        }
                    };
                };
            });
        }
    }

    /* ==========================================================================
       FILE HANDLING & PARSING
       ========================================================================== */

    // Helper to safely get cell value
    const getVal = (row, colIndex) => {
        const cell = row.getCell(colIndex);
        if (!cell || cell.value === null) return '';
        if (cell.value instanceof Date) {
            return cell.value.toLocaleDateString('fr-FR');
        }
        if (typeof cell.value === 'object') {
            if (cell.value.richText) return cell.value.richText.map(t => t.text).join('');
            if (cell.value.text) return cell.value.text;
            if (cell.value.result !== undefined) {
                if (cell.value.result instanceof Date) return cell.value.result.toLocaleDateString('fr-FR');
                return cell.value.result.toString();
            }
        }
        return cell.value.toString();
    };

    let _replaceAllFlag = false;

    async function handleFile(file, isNewUpload = false, existingSessionId = null) {
        try {
            let arrayBuffer = await file.arrayBuffer();
            let sessionId = existingSessionId;

            if (isNewUpload && !sessionId) {
                // Check if already open in a tab — let the DB modal handle it if they upload again
                // so they have the choice to replace/merge.

                // Check if a save with the same fileName exists in DB
                if (db) {
                    const existingSave = await new Promise(resolve => {
                        const tx = db.transaction([CONFIG.DB.STORE_FILE], 'readonly');
                        const store = tx.objectStore(CONFIG.DB.STORE_FILE);
                        const cursorReq = store.openCursor();
                        cursorReq.onsuccess = (e) => {
                            const cursor = e.target.result;
                            if (cursor) {
                                if (cursor.value && cursor.value.name === file.name) {
                                    resolve({ key: cursor.key, data: cursor.value });
                                } else {
                                    cursor.continue();
                                }
                            } else {
                                resolve(null);
                            }
                        };
                    });

                    if (existingSave) {
                        // Get edits for existing save
                        const savedEdits = await new Promise(resolve => {
                            const edits = [];
                            const tx = db.transaction([CONFIG.DB.STORE_EDITS], 'readonly');
                            const cursorReq = tx.objectStore(CONFIG.DB.STORE_EDITS).openCursor();
                            cursorReq.onsuccess = (e) => {
                                const cursor = e.target.result;
                                if (cursor) {
                                    if (cursor.value.sessionId === existingSave.key || cursor.key.startsWith(existingSave.key + '::')) {
                                        edits.push(cursor.value);
                                    }
                                    cursor.continue();
                                } else {
                                    resolve(edits);
                                }
                            };
                        });

                        // Calculate stats for existing save
                        const savedWb = new ExcelJS.Workbook();
                        await savedWb.xlsx.load(existingSave.data.buffer.slice(0));
                        const savedWs = savedWb.worksheets[0];
                        const { schema: savedSchema, dataRows: savedRows } = parseMatrixStructure(savedWs);

                        // Apply edits to virtual worksheet
                        savedEdits.forEach(edit => {
                            const cell = savedWs.getRow(edit.row).getCell(edit.col);
                            cell.value = edit.val;
                            if (edit.isGreyed) cell.style = { ...cell.style, fill: CONFIG.STYLES.GREY_PATTERN.fill };
                        });

                        let savedFilledCount = 0;
                        const savedAuditoires = new Set();
                        savedRows.forEach(dr => {
                            const row = savedWs.getRow(dr.rowIndex);
                            let hasEdits = false;
                            savedSchema.forEach(f => {
                                if (!f.question) return;
                                const v = getVal(row, f.colIndex);
                                if (v && v.toString().trim() !== '') {
                                    savedFilledCount++;
                                    hasEdits = true;
                                }
                            });
                            if (hasEdits) savedAuditoires.add(dr.name);
                        });

                        // Calculate stats for the new file
                        const tempWb = new ExcelJS.Workbook();
                        await tempWb.xlsx.load(arrayBuffer.slice(0)); // use a copy
                        const tempWs = tempWb.worksheets[0];
                        const { schema: tempSchema, dataRows: tempRows } = parseMatrixStructure(tempWs);
                        let newFilledCount = 0;
                        const newAuditoires = new Set();
                        tempRows.forEach(dr => {
                            const row = tempWs.getRow(dr.rowIndex);
                            let hasEdits = false;
                            tempSchema.forEach(f => {
                                if (!f.question) return;
                                const v = getVal(row, f.colIndex);
                                if (v && v.toString().trim() !== '') {
                                    newFilledCount++;
                                    hasEdits = true;
                                }
                            });
                            if (hasEdits) newAuditoires.add(dr.name);
                        });

                        const savedDate = existingSave.data.timestamp
                            ? new Date(existingSave.data.timestamp).toLocaleDateString('fr-BE', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })
                            : 'inconnue';

                        const formatAudList = (set) => {
                            const arr = Array.from(set);
                            if (arr.length === 0) return 'Aucun';
                            if (arr.length <= 4) return arr.join(', ');
                            return arr.slice(0, 4).join(', ') + ` et ${arr.length - 4} autre(s)`;
                        };

                        // Build comparison UI
                        const cmpDiv = document.createElement('div');
                        cmpDiv.style.cssText = 'text-align:left;font-size:0.85rem;line-height:1.6;';
                        cmpDiv.innerHTML = `
                            <div style="display:flex;gap:0.75rem;margin-bottom:0.5rem">
                                <div style="flex:1;padding:0.6rem;border-radius:8px;background:rgba(99,102,241,0.08);border:1px solid rgba(99,102,241,0.2)">
                                    <div style="font-weight:700;font-size:0.78rem;color:var(--primary-light);margin-bottom:0.3rem">💾 Sauvegarde existante</div>
                                    <div style="font-size:0.78rem;color:var(--text-secondary)">Date : ${savedDate}</div>
                                    <div style="font-size:0.78rem;color:var(--text-secondary)">${savedFilledCount} champ(s) rempli(s)</div>
                                    <div style="font-size:0.7rem;color:var(--text-muted);margin-top:0.3rem"><strong>Auditoires :</strong> ${formatAudList(savedAuditoires)}</div>
                                </div>
                                <div style="flex:1;padding:0.6rem;border-radius:8px;background:rgba(34,211,238,0.06);border:1px solid rgba(34,211,238,0.15)">
                                    <div style="font-weight:700;font-size:0.78rem;color:var(--accent-color);margin-bottom:0.3rem">📄 Nouveau fichier</div>
                                    <div style="font-size:0.78rem;color:var(--text-secondary)">${newFilledCount} champ(s) rempli(s)</div>
                                    <div style="font-size:0.7rem;color:var(--text-muted);margin-top:0.3rem"><strong>Auditoires :</strong> ${formatAudList(newAuditoires)}</div>
                                </div>
                            </div>
                            <div style="font-size:0.78rem;color:var(--text-muted);text-align:center">Que souhaitez-vous faire ?</div>
                        `;

                        let choice;
                        if (_replaceAllFlag) {
                            choice = 'replace';
                        } else {
                            choice = await showModal({
                                icon: '⚖️',
                                title: `"${file.name}" existe déjà`,
                                message: cmpDiv,
                                buttons: [
                                    { label: 'Annuler', value: 'cancel' },
                                    { label: 'Garder les deux', value: 'keep' },
                                    { label: 'Remplacer', value: 'replace' },
                                    { label: 'Tout remplacer', value: 'replace_all' },
                                    { label: 'Fusionner', value: 'merge', primary: true }
                                ]
                            });
                        }

                        if (choice === 'cancel') return;
                        if (choice === 'replace_all') {
                            _replaceAllFlag = true;
                            choice = 'replace';
                        }
                        if (choice === 'replace') {
                            await clearSessionFromDB(existingSave.key);
                            sessionId = existingSave.key;
                        }
                        if (choice === 'merge') {
                            sessionId = existingSave.key;
                            // Inject values from new file into existing save's DB edits
                            tempRows.forEach(dr => {
                                const row = tempWs.getRow(dr.rowIndex);
                                tempSchema.forEach(f => {
                                    if (!f.question) return;
                                    const val = getVal(row, f.colIndex);
                                    if (val && val.toString().trim() !== '') {
                                        // Note: this overwrites any previous edit for the same cell
                                        saveEditToDB(sessionId, dr.rowIndex, f.colIndex, val, false);
                                    }
                                });
                            });
                            // Use the old base file so we just stack edits on top
                            arrayBuffer = existingSave.data.buffer.slice(0);
                        }
                        // 'keep' → sessionId stays null, a new one will be generated
                    }
                }

                if (!sessionId) sessionId = generateSessionId();
                saveFileToDB(sessionId, arrayBuffer, file.name);
            }

            if (!sessionId) sessionId = generateSessionId();

            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            const worksheet = workbook.worksheets[0];

            const { schema: parsedSchema, dataRows: parsedRows } = parseMatrixStructure(worksheet);

            // Create session
            const session = {
                id: sessionId,
                fileName: file.name,
                workbook: workbook,
                worksheet: worksheet,
                schema: parsedSchema,
                dataRows: parsedRows,
                currentRowIndex: null
            };
            sessions.set(sessionId, session);

            // Switch to this session
            activeSessionId = sessionId;
            renderTabs();
            renderSidebar();

            // Reset Form View
            formContainer.innerHTML = '<div class="empty-state">Veuillez sélectionner un auditoire dans la liste à gauche.</div>';
            currentAuditoireTitle.textContent = 'Sélectionnez un auditoire';
            if (searchContainer) searchContainer.style.display = 'none';
            if (progressBarContainer) progressBarContainer.style.display = 'none';
            if (progressBadge) progressBadge.style.display = 'none';

            uploadSection.classList.add('hidden');
            fileTabsBar.classList.remove('hidden');
            if (modeSwitcher) modeSwitcher.classList.remove('hidden');

            if (currentMode === 'post') {
                splitView.classList.add('hidden');
                postReleveSection.classList.remove('hidden');
                burgerBtn.classList.add('hidden');
                renderPostReleveDashboard();
            } else {
                splitView.classList.remove('hidden');
                postReleveSection.classList.add('hidden');
                burgerBtn.classList.remove('hidden');
            }

        } catch (error) {
            console.error('Error parsing file:', error);
            showAlert('Structure du fichier non reconnue ou fichier invalide.', '❌', 'Erreur');
        }
    }

    /**
     * Parses the Excel Header Rows to Determine Schema.
     * Returns { schema, dataRows } instead of setting globals.
     */
    function parseMatrixStructure(sheet) {
        const schemaResult = [];
        const dataRowsResult = [];

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
                schemaResult.push({
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
                dataRowsResult.push({ rowIndex: rowIndex, name: name });
            }
        });

        return { schema: schemaResult, dataRows: dataRowsResult };
    }

    /* ==========================================================================
       TAB MANAGEMENT
       ========================================================================== */

    function renderTabs() {
        fileTabs.innerHTML = '';
        sessions.forEach((session) => {
            const tab = document.createElement('div');
            tab.className = 'file-tab' + (session.id === activeSessionId ? ' active' : '');
            tab.dataset.sessionId = session.id;

            const nameSpan = document.createElement('span');
            nameSpan.className = 'file-tab-name';
            nameSpan.textContent = session.fileName.replace(/\.xlsx$/i, '');
            nameSpan.title = session.fileName;
            tab.appendChild(nameSpan);

            const closeBtn = document.createElement('button');
            closeBtn.className = 'file-tab-close';
            closeBtn.textContent = '✕';
            closeBtn.title = 'Fermer';
            closeBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                closeSession(session.id);
            });
            tab.appendChild(closeBtn);

            tab.addEventListener('click', () => switchToSession(session.id));
            fileTabs.appendChild(tab);
        });
    }

    function switchToSession(sessionId) {
        if (sessionId === activeSessionId) return;
        activeSessionId = sessionId;
        const session = S();
        if (!session) return;

        renderTabs();

        if (currentMode === 'post') {
            renderPostReleveDashboard();
            return;
        }

        renderSidebar();

        // Restore form state
        if (session.currentRowIndex) {
            const idx = session.dataRows.findIndex(r => r.rowIndex === session.currentRowIndex);
            if (idx !== -1) {
                selectAuditoire(idx);
                return;
            }
        }
        // No auditoire selected yet
        formContainer.innerHTML = '<div class="empty-state">Veuillez sélectionner un auditoire dans la liste à gauche.</div>';
        currentAuditoireTitle.textContent = 'Sélectionnez un auditoire';
        if (searchContainer) searchContainer.style.display = 'none';
        if (progressBarContainer) progressBarContainer.style.display = 'none';
        if (progressBadge) progressBadge.style.display = 'none';
    }

    async function closeSession(sessionId) {
        const session = sessions.get(sessionId);
        if (!session) return;

        const shouldClose = await showConfirm(
            `Fermer "${session.fileName}" ? La sauvegarde sera conservée.`,
            '📁', 'Fermer le fichier'
        );
        if (!shouldClose) return;

        // Keep save in DB (don't call clearSessionFromDB)
        sessions.delete(sessionId);

        if (sessions.size === 0) {
            activeSessionId = null;
            if (modeSwitcher) modeSwitcher.classList.add('hidden');
            switchMode('releve');
            splitView.classList.add('hidden');
            postReleveSection.classList.add('hidden');
            fileTabsBar.classList.add('hidden');
            sidebar.classList.add('hidden');
            burgerBtn.classList.add('hidden');
            if (mobileFab) mobileFab.classList.remove('visible');
            fileTabs.innerHTML = '';
            renderSavedSessionsList(); // Refresh list to show closed session back
            return;
        }

        // Switch to another tab if we closed the active one
        if (sessionId === activeSessionId) {
            const firstKey = sessions.keys().next().value;
            switchToSession(firstKey);
        } else {
            renderTabs();
        }
    }

    /* ==========================================================================
       UI RENDERING
       ========================================================================== */

    async function renderSidebar() {
        auditoireList.innerHTML = '';
        const session = S();
        if (!session) return;
        
        const remarksArray = await loadAllRemarksForSession(session.id);
        const remarksMap = new Set(remarksArray.filter(r => r.remarqueGenerale || r.debriefing).map(r => r.roomName));

        session.dataRows.forEach((item, index) => {
            const li = document.createElement('li');
            li.className = 'sidebar-item';
            li.dataset.rowIndex = item.rowIndex;

            const spanName = document.createElement('span');
            spanName.textContent = item.name;
            li.appendChild(spanName);

            const badgesContainer = document.createElement('div');
            badgesContainer.style.display = 'flex';
            badgesContainer.style.alignItems = 'center';
            badgesContainer.style.gap = '0.25rem';

            if (remarksMap.has(item.name)) {
                const remarkIcon = document.createElement('span');
                remarkIcon.innerHTML = '💬';
                remarkIcon.style.fontSize = '0.8rem';
                remarkIcon.title = 'Cet auditoire possède une remarque/débriefing';
                badgesContainer.appendChild(remarkIcon);
            }

            const check = document.createElement('span');
            check.className = 'status-indicator';
            check.innerHTML = '✔';

            const name = item.name.toLowerCase();
            const threshold = name.includes('hall, sanitaire') ? 40 : 60;

            const percentage = calculateCompletion(item.rowIndex);
            if (percentage >= threshold) check.classList.add('visible');

            badgesContainer.appendChild(check);
            li.appendChild(badgesContainer);

            li.onclick = () => selectAuditoire(index);
            auditoireList.appendChild(li);
        });
    }

    function getRowCompletion(session, rowIndex) {
        if (!session || !session.worksheet) return { percentage: 0, filled: 0, total: 0 };
        const row = session.worksheet.getRow(rowIndex);
        let total = 0;
        let filled = 0;

        session.schema.forEach(field => {
            if (!field.question) return;

            const cell = row.getCell(field.colIndex);

            let isExempt = false;
            if (cell && cell.style && cell.style.fill) {
                const f = cell.style.fill;
                if (f.type === 'pattern' && f.pattern && f.pattern !== 'none' && f.pattern !== 'solid') {
                    isExempt = true;
                }
            }

            if (isExempt) return;

            total++;
            const val = getVal(row, field.colIndex);
            if (val && val.toString().trim() !== '') {
                filled++;
            }
        });

        const pct = total === 0 ? 0 : Math.round((filled / total) * 100);
        return { percentage: pct, filled, total };
    }

    function calculateCompletion(rowIndex) {
        const session = S();
        return getRowCompletion(session, rowIndex).percentage;
    }


    function updateSidebarStatus(rowIndex) {
        const li = Array.from(auditoireList.children).find(el => parseInt(el.dataset.rowIndex) === rowIndex);
        if (li) {
            const session = S();
            const rowData = session ? session.dataRows.find(r => r.rowIndex === rowIndex) : null;
            const name = rowData ? rowData.name.toLowerCase() : '';
            const threshold = name.includes('hall, sanitaire') ? 40 : 60;

            const percentage = calculateCompletion(rowIndex);
            const check = li.querySelector('.status-indicator');
            if (percentage >= threshold) check.classList.add('visible');
            else check.classList.remove('visible');
        }
    }

    function selectAuditoire(index) {
        const session = S();
        if (!session) return;
        document.querySelectorAll('.sidebar-item').forEach(el => el.classList.remove('active'));
        auditoireList.children[index].classList.add('active');

        const item = session.dataRows[index];
        session.currentRowIndex = item.rowIndex;
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
        if (progressBadge) progressBadge.style.display = '';
        closeDrawer();
    }

    function updateProgressBar(rowIndex) {
        const session = S();
        if (!session || !session.worksheet || !progressFill) return;
        const row = session.worksheet.getRow(rowIndex);
        let total = 0, filled = 0;
        session.schema.forEach(field => {
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
        if (progressBadge) progressBadge.textContent = filled + '/' + total;
    }

    function renderForm(rowIndex) {
        const session = S();
        if (!session) return;
        formContainer.innerHTML = '';
        const row = session.worksheet.getRow(rowIndex);

        // Extract structurally read-only fields for ID Card
        const forceEditMode = document.body.classList.contains('force-edit-mode');
        const idFields = [];
        const normalFields = [];

        session.schema.forEach(field => {
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
                        const capField = session.schema.find(f => {
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
        const session = S();
        if (!session || !session.currentRowIndex) return;
        const row = session.worksheet.getRow(session.currentRowIndex);
        const cell = row.getCell(colIndex);

        // Force Date objects to strict "dd/mm/yyyy" string format
        if (value instanceof Date) {
            const day = String(value.getDate()).padStart(2, '0');
            const month = String(value.getMonth() + 1).padStart(2, '0');
            const year = value.getFullYear();
            value = `${day}/${month}/${year}`;
        }

        cell.value = value;

        saveEditToDB(session.id, session.currentRowIndex, colIndex, value);
        updateSidebarStatus(session.currentRowIndex);
        updateProgressBar(session.currentRowIndex);
    }

    /* ==========================================================================
       INTERACTIVE FEATURES
       ========================================================================== */

    // Export Logic
    exportXlsxBtn.addEventListener('click', async () => {
        const session = S();
        if (!session || !session.workbook) return;
        try {
            const buffer = await session.workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            // Use original filename with _Completed suffix
            const baseName = session.fileName.replace(/\.xlsx$/i, '');
            a.download = `${baseName}_Completed.xlsx`;
            a.click();
            window.URL.revokeObjectURL(url);
        } catch (error) {
            console.error(error);
            showAlert('L\'export a échoué. Veuillez réessayer.', '❌', 'Erreur d\'export');
        }
    });

    clearBtn.addEventListener('click', () => {
        if (activeSessionId) {
            closeSession(activeSessionId);
        }
    });

    // 1. Force Edit Mode
    unlockBtn.addEventListener('click', async () => {
        const shouldUnlock = await showConfirm(
            "Cela déverrouillera tous les champs structurels (Bâtiments, Capacité, etc.).",
            '🔓', 'Activer l\'Édition Forcée ?'
        );
        if (shouldUnlock) {
            document.body.classList.toggle('force-edit-mode');
            const session = S();
            if (session && session.currentRowIndex) renderForm(session.currentRowIndex);
        }
    });

    // 2. Fill Defaults
    fillDefaultsBtn.addEventListener('click', () => {
        const session = S();
        if (!session || !session.currentRowIndex || !session.worksheet) return;

        const row = session.worksheet.getRow(session.currentRowIndex);
        let editsMade = 0;

        // Find Value of "Announced Capacity"
        let capaAnnonceVal = '';
        const capaField = session.schema.find(f => {
            const c = (f.category || '').toLowerCase();
            const q = (f.question || '').toLowerCase();
            return c === CONFIG.AUTO_FILL.CAPACITY_SOURCE || q === CONFIG.AUTO_FILL.CAPACITY_SOURCE;
        });
        if (capaField) {
            capaAnnonceVal = getVal(row, capaField.colIndex);
        }

        session.schema.forEach(field => {
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
            renderForm(session.currentRowIndex);
            updateSidebarStatus(session.currentRowIndex);
            updateProgressBar(session.currentRowIndex);
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

    // Drag/Drop Listeners (supports multiple files)
    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
    dropZone.addEventListener('drop', async (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.xlsx'));
        _replaceAllFlag = false;
        for (const f of files) {
            await handleFile(f, true);
        }
        _replaceAllFlag = false;
    });
    fileInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        _replaceAllFlag = false;
        for (const f of files) {
            await handleFile(f, true);
        }
        _replaceAllFlag = false;
        e.target.value = '';
    });

    // Add File button (in tab bar)
    addFileBtn.addEventListener('click', () => addFileInput.click());
    addFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) handleFile(e.target.files[0], true);
        e.target.value = '';
    });


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
        const session = S();
        if (!session || !session.worksheet || session.currentRowIndex === null) return;
        const row = session.worksheet.getRow(session.currentRowIndex);
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

        saveEditToDB(session.id, session.currentRowIndex, colIndex, cell.value, isGreyed);
        updateSidebarStatus(session.currentRowIndex);
        updateProgressBar(session.currentRowIndex);

        // Re-render form to ensure clean state (avoids opacity stacking bugs)
        renderForm(session.currentRowIndex);
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

    /* ==========================================================================
       POST-RELEVE CORE FUNCTIONALITY
       ========================================================================== */

    function switchMode(mode) {
        currentMode = mode;
        modeReleveBtn.classList.remove('active');
        modePostBtn.classList.remove('active');
        if (modeStatsBtn) modeStatsBtn.classList.remove('active');
        if (modeIdentityBtn) modeIdentityBtn.classList.remove('active');

        splitView.classList.add('hidden');
        postReleveSection.classList.add('hidden');
        if (statsSection) statsSection.classList.add('hidden');
        if (identitySection) identitySection.classList.add('hidden');

        if (mode === 'releve') {
            modeReleveBtn.classList.add('active');
            splitView.classList.remove('hidden');
            fileTabsBar.classList.remove('hidden');
            
            // Re-render sidebar/form for current active session
            renderSidebar();
            const session = S();
            if (session && session.currentRowIndex) {
                const idx = session.dataRows.findIndex(r => r.rowIndex === session.currentRowIndex);
                if (idx !== -1) {
                    selectAuditoire(idx);
                }
            } else {
                formContainer.innerHTML = '<div class="empty-state">Veuillez sélectionner un auditoire dans la liste à gauche.</div>';
                currentAuditoireTitle.textContent = 'Sélectionnez un auditoire';
                if (searchContainer) searchContainer.style.display = 'none';
                if (progressBarContainer) progressBarContainer.style.display = 'none';
                if (progressBadge) progressBadge.style.display = 'none';
            }
            if (window.innerWidth <= 768) {
                burgerBtn.classList.remove('hidden');
                if (mobileFab) mobileFab.classList.add('visible');
            }
        } else if (mode === 'post') {
            modePostBtn.classList.add('active');
            postReleveSection.classList.remove('hidden');
            burgerBtn.classList.add('hidden');
            if (mobileFab) mobileFab.classList.remove('visible');
            closeDrawer();
            renderPostReleveDashboard();
        } else if (mode === 'stats') {
            if (modeStatsBtn) modeStatsBtn.classList.add('active');
            if (statsSection) statsSection.classList.remove('hidden');
            burgerBtn.classList.add('hidden');
            if (mobileFab) mobileFab.classList.remove('visible');
            closeDrawer();
            renderStatsGraphs();
        } else if (mode === 'identity') {
            if (modeIdentityBtn) modeIdentityBtn.classList.add('active');
            if (identitySection) identitySection.classList.remove('hidden');
            burgerBtn.classList.add('hidden');
            if (mobileFab) mobileFab.classList.remove('visible');
            closeDrawer();
            renderIdentityCards();
        }
    }

    function createDonutChart(title, valueStr, labelStr, dataArray, colors) {
        let conicStops = [];
        let currentPct = 0;
        let legendsHtml = '';
        
        let total = dataArray.reduce((acc, item) => acc + item.value, 0);
        
        dataArray.forEach((item, index) => {
            const pct = total === 0 ? 0 : (item.value / total) * 100;
            const color = colors[index % colors.length];
            if (total > 0) {
                conicStops.push(`${color} ${currentPct}% ${currentPct + pct}%`);
            }
            currentPct += pct;
            
            legendsHtml += `
                <div class="donut-legend-item">
                    <div style="display:flex; align-items:center; overflow:hidden;">
                        <div class="donut-legend-color" style="background:${color};"></div>
                        <span class="donut-legend-text" title="${item.label}">${item.label}</span>
                    </div>
                    <span class="donut-legend-value">${item.value} (${Math.round(pct)}%)</span>
                </div>
            `;
        });
        
        const gradient = total > 0 ? `conic-gradient(${conicStops.join(', ')})` : `conic-gradient(var(--surface-hover) 0% 100%)`;
        
        return `
            <div class="donut-chart-container">
                <h3>${title}</h3>
                <div class="donut-chart" style="background: ${gradient};">
                    <div class="donut-inner">
                        <span class="donut-inner-value">${valueStr}</span>
                        <span class="donut-inner-label">${labelStr}</span>
                    </div>
                </div>
                <div class="donut-legend">
                    ${total > 0 ? legendsHtml : '<div class="donut-legend-item" style="justify-content:center;">Aucune donnée</div>'}
                </div>
            </div>
        `;
    }

    function renderStatsGraphs() {
        const statsChartsContainer = document.getElementById('stats-charts-container');
        if (!statsChartsContainer) return;

        if (sessions.size === 0) {
            statsChartsContainer.innerHTML = `
                <div class="empty-dashboard-state">
                    <span style="font-size: 3rem;">📊</span>
                    <h2>Aucun fichier actif</h2>
                    <p>Veuillez charger des relevés Excel pour voir les graphiques.</p>
                </div>
            `;
            return;
        }

        let totalRooms = 0;
        let completedRooms = 0;
        let roomsWithAnomalies = 0;
        let totalAnomalies = 0;
        
        const problemCats = new Map();
        const poolAnomalies = new Map();
        
        sessions.forEach(session => {
            let sessionAnomalies = 0;
            
            session.dataRows.forEach(item => {
                totalRooms++;
                const nameNorm = item.name.toLowerCase();
                const threshold = nameNorm.includes('hall, sanitaire') ? 40 : 60;
                if (getRowCompletion(session, item.rowIndex).percentage >= threshold) {
                    completedRooms++;
                }

                const roomAnomalies = getRowAnomalies(session, item.rowIndex);
                if (roomAnomalies.length > 0) {
                    roomsWithAnomalies++;
                    sessionAnomalies += roomAnomalies.length;
                    totalAnomalies += roomAnomalies.length;
                }
                
                roomAnomalies.forEach(anom => {
                    const catName = anom.category || 'Général';
                    const catNorm = catName.toLowerCase();
                    const qNorm = (anom.question || '').toLowerCase();
                    
                    const isComment = catNorm.includes('remarque') || catNorm.includes('constat') || catNorm.includes('commentaire') ||
                                      qNorm.includes('remarque') || qNorm.includes('constat') || qNorm.includes('commentaire');
                    
                    if (!isComment) {
                        problemCats.set(catName, (problemCats.get(catName) || 0) + 1);
                    }
                });
            });
            
            poolAnomalies.set(session.fileName, sessionAnomalies);
        });

        const colors = ['#6366f1', '#22d3ee', '#f43f5e', '#fbbf24', '#34d399', '#a855f7', '#ec4899'];
        
        // 1. Taux de complétion
        const completionData = [
            { label: 'Complétés', value: completedRooms },
            { label: 'Incomplets', value: totalRooms - completedRooms }
        ];
        const completionPct = totalRooms === 0 ? 0 : Math.round((completedRooms / totalRooms) * 100);
        const htmlCompletion = createDonutChart('Taux de complétion', `${completionPct}%`, 'Global', completionData, ['#34d399', 'var(--surface-hover)']);
        
        // 2. Salles avec problèmes
        const problemRoomsData = [
            { label: 'Avec problèmes', value: roomsWithAnomalies },
            { label: 'Sans problème', value: totalRooms - roomsWithAnomalies }
        ];
        const problemPct = totalRooms === 0 ? 0 : Math.round((roomsWithAnomalies / totalRooms) * 100);
        const htmlProblemRooms = createDonutChart('Auditoires avec problèmes', `${problemPct}%`, `${roomsWithAnomalies}/${totalRooms}`, problemRoomsData, ['#f43f5e', '#22d3ee']);
        
        // 3. Top problèmes (hors commentaires)
        const sortedProbs = Array.from(problemCats.entries()).sort((a, b) => b[1] - a[1]);
        const topProbs = sortedProbs.slice(0, 4);
        const otherProbs = sortedProbs.slice(4).reduce((sum, item) => sum + item[1], 0);
        const topProbsData = topProbs.map(p => ({ label: p[0], value: p[1] }));
        if (otherProbs > 0) topProbsData.push({ label: 'Autres catégories', value: otherProbs });
        
        const topProbTotal = topProbsData.reduce((sum, i) => sum + i.value, 0);
        const htmlTopProblems = createDonutChart('Classement des problèmes', `${topProbTotal}`, 'Anomalies', topProbsData, colors);
        
        // 4. Problèmes par Pool
        const poolData = Array.from(poolAnomalies.entries())
            .filter(p => p[1] > 0)
            .sort((a, b) => b[1] - a[1])
            .map(p => ({ label: p[0].replace(/\.xlsx$/i, ''), value: p[1] }));
            
        const htmlPools = createDonutChart('Anomalies par Pool', `${totalAnomalies}`, 'Total', poolData, colors);

        statsChartsContainer.innerHTML = htmlCompletion + htmlProblemRooms + htmlTopProblems + htmlPools;
    }

    function renderIdentityCards() {
        const container = document.getElementById('identity-cards-container');
        const filterSelect = document.getElementById('identity-filter-file');
        if (!container) return;

        if (sessions.size === 0) {
            container.innerHTML = `
                <div class="empty-dashboard-state">
                    <span style="font-size: 3rem;">🏷️</span>
                    <h2>Aucun fichier actif</h2>
                    <p>Veuillez charger des relevés Excel pour générer les cartes d'identité.</p>
                </div>
            `;
            if (filterSelect) filterSelect.innerHTML = '<option value="all">Tous les fichiers</option>';
            return;
        }

        const filterVal = filterSelect ? filterSelect.value : 'all';
        if (filterSelect) {
            filterSelect.innerHTML = '<option value="all">Tous les fichiers (Pools)</option>';
        }

        const searchInput = document.getElementById('identity-search-input');
        const searchVal = searchInput ? searchInput.value.toLowerCase().trim() : '';

        let html = '';
        
        sessions.forEach(session => {
            if (filterSelect) {
                const opt = document.createElement('option');
                opt.value = session.id;
                opt.textContent = session.fileName;
                filterSelect.appendChild(opt);
            }

            if (filterVal !== 'all' && session.id !== filterVal) return;

            let roomCount = 0;
            let poolHtml = '';

            session.dataRows.forEach(dataRow => {
                const roomName = dataRow.name;
                if (searchVal && !roomName.toLowerCase().includes(searchVal)) return;

                if (!session.worksheet) return;
                const wsRow = session.worksheet.getRow(dataRow.rowIndex);
                roomCount++;
                
                let statsHtml = '';
                let capAnnoncee = '-';
                let batiment = '-';
                let fmg = '-';

                // Find announced capacity first for anomaly checks
                const capField = session.schema.find(f => {
                    const fc = (f.category || '').toLowerCase();
                    const fq = (f.question || '').toLowerCase();
                    return fc.includes('capacité annoncée') || fq.includes('capacité annoncée');
                });
                const announcedCap = capField ? getVal(wsRow, capField.colIndex) : '';

                session.schema.forEach(field => {
                    const val = getVal(wsRow, field.colIndex);
                    if (val === null || val === undefined || val.toString().trim() === '') return;
                    
                    const qNorm = (field.question || '').toLowerCase();
                    const catNorm = (field.category || '').toLowerCase();
                    const isAno = isValAnomaly(field, val, announcedCap);
                    
                    // Extract core info
                    if (catNorm.includes('capacité annoncée') || qNorm.includes('capacité annoncée')) {
                        capAnnoncee = val;
                    } else if (catNorm.includes('bâtiment') || qNorm.includes('bâtiment') || qNorm.includes('batiment')) {
                        batiment = val;
                    } else if (field.type === CONFIG.TYPES.GMF || qNorm.includes('gradin') || qNorm.includes('fixe') || qNorm.includes('mobile')) {
                        fmg = val;
                    }
                    
                    if (!isAno && field.type !== CONFIG.TYPES.YES_NO && field.type !== CONFIG.TYPES.TRUE_FALSE) {
                        if (!catNorm.includes('capacité annoncée') && !qNorm.includes('capacité annoncée') &&
                            !catNorm.includes('bâtiment') && !qNorm.includes('bâtiment') && !qNorm.includes('batiment') &&
                            field.type !== CONFIG.TYPES.GMF && !qNorm.includes('gradin') && !qNorm.includes('fixe') && !qNorm.includes('mobile') &&
                            !CONFIG.KEYWORDS.IDENTITY_COL.some(k => catNorm.includes(k))) {
                            
                            statsHtml += `
                                <div class="identity-stat-item">
                                    <span class="identity-stat-label">${field.question || field.category}</span>
                                    <span class="identity-stat-value">${val}</span>
                                </div>
                            `;
                        }
                    }
                });

                if (!statsHtml) {
                    statsHtml = '<div style="font-size:0.8rem; color:var(--text-muted); font-style:italic; padding: 0.5rem 0;">Aucune donnée saisie</div>';
                }

                poolHtml += `
                    <div class="identity-card">
                        <div class="identity-card-header">
                            <h4 class="identity-card-title">${roomName}</h4>
                            <div class="identity-card-badges">
                                <span class="identity-badge badge-building">🏢 ${batiment}</span>
                                <span class="identity-badge badge-capacity">👥 ${capAnnoncee} places</span>
                                <span class="identity-badge badge-type">🪑 ${fmg}</span>
                            </div>
                        </div>
                        <div class="identity-card-body">
                            <div class="identity-stat-list">
                                ${statsHtml}
                            </div>
                        </div>
                    </div>
                `;
            });

            if (roomCount > 0) {
                const poolName = session.fileName.replace(/\.xlsx$/i, '');
                html += `
                    <div class="identity-pool-group">
                        <h3 class="identity-pool-header">📂 Pool : ${poolName} <span class="pool-count">${roomCount} auditoires</span></h3>
                        <div class="identity-cards-grid">
                            ${poolHtml}
                        </div>
                    </div>
                `;
            }
        });

        if (!html) {
            html = `
                <div class="empty-dashboard-state">
                    <span style="font-size: 2rem;">🏷️</span>
                    <p>Aucun auditoire ne correspond à votre recherche.</p>
                </div>
            `;
        }

        container.innerHTML = html;
        
        if (filterSelect && !filterSelect.dataset.listenerAttached) {
            filterSelect.dataset.listenerAttached = 'true';
            filterSelect.addEventListener('change', renderIdentityCards);
        }
    }

    function isValAnomaly(field, val, announcedCapacity) {
        if (val === null || val === undefined) return false;
        const valStr = val.toString().trim();
        if (valStr === '') return false;

        const qNorm = (field.question || '').toLowerCase();
        const catNorm = (field.category || '').toLowerCase();
        const valNorm = valStr.toLowerCase();

        // 1. o/n (Oui/Non) fields
        if (field.type === CONFIG.TYPES.YES_NO) {
            const isNegative = CONFIG.AUTO_FILL.NEGATIVE_KEYWORDS.some(k => qNorm.includes(k) || catNorm.includes(k));
            if (isNegative) {
                return valNorm === 'o' || valNorm === 'oui' || valNorm === 'v' || valNorm === 'vrai';
            } else {
                return valNorm === 'n' || valNorm === 'non' || valNorm === 'f' || valNorm === 'faux';
            }
        }

        // 2. v/f (Vrai/Faux) fields
        if (field.type === CONFIG.TYPES.TRUE_FALSE) {
            const isNegative = CONFIG.AUTO_FILL.NEGATIVE_KEYWORDS.some(k => qNorm.includes(k) || catNorm.includes(k));
            if (isNegative) {
                return valNorm === 'v' || valNorm === 'vrai' || valNorm === 'o' || valNorm === 'oui';
            } else {
                return valNorm === 'f' || valNorm === 'faux' || valNorm === 'n' || valNorm === 'non';
            }
        }

        // 3. Number fields
        if (field.type === CONFIG.TYPES.NUMBER) {
            const isCapaciteReelle = CONFIG.AUTO_FILL.CAPACITY_TARGET.some(t => qNorm.includes(t)) || catNorm.includes('capacité réelle');
            if (isCapaciteReelle && announcedCapacity !== undefined && announcedCapacity !== null && announcedCapacity !== '') {
                const real = Number(val);
                const ann = Number(announcedCapacity);
                if (!isNaN(real) && !isNaN(ann) && real < ann) {
                    return true;
                }
            }
            // Other numeric anomaly
            const isNegative = CONFIG.AUTO_FILL.NEGATIVE_KEYWORDS.some(k => qNorm.includes(k) || catNorm.includes(k));
            if (isNegative) {
                const num = Number(val);
                if (!isNaN(num) && num > 0) {
                    return true;
                }
            }
        }

        // 4. Comment fields
        const isComment = catNorm.includes('remarque') || catNorm.includes('constat') || catNorm.includes('commentaire') ||
                          qNorm.includes('remarque') || qNorm.includes('constat') || qNorm.includes('commentaire');
        if (isComment && valStr.length > 0) {
            return true;
        }

        return false;
    }

    function getRowAnomalies(session, rowIndex) {
        if (!session || !session.worksheet) return [];
        const row = session.worksheet.getRow(rowIndex);
        const anomaliesList = [];

        let announcedCap = '';
        const capField = session.schema.find(f => {
            const fc = (f.category || '').toLowerCase();
            const fq = (f.question || '').toLowerCase();
            return fc === CONFIG.AUTO_FILL.CAPACITY_SOURCE || fq === CONFIG.AUTO_FILL.CAPACITY_SOURCE;
        });
        if (capField) {
            announcedCap = getVal(row, capField.colIndex);
        }

        session.schema.forEach(field => {
            if (!field.question) return;

            const cell = row.getCell(field.colIndex);
            let isExempt = false;
            if (cell && cell.style && cell.style.fill) {
                const f = cell.style.fill;
                if (f.type === 'pattern' && f.pattern && f.pattern !== 'none' && f.pattern !== 'solid') {
                    isExempt = true;
                }
            }

            if (isExempt) return;

            const val = getVal(row, field.colIndex);
            if (isValAnomaly(field, val, announcedCap)) {
                const qNorm = (field.question || '').toLowerCase();
                const catNorm = (field.category || '').toLowerCase();
                const isComment = catNorm.includes('remarque') || catNorm.includes('constat') || catNorm.includes('commentaire') ||
                                  qNorm.includes('remarque') || qNorm.includes('constat') || qNorm.includes('commentaire');
                
                anomaliesList.push({
                    colIndex: field.colIndex,
                    category: field.category,
                    question: field.question,
                    value: val,
                    type: field.type,
                    isComment: isComment
                });
            }
        });

        return anomaliesList;
    }

    function renderPostReleveDashboard() {
        if (sessions.size === 0) {
            postReleveSection.innerHTML = `
                <div class="empty-dashboard-state">
                    <span style="font-size: 3rem;">📊</span>
                    <h2>Aucun fichier actif</h2>
                    <p>Veuillez charger des relevés Excel pour voir les analyses.</p>
                </div>
            `;
            return;
        }

        let totalFiles = sessions.size;
        let totalRooms = 0;
        let completedRoomsCount = 0;
        let totalAnomaliesCount = 0;

        const fileBreakdown = [];
        const categoryStatsMap = new Map();
        const allAnomalies = [];

        const filterSelectVal = anomaliesFilterFile.value;
        anomaliesFilterFile.innerHTML = '<option value="all">Tous les fichiers</option>';

        sessions.forEach(session => {
            const opt = document.createElement('option');
            opt.value = session.id;
            opt.textContent = session.fileName;
            anomaliesFilterFile.appendChild(opt);

            let sessionRoomsCount = session.dataRows.length;
            let sessionCompletedCount = 0;
            let sessionAnomaliesCount = 0;

            totalRooms += sessionRoomsCount;

            session.dataRows.forEach(item => {
                const nameNorm = item.name.toLowerCase();
                const threshold = nameNorm.includes('hall, sanitaire') ? 40 : 60;
                const completion = getRowCompletion(session, item.rowIndex).percentage;

                if (completion >= threshold) {
                    completedRoomsCount++;
                    sessionCompletedCount++;
                }

                const roomAnomalies = getRowAnomalies(session, item.rowIndex);
                sessionAnomaliesCount += roomAnomalies.length;
                totalAnomaliesCount += roomAnomalies.length;

                roomAnomalies.forEach(anom => {
                    allAnomalies.push({
                        sessionId: session.id,
                        fileName: session.fileName,
                        rowIndex: item.rowIndex,
                        roomName: item.name,
                        ...anom
                    });

                    const catName = anom.category || 'Général';
                    const statKey = `${catName}::${anom.question}`;
                    if (!categoryStatsMap.has(statKey)) {
                        categoryStatsMap.set(statKey, {
                            category: catName,
                            question: anom.question,
                            count: 0
                        });
                    }
                    categoryStatsMap.get(statKey).count++;
                });
            });

            const sessionPct = sessionRoomsCount === 0 ? 0 : Math.round((sessionCompletedCount / sessionRoomsCount) * 100);
            fileBreakdown.push({
                id: session.id,
                name: session.fileName,
                total: sessionRoomsCount,
                completed: sessionCompletedCount,
                percentage: sessionPct,
                anomalies: sessionAnomaliesCount
            });
        });

        if (Array.from(anomaliesFilterFile.options).some(o => o.value === filterSelectVal)) {
            anomaliesFilterFile.value = filterSelectVal;
        } else {
            anomaliesFilterFile.value = 'all';
        }

        const globalCompletionPct = totalRooms === 0 ? 0 : Math.round((completedRoomsCount / totalRooms) * 100);
        dashboardMetricsRow.innerHTML = `
            <div class="metric-card purple">
                <div class="metric-icon">📂</div>
                <div class="metric-info">
                    <span class="metric-label">Fichiers Actifs</span>
                    <span class="metric-value">${totalFiles}</span>
                </div>
            </div>
            <div class="metric-card cyan">
                <div class="metric-icon">🎓</div>
                <div class="metric-info">
                    <span class="metric-label">Auditoires Total</span>
                    <span class="metric-value">${totalRooms}</span>
                </div>
            </div>
            <div class="metric-card green">
                <div class="metric-icon">📈</div>
                <div class="metric-info">
                    <span class="metric-label">Salles Complétées</span>
                    <span class="metric-value">${completedRoomsCount} / ${totalRooms} (${globalCompletionPct}%)</span>
                </div>
            </div>
            <div class="metric-card pink">
                <div class="metric-icon">🚨</div>
                <div class="metric-info">
                    <span class="metric-label">Anomalies</span>
                    <span class="metric-value">${totalAnomaliesCount}</span>
                </div>
            </div>
        `;

        dashboardFilesList.innerHTML = '';
        fileBreakdown.forEach(file => {
            const card = document.createElement('div');
            card.className = 'excel-card';
            card.innerHTML = `
                <div class="excel-card-header">
                    <span class="excel-card-title" title="${file.name}">${file.name.replace(/\.xlsx$/i, '')}</span>
                    <span class="excel-card-stats">${file.completed}/${file.total} salles</span>
                </div>
                <div class="excel-card-progress">
                    <div class="excel-card-track">
                        <div class="excel-card-fill" style="width: ${file.percentage}%;"></div>
                    </div>
                    <span class="excel-card-pct">${file.percentage}%</span>
                </div>
                <div style="display:flex; justify-content:space-between; align-items:center; margin-top:0.25rem;">
                    <span style="font-size:0.7rem; color:var(--text-muted);">⚠️ ${file.anomalies} anomalie(s)</span>
                    <button class="anomaly-goto-btn" type="button" style="padding:0.2rem 0.5rem; font-size:0.65rem;">
                        Ouvrir
                    </button>
                </div>
            `;

            card.querySelector('.anomaly-goto-btn').addEventListener('click', () => {
                switchToSession(file.id);
                switchMode('releve');
            });

            dashboardFilesList.appendChild(card);
        });

        dashboardStatsList.innerHTML = '';
        const statsListSorted = Array.from(categoryStatsMap.values())
            .filter(item => item.count > 0)
            .sort((a, b) => b.count - a.count);

        if (statsListSorted.length === 0) {
            dashboardStatsList.innerHTML = '<div style="font-size:0.75rem; color:var(--text-muted); text-align:center; padding:1rem 0;">Aucune anomalie détectée !</div>';
        } else {
            statsListSorted.forEach(stat => {
                const row = document.createElement('div');
                row.className = 'stat-row';
                row.style.cursor = 'pointer';
                row.title = 'Filtrer sur cette question';
                row.innerHTML = `
                    <span class="stat-name">
                        <span class="anomaly-cat-icon">${getCategoryIcon(stat.category)}</span>
                        <strong style="color:var(--text-primary); margin-right:3px;">${stat.category}</strong>
                        <span>- ${stat.question}</span>
                    </span>
                    <span class="stat-count">${stat.count}</span>
                `;

                row.addEventListener('click', () => {
                    anomaliesSearchInput.value = stat.question;
                    renderAnomaliesFeed(allAnomalies);
                });

                dashboardStatsList.appendChild(row);
            });
        }

        renderAnomaliesFeed(allAnomalies);
        renderExportCheckboxes();
        populateRemarksDropdown();
    }

    function renderAnomaliesFeed(allAnomalies) {
        anomaliesFeed.innerHTML = '';
        const query = anomaliesSearchInput.value.toLowerCase().trim();
        const fileFilter = anomaliesFilterFile.value;

        const filtered = allAnomalies.filter(anom => {
            if (fileFilter !== 'all' && anom.sessionId !== fileFilter) return false;
            
            if (query !== '') {
                const searchSpace = `${anom.roomName} ${anom.category} ${anom.question} ${anom.value}`.toLowerCase();
                if (!searchSpace.includes(query)) return false;
            }
            return true;
        });

        if (filtered.length === 0) {
            anomaliesFeed.innerHTML = '<div class="empty-state" style="margin-top:15%">Aucune anomalie trouvée.</div>';
            return;
        }

        filtered.sort((a, b) => {
            if (a.isComment && !b.isComment) return 1;
            if (!a.isComment && b.isComment) return -1;
            return 0;
        });

        filtered.forEach(anom => {
            const card = document.createElement('div');
            card.className = 'anomaly-card';
            if (anom.isComment) {
                card.style.borderColor = 'rgba(251, 191, 36, 0.15)';
                card.style.background = 'rgba(251, 191, 36, 0.01)';
            }

            const cleanFileName = anom.fileName.replace(/\.xlsx$/i, '');

            card.innerHTML = `
                <div class="anomaly-card-top">
                    <div class="anomaly-location">
                        <span class="anomaly-building">${cleanFileName}</span>
                        <span class="anomaly-room">${anom.roomName}</span>
                    </div>
                    <button class="anomaly-goto-btn" type="button">
                        ✏️ Corriger
                    </button>
                </div>
                <div class="anomaly-details" style="${anom.isComment ? 'border-color: rgba(251, 191, 36, 0.1);' : ''}">
                    <span class="anomaly-cat-icon">${getCategoryIcon(anom.category)}</span>
                    <span class="anomaly-q"><strong>${anom.category}</strong> - ${anom.question}</span>
                    <span class="anomaly-bad-val" style="${anom.isComment ? 'color: var(--warning-color); background: rgba(251, 191, 36, 0.12);' : ''}">
                        ${anom.value}
                    </span>
                </div>
            `;

            card.querySelector('.anomaly-goto-btn').addEventListener('click', () => {
                switchToSession(anom.sessionId);
                const session = sessions.get(anom.sessionId);
                if (session) {
                    const idx = session.dataRows.findIndex(r => r.rowIndex === anom.rowIndex);
                    if (idx !== -1) {
                        selectAuditoire(idx);
                        switchMode('releve');

                        setTimeout(() => {
                            const fieldName = `field-${anom.colIndex}`;
                            const input = formContainer.querySelector(`input[name="${fieldName}"], textarea[name="${fieldName}"]`);
                            if (input) {
                                input.scrollIntoView({ behavior: 'smooth', block: 'center' });
                                if (input.type === 'radio') {
                                    const radios = formContainer.querySelectorAll(`input[name="${fieldName}"]`);
                                    const checked = Array.from(radios).find(r => r.value === anom.value.toString().toLowerCase());
                                    if (checked) checked.focus();
                                    else if (radios.length > 0) radios[0].focus();
                                } else {
                                    input.focus();
                                }
                                const group = input.closest('.field-group');
                                if (group) {
                                    group.style.borderColor = 'var(--primary-color)';
                                    group.style.boxShadow = '0 0 12px var(--primary-glow)';
                                    setTimeout(() => {
                                        group.style.borderColor = '';
                                        group.style.boxShadow = '';
                                    }, 2000);
                                }
                            }
                        }, 300);
                    }
                }
            });

            anomaliesFeed.appendChild(card);
        });
    }

    function renderExportCheckboxes() {
        const prevCheckedFiles = new Set(Array.from(exportFilesCheckboxes.querySelectorAll('input:checked')).map(i => i.value));
        exportFilesCheckboxes.innerHTML = '';
        
        sessions.forEach(session => {
            const item = document.createElement('label');
            item.className = 'export-checkbox-item';
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = session.id;
            checkbox.checked = prevCheckedFiles.size === 0 || prevCheckedFiles.has(session.id);
            
            const span = document.createElement('span');
            span.textContent = session.fileName;
            
            item.appendChild(checkbox);
            item.appendChild(span);
            exportFilesCheckboxes.appendChild(item);
        });

        const categories = new Set();
        sessions.forEach(session => {
            session.schema.forEach(field => {
                if (field.category) categories.add(field.category);
            });
        });

        const prevCheckedCats = new Set(Array.from(exportCategoriesCheckboxes.querySelectorAll('input:checked')).map(i => i.value));
        exportCategoriesCheckboxes.innerHTML = '';

        Array.from(categories).sort().forEach(cat => {
            const item = document.createElement('label');
            item.className = 'export-checkbox-item';
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = cat;
            checkbox.checked = prevCheckedCats.size === 0 || prevCheckedCats.has(cat);
            
            const span = document.createElement('span');
            span.textContent = `${getCategoryIcon(cat)} ${cat}`;
            
            item.appendChild(checkbox);
            item.appendChild(span);
            exportCategoriesCheckboxes.appendChild(item);
        });
    }

    function populateRemarksDropdown() {
        const select = document.getElementById('remark-auditoire-select');
        const genInput = document.getElementById('remarque-generale-input');
        const debInput = document.getElementById('debriefing-input');
        if (!select || !genInput || !debInput) return;

        select.innerHTML = '<option value="">-- Choisir un auditoire --</option>';
        
        sessions.forEach(session => {
            const optgroup = document.createElement('optgroup');
            optgroup.label = session.fileName.replace(/\.xlsx$/i, '');
            
            session.dataRows.forEach(row => {
                const opt = document.createElement('option');
                opt.value = `${session.id}::${row.name}`;
                opt.textContent = row.name;
                optgroup.appendChild(opt);
            });
            
            select.appendChild(optgroup);
        });

        // Reset inputs
        genInput.value = '';
        debInput.value = '';
    }

    const remarkSelect = document.getElementById('remark-auditoire-select');
    const remarkGenInput = document.getElementById('remarque-generale-input');
    const remarkDebInput = document.getElementById('debriefing-input');
    const saveRemarkBtn = document.getElementById('save-remark-btn');

    if (remarkSelect && remarkGenInput && remarkDebInput && saveRemarkBtn) {
        remarkSelect.addEventListener('change', async () => {
            const val = remarkSelect.value;
            if (!val) {
                remarkGenInput.value = '';
                remarkDebInput.value = '';
                return;
            }
            const [sessionId, roomName] = val.split('::');
            const remarkData = await loadRemarkFromDB(sessionId, roomName);
            if (remarkData) {
                remarkGenInput.value = remarkData.remarqueGenerale || '';
                remarkDebInput.value = remarkData.debriefing || '';
            } else {
                remarkGenInput.value = '';
                remarkDebInput.value = '';
            }
        });

        saveRemarkBtn.addEventListener('click', () => {
            const val = remarkSelect.value;
            if (!val) {
                showAlert('Veuillez sélectionner un auditoire.', '⚠️', 'Attention');
                return;
            }
            const [sessionId, roomName] = val.split('::');
            const genText = remarkGenInput.value.trim();
            const debText = remarkDebInput.value.trim();
            
            saveRemarkToDB(sessionId, roomName, genText, debText);
            
            const originalText = saveRemarkBtn.innerHTML;
            saveRemarkBtn.innerHTML = '✅ Enregistré !';
            setTimeout(() => {
                saveRemarkBtn.innerHTML = originalText;
                renderSidebar(); // Update the bubble indicator in sidebar
            }, 2000);
        });
    }



    async function exportAnomaliesReport() {
        const checkedFileIds = Array.from(exportFilesCheckboxes.querySelectorAll('input:checked')).map(i => i.value);
        const checkedCategories = new Set(Array.from(exportCategoriesCheckboxes.querySelectorAll('input:checked')).map(i => i.value));
        const poolType = document.querySelector('input[name="export-pool"]:checked').value;

        if (checkedFileIds.length === 0) {
            showAlert('Veuillez sélectionner au moins un fichier à inclure.', '⚠️', 'Export impossible');
            return;
        }
        if (checkedCategories.size === 0) {
            showAlert('Veuillez sélectionner au moins une catégorie à inclure.', '⚠️', 'Export impossible');
            return;
        }

        try {
            const outWb = new ExcelJS.Workbook();
            const outWs = outWb.addWorksheet('Anomalies');

            outWs.columns = [
                { header: 'Bâtiment / Fichier', key: 'file', width: 22 },
                { header: 'Auditoire', key: 'room', width: 20 },
                { header: 'Catégorie', key: 'category', width: 20 },
                { header: 'Question', key: 'question', width: 35 },
                { header: 'Valeur Relevée', key: 'value', width: 16 },
                { header: 'Type de constat', key: 'type', width: 16 },
                { header: 'Remarque générale', key: 'remarque', width: 30 },
                { header: 'Débriefing', key: 'debriefing', width: 30 }
            ];

            const headerRow = outWs.getRow(1);
            headerRow.height = 24;
            headerRow.eachCell(cell => {
                cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, name: 'Segoe UI', size: 11 };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF1E2234' }
                };
                cell.alignment = { vertical: 'middle', horizontal: 'left' };
                cell.border = {
                    bottom: { style: 'medium', color: { argb: 'FF6366F1' } }
                };
            });

            let count = 0;
            for (const sessId of checkedFileIds) {
                const session = sessions.get(sessId);
                if (!session) continue;

                const remarksArray = await loadAllRemarksForSession(sessId);
                const remarksMap = new Map();
                remarksArray.forEach(r => remarksMap.set(r.roomName, r));

                session.dataRows.forEach(item => {
                    const completion = getRowCompletion(session, item.rowIndex).percentage;
                    const roomAnomalies = getRowAnomalies(session, item.rowIndex);

                    if (poolType === 'anomalies' && roomAnomalies.length === 0) return;
                    
                    const nameNorm = item.name.toLowerCase();
                    const threshold = nameNorm.includes('hall, sanitaire') ? 40 : 60;
                    if (poolType === 'completed' && completion < threshold) return;

                    const roomRemark = remarksMap.get(item.name) || {};

                    roomAnomalies.forEach(anom => {
                        if (!checkedCategories.has(anom.category)) return;

                        const row = outWs.addRow({
                            file: session.fileName.replace(/\.xlsx$/i, ''),
                            room: item.name,
                            category: anom.category,
                            question: anom.question,
                            value: anom.value,
                            type: anom.isComment ? 'Remarque / Note' : 'Anomalie',
                            remarque: roomRemark.remarqueGenerale || '',
                            debriefing: roomRemark.debriefing || ''
                        });

                        row.eachCell((cell, colNum) => {
                            cell.font = { name: 'Segoe UI', size: 10 };
                            if (colNum === 5 || colNum === 6) {
                                if (anom.isComment) {
                                    cell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FFD97706' }, bold: true };
                                } else {
                                    cell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FFEF4444' }, bold: true };
                                }
                            }
                            cell.border = {
                                bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } }
                            };
                            cell.alignment = { vertical: 'middle' };
                        });
                        count++;
                    });
                });
            }

            if (count === 0) {
                showAlert('Aucune anomalie ne correspond aux filtres d\'export sélectionnés.', 'ℹ️', 'Export vide');
                return;
            }

            const buffer = await outWb.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Rapport_Anomalies_Consolide.xlsx`;
            a.click();
            window.URL.revokeObjectURL(url);

            showAlert(`Rapport d'anomalies consolidé généré avec succès (${count} anomalies exportées).`, '✅', 'Export terminé');

        } catch (error) {
            console.error('Error generating anomalies report:', error);
            showAlert('Une erreur est survenue lors de la génération du rapport.', '❌', 'Erreur export');
        }
    }

    async function exportFilteredWorkbooks() {
        const checkedFileIds = Array.from(exportFilesCheckboxes.querySelectorAll('input:checked')).map(i => i.value);
        const checkedCategories = new Set(Array.from(exportCategoriesCheckboxes.querySelectorAll('input:checked')).map(i => i.value));
        const poolType = document.querySelector('input[name="export-pool"]:checked').value;

        if (checkedFileIds.length === 0) {
            showAlert('Veuillez sélectionner au moins un fichier à inclure.', '⚠️', 'Export impossible');
            return;
        }

        try {
            for (const sessId of checkedFileIds) {
                const session = sessions.get(sessId);
                if (!session) continue;

                const originalBuffer = await session.workbook.xlsx.writeBuffer();
                const cloneWb = new ExcelJS.Workbook();
                await cloneWb.xlsx.load(originalBuffer);
                const cloneWs = cloneWb.worksheets[0];

                const rowsToDelete = [];
                session.dataRows.forEach(item => {
                    const completion = getRowCompletion(session, item.rowIndex).percentage;
                    const roomAnomalies = getRowAnomalies(session, item.rowIndex);

                    let keep = true;
                    if (poolType === 'anomalies' && roomAnomalies.length === 0) keep = false;
                    
                    const nameNorm = item.name.toLowerCase();
                    const threshold = nameNorm.includes('hall, sanitaire') ? 40 : 60;
                    if (poolType === 'completed' && completion < threshold) keep = false;

                    if (!keep) {
                        rowsToDelete.push(item.rowIndex);
                    }
                });

                rowsToDelete.sort((a, b) => b - a);
                rowsToDelete.forEach(idx => {
                    cloneWs.spliceRows(idx, 1);
                });

                session.schema.forEach(field => {
                    const catNorm = (field.category || '').toLowerCase().trim();
                    const questNorm = (field.question || '').toLowerCase().trim();
                    const isBatiment = catNorm === 'bâtiments' || catNorm === 'batiments' || questNorm === 'bâtiments' || questNorm === 'batiments';
                    const isAuditoire = catNorm === 'auditoires' || questNorm === 'auditoires';
                    const isCapacite = catNorm.includes(CONFIG.AUTO_FILL.CAPACITY_SOURCE) || questNorm.includes(CONFIG.AUTO_FILL.CAPACITY_SOURCE);

                    if (isBatiment || isAuditoire || isCapacite) {
                        return;
                    }

                    if (!checkedCategories.has(field.category)) {
                        cloneWs.getColumn(field.colIndex).hidden = true;
                    }
                });

                const buffer = await cloneWb.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                const baseName = session.fileName.replace(/\.xlsx$/i, '');
                a.download = `${baseName}_Filtre.xlsx`;
                a.click();
                window.URL.revokeObjectURL(url);
            }

            showAlert('Les fichiers filtrés ont été générés et téléchargés avec succès.', '✅', 'Export terminé');

        } catch (error) {
            console.error('Error exporting filtered workbooks:', error);
            showAlert('Une erreur est survenue lors de l\'export des fichiers filtrés.', '❌', 'Erreur export');
        }
    }

    // Event Listeners for Switcher & Dashboard
    if (modeReleveBtn) modeReleveBtn.addEventListener('click', () => { switchMode('releve'); if (window.innerWidth <= 768) closeDrawer(); });
    if (modePostBtn) modePostBtn.addEventListener('click', () => { switchMode('post'); if (window.innerWidth <= 768) closeDrawer(); });
    if (modeStatsBtn) modeStatsBtn.addEventListener('click', () => { switchMode('stats'); if (window.innerWidth <= 768) closeDrawer(); });
    if (modeIdentityBtn) modeIdentityBtn.addEventListener('click', () => { switchMode('identity'); if (window.innerWidth <= 768) closeDrawer(); });

    /* ==========================================================================
       MOBILE LAYOUT: Move mode-switcher & file-tabs into sidebar drawer
       ========================================================================== */
    const mobileQuery = window.matchMedia('(max-width: 768px)');
    const appHeader = document.querySelector('.app-header');
    const headerControls = document.querySelector('.header-controls');
    const mainContent = document.querySelector('.main-content');

    function relocateMobileControls(isMobile) {
        if (!sidebar) return;
        const sidebarHeader = sidebar.querySelector('.sidebar-header');
        if (!sidebarHeader) return;

        if (isMobile) {
            if (modeSwitcher && modeSwitcher.parentElement !== sidebar) {
                sidebar.insertBefore(modeSwitcher, sidebarHeader);
            }
            if (fileTabsBar && fileTabsBar.parentElement !== sidebar) {
                sidebarHeader.before(fileTabsBar);
            }
        } else {
            if (modeSwitcher && modeSwitcher.parentElement === sidebar) {
                appHeader.insertBefore(modeSwitcher, headerControls);
            }
            if (fileTabsBar && fileTabsBar.parentElement === sidebar) {
                mainContent.insertBefore(fileTabsBar, splitView);
            }
        }
    }

    mobileQuery.addEventListener('change', (e) => relocateMobileControls(e.matches));
    relocateMobileControls(mobileQuery.matches);

    // Close drawer when clicking a file tab on mobile
    if (fileTabs) {
        fileTabs.addEventListener('click', (e) => {
            if (window.innerWidth <= 768 && e.target.closest('.file-tab')) {
                closeDrawer();
            }
        });
    }

    /* ==========================================================================
       MAGIC SEARCH BAR: Auto-focus search on typing when no field is active
       ========================================================================== */
    document.addEventListener('keydown', (e) => {
        // Only when an auditoire is selected and search is visible
        const session = S();
        if (!session || !session.currentRowIndex) return;
        if (!searchContainer || searchContainer.style.display === 'none') return;
        if (!searchInput) return;

        // Already focused on search
        if (document.activeElement === searchInput) return;

        // Don't interfere with form fields
        const active = document.activeElement;
        if (active && (
            active.tagName === 'INPUT' ||
            active.tagName === 'TEXTAREA' ||
            active.tagName === 'SELECT' ||
            active.isContentEditable
        )) return;

        // Only printable characters, no modifiers
        if (e.ctrlKey || e.metaKey || e.altKey) return;
        if (e.key.length !== 1) return;

        // Focus search and let the character be typed
        searchInput.focus();
    });
    
    if (anomaliesSearchInput) {
        anomaliesSearchInput.addEventListener('input', () => {
            const allAnoms = [];
            sessions.forEach(session => {
                session.dataRows.forEach(item => {
                    const roomAnoms = getRowAnomalies(session, item.rowIndex);
                    roomAnoms.forEach(anom => {
                        allAnoms.push({
                            sessionId: session.id,
                            fileName: session.fileName,
                            rowIndex: item.rowIndex,
                            roomName: item.name,
                            ...anom
                        });
                    });
                });
            });
            renderAnomaliesFeed(allAnoms);
        });
    }

    if (anomaliesFilterFile) {
        anomaliesFilterFile.addEventListener('change', () => {
            const allAnoms = [];
            sessions.forEach(session => {
                session.dataRows.forEach(item => {
                    const roomAnoms = getRowAnomalies(session, item.rowIndex);
                    roomAnoms.forEach(anom => {
                        allAnoms.push({
                            sessionId: session.id,
                            fileName: session.fileName,
                            rowIndex: item.rowIndex,
                            roomName: item.name,
                            ...anom
                        });
                    });
                });
            });
            renderAnomaliesFeed(allAnoms);
        });
    }

    if (generateSampleBtn) generateSampleBtn.addEventListener('click', () => generateVerificationSample());
    if (exportAnomaliesXlsxBtn) exportAnomaliesXlsxBtn.addEventListener('click', () => exportAnomaliesReport());
    if (exportFilteredXlsxBtn) exportFilteredXlsxBtn.addEventListener('click', () => exportFilteredWorkbooks());

    // Identity Search
    const identitySearchInput = document.getElementById('identity-search-input');
    if (identitySearchInput) {
        identitySearchInput.addEventListener('input', renderIdentityCards);
    }

    // Toggle All Checkboxes
    document.addEventListener('click', (e) => {
        if (e.target && e.target.classList.contains('toggle-all-btn')) {
            const targetId = e.target.getAttribute('data-target');
            if (targetId) {
                const container = document.getElementById(targetId);
                if (container) {
                    const checkboxes = container.querySelectorAll('input[type="checkbox"]');
                    const allChecked = Array.from(checkboxes).every(cb => cb.checked);
                    checkboxes.forEach(cb => cb.checked = !allChecked);
                }
            }
        }
    });

    initDB();

    // Register Service Worker for PWA with OTA Update Handling
    if ('serviceWorker' in navigator) {
        let refreshing = false;

        // When the new worker takes control, reload the page
        navigator.serviceWorker.addEventListener('controllerchange', () => {
            if (!refreshing) {
                refreshing = true;
                window.location.reload();
            }
        });

        navigator.serviceWorker.register('./sw.js').then(reg => {
            console.log('ServiceWorker registration successful:', reg.scope);

            // Function to ask user if they want to update
            const promptUserToUpdate = (worker) => {
                showModal({
                    icon: '🚀',
                    title: 'Mise à jour disponible !',
                    message: "Une nouvelle version de l'application est prête. Vos données et sauvegardes ne seront pas perdues. Voulez-vous mettre à jour maintenant ?",
                    buttons: [
                        { label: 'Plus tard', value: 'cancel' },
                        { label: 'Mettre à jour', value: 'update', primary: true }
                    ]
                }).then(choice => {
                    if (choice === 'update') {
                        // Tell the new worker to skip waiting
                        worker.postMessage({ action: 'skipWaiting' });
                    }
                });
            };

            // If a worker is already waiting to take over
            if (reg.waiting) {
                promptUserToUpdate(reg.waiting);
                return;
            }

            // If a worker is installing, wait for it to finish
            if (reg.installing) {
                const worker = reg.installing;
                worker.addEventListener('statechange', () => {
                    if (worker.state === 'installed' && navigator.serviceWorker.controller) {
                        promptUserToUpdate(worker);
                    }
                });
                return;
            }

            // Listen for new workers arriving
            reg.addEventListener('updatefound', () => {
                const worker = reg.installing;
                worker.addEventListener('statechange', () => {
                    if (worker.state === 'installed' && navigator.serviceWorker.controller) {
                        // A new worker is installed and waiting
                        promptUserToUpdate(worker);
                    }
                });
            });

        }).catch(err => {
            console.log('ServiceWorker registration failed:', err);
        });
    }

});
