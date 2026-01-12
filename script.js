document.addEventListener('DOMContentLoaded', () => {
    // Elements
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const uploadSection = document.getElementById('upload-section');
    const splitView = document.getElementById('split-view');
    const auditoireList = document.getElementById('auditoire-list');
    const formContainer = document.getElementById('form-container');
    const currentAuditoireTitle = document.getElementById('current-auditoire-title');
    // fileNameDisplay removed from HTML
    const exportXlsxBtn = document.getElementById('export-xlsx-btn');
    const exportJsonBtn = document.getElementById('export-json-btn');
    const clearBtn = document.getElementById('clear-btn');

    // Persistence Config
    const DB_NAME = 'TavlDB';
    const STORE_FILE = 'file';
    const STORE_EDITS = 'edits'; // We will store cell edits as { sheetName, row, col, val }

    // State
    let currentWorkbook = null;
    let mainWorksheet = null;
    let schema = []; // Array of { colIndex, category, question, type }
    let dataRows = []; // Array of { rowIndex, auditoireName, ... }
    let currentRowIndex = null;
    let db = null;

    // Init DB
    const initDB = () => {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, 2); // Bumpped version to ensure clean store if needed (auto-handle via upgradeneeded if structure changed or just overwrite)
            request.onupgradeneeded = (e) => {
                db = e.target.result;
                if (!db.objectStoreNames.contains(STORE_FILE)) db.createObjectStore(STORE_FILE);
                if (!db.objectStoreNames.contains(STORE_EDITS)) db.createObjectStore(STORE_EDITS, { keyPath: 'id' });
            };
            request.onsuccess = (e) => {
                db = e.target.result;
                resolve(db);
                checkSavedSession();
            };
            request.onerror = (e) => reject(e);
        });
    };

    // Helper: Safe Cell Value
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

    // Drag & Drop Handlers
    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0], true);
    });
    fileInput.addEventListener('change', (e) => { if (e.target.files.length > 0) handleFile(e.target.files[0], true); });

    // Persistence Logic
    async function saveFileToDB(fileBuffer, fileName) {
        if (!db) return;
        const tx = db.transaction([STORE_FILE], 'readwrite');
        // Store as a single object with a fixed key 'current'
        tx.objectStore(STORE_FILE).put({ buffer: fileBuffer, name: fileName }, 'current');
    }

    async function saveEditToDB(row, col, val) {
        if (!db) return;
        const tx = db.transaction([STORE_EDITS], 'readwrite');
        const id = `${row}-${col}`;
        tx.objectStore(STORE_EDITS).put({ id, row, col, val });
    }

    async function clearDB() {
        if (!db) return;
        const tx = db.transaction([STORE_FILE, STORE_EDITS], 'readwrite');
        tx.objectStore(STORE_FILE).clear();
        tx.objectStore(STORE_EDITS).clear();
    }

    async function checkSavedSession() {
        if (!db) return;
        const tx = db.transaction([STORE_FILE], 'readonly');
        const req = tx.objectStore(STORE_FILE).get('current');

        req.onsuccess = async () => {
            if (req.result && req.result.buffer) {
                // Found a saved file
                if (confirm('Une session précédente (' + req.result.name + ') a été trouvée. Voulez-vous la restaurer ?')) {
                    const blob = new Blob([req.result.buffer]);
                    const file = new File([blob], req.result.name);
                    await handleFile(file, false); // False = don't overwrite DB yet

                    // Apply Edits
                    const tx2 = db.transaction([STORE_EDITS], 'readonly');
                    const cursorReq = tx2.objectStore(STORE_EDITS).openCursor();
                    cursorReq.onsuccess = (e) => {
                        const cursor = e.target.result;
                        if (cursor) {
                            const { row, col, val } = cursor.value;
                            const r = mainWorksheet.getRow(row);
                            const c = r.getCell(col);
                            c.value = val; // Apply edit
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

    // File Handling
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

            // Switch View
            uploadSection.classList.add('hidden');
            splitView.classList.remove('hidden');
            // fileNameDisplay removed

        } catch (error) {
            console.error('Error parsing file:', error);
            alert('Erreur: Structure du fichier non reconnue ou fichier invalide.');
        }
    }

    // Parsing Logic (Matrix)
    function parseMatrixStructure(sheet) {
        schema = [];
        dataRows = [];

        // Row 3: Categories (e.g., "Capacité réelle", "Mobilier")
        // Row 4: Questions (e.g., "Nombre de places...", "Horloge opérationnelle")
        // Row 5: Types (e.g., "Nombre", "v/f", "o/n")

        // We start scanning columns from index 1.
        // We propagate "Category" forward if empty (merged cell behavior simulation)

        // Detect specific columns based on Row 3 headers
        let auditoireColIndex = 3; // Default to C

        let lastCategory = '';
        const row3 = sheet.getRow(3);
        const row4 = sheet.getRow(4);
        const row5 = sheet.getRow(5);

        // Scan columns involved (assume max 100 columns for safety or until empty block)
        const colCount = sheet.columnCount;

        for (let c = 1; c <= colCount; c++) {
            let cat = getVal(row3, c);
            if (cat) lastCategory = cat;

            // Check if this is the Identity Column
            if (cat && cat.toLowerCase().includes('auditoire')) {
                auditoireColIndex = c;
            }

            let question = getVal(row4, c);
            let typeRaw = getVal(row5, c);
            let type = typeRaw.toLowerCase().trim();

            // Normalization
            if (type.includes('..') || type.includes('date')) type = 'date';

            // Specific overrides based on Headers
            if ((cat && cat.includes('Gradin')) || (question && question.includes('Gradin'))) {
                type = 'gmf';
            }

            // We include columns that have a Question OR a Type
            if (question || type) {
                schema.push({
                    colIndex: c,
                    category: lastCategory,
                    question: question,
                    type: type || 'text' // Default to text
                });
            }
        }

        // Scan Data Rows (Row 6 onwards)
        // We assume Row 6 until end are Auditoires.
        // We identify the Auditoire Name usually in Column 1 or 2 (Batiments / Auditoires).
        // Based on user screenshot: Col B = Bâtiments, Col C = Auditoires.
        // We'll search for the "Auditoires" column in Row 3 headers if possible, or just guess Col 3.

        // Let's rely on finding "Auditoire" in row 3 header? Or just check first few cols.
        // Screenshot shows: Col C is "Auditoires".

        sheet.eachRow((row, rowIndex) => {
            if (rowIndex < 6) return;

            // Try to find a name
            const name = getVal(row, auditoireColIndex);
            if (name) {
                dataRows.push({
                    rowIndex: rowIndex,
                    name: name
                });
            }
        });
    }

    // Render Sidebar
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

    // Select Auditoire
    function selectAuditoire(index) {
        // UI Active State
        document.querySelectorAll('.sidebar-item').forEach(el => el.classList.remove('active'));
        auditoireList.children[index].classList.add('active');

        const item = dataRows[index];
        currentRowIndex = item.rowIndex;
        currentAuditoireTitle.textContent = item.name;

        renderForm(item.rowIndex);
    }

    // Render Form
    function renderForm(rowIndex) {
        formContainer.innerHTML = '';
        const row = mainWorksheet.getRow(rowIndex);

        // Group by Category
        const byCategory = {};
        schema.forEach(field => {
            if (!byCategory[field.category]) byCategory[field.category] = [];
            byCategory[field.category].push(field);
        });

        for (const [category, fields] of Object.entries(byCategory)) {
            // Skip if category is likely metadata columns (like Batiments, Auditoires) 
            // unless they are marked as editable types?
            // Actually, we should filter out the Identity columns if they don't have a "Type" in Row 5.
            // Screenshot shows Type "F" for Gradin/Mobile/Fixe in Row 5?
            // Let's assume everything in Schema is worth showing.

            const catDiv = document.createElement('div');
            catDiv.className = 'form-category';
            const h3 = document.createElement('h3');
            h3.textContent = category || 'Général';
            catDiv.appendChild(h3);

            let hasVisibleFields = false;

            fields.forEach(field => {
                // If question is empty and type is empty, skip
                if (!field.question && !field.type) return;

                hasVisibleFields = true;
                const cell = row.getCell(field.colIndex); // Get actual cell to check style
                const val = getVal(row, field.colIndex);

                // Detection logic for "Optional" (Pattern fill based on user input)
                const isFacultatif = (c) => {
                    if (!c || !c.style || !c.style.fill) return false;
                    const f = c.style.fill;
                    // User showed a pattern fill (dots/grid).
                    // ExcelJS maps patterns like 'gray125', 'gray0625', 'mediumGray', etc.
                    // We check if type is value pattern and it is NOT 'none' or 'solid' (unless solid grey, but user emphasized pattern)
                    if (f.type === 'pattern' && f.pattern && f.pattern !== 'none' && f.pattern !== 'solid') {
                        return true;
                    }
                    return false;
                };

                const isOptional = isFacultatif(cell);

                const group = document.createElement('div');
                group.className = 'field-group';

                const label = document.createElement('label');
                label.className = 'field-question';
                label.textContent = field.question || field.category; // Fallback

                if (isOptional) {
                    const badge = document.createElement('span');
                    badge.className = 'badge-optional';
                    badge.textContent = 'Facultatif';
                    label.appendChild(badge);
                }

                group.appendChild(label);

                // Read-Only Logic
                // Normalize strings to check against user list
                const catNorm = (field.category || '').toLowerCase().trim();
                const questNorm = (field.question || '').toLowerCase().trim();

                // List of keywords that trigger read-only
                const readOnlyKeywords = [
                    'bâtiment', 'batiment',
                    'auditoire',
                    'capacité annoncée',
                    'gradin', 'mobile', 'fixe', 'gradin/mobile/fixe'
                ];

                // Check if category or question contains these keywords strictly for the user context
                // User gave strict list: "Bâtiments", "Auditoires", "Capacité annoncée", "Gradin/Mobile/Fixe"
                // We'll try to match inclusive.
                let isReadOnly = false;
                if (
                    catNorm.includes('bâtiment') ||
                    catNorm.includes('auditoires') ||
                    catNorm.includes('capacité annoncée') ||
                    questNorm.includes('bâtiment') ||
                    questNorm.includes('auditoires') ||
                    questNorm.includes('capacité annoncée') ||
                    (catNorm.includes('gradin') && catNorm.includes('mobile')) // Specific for G/M/F likely
                ) {
                    isReadOnly = true;
                }

                // Special check for GMF if type is 'gmf' and context implies it's the structural definition
                if (field.type === 'gmf' && (catNorm.includes('gradin') || questNorm.includes('gradin'))) {
                    isReadOnly = true;
                }


                // Render Input based on type
                const type = field.type;

                if (type === 'v/f' || type === 'o/n' || type === 'gmf') {
                    // Radio Group
                    let options = [];
                    let labels = {};

                    if (type === 'v/f') { options = ['v', 'f']; labels = { v: 'Vrai', f: 'Faux' }; }
                    else if (type === 'o/n') { options = ['o', 'n']; labels = { o: 'Oui', n: 'Non' }; }
                    else if (type === 'gmf') { options = ['G', 'M', 'F']; labels = { G: 'Gradin', M: 'Mobile', F: 'Fixe' }; }

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

                        // Case insensitive compare
                        if (val.toString().toLowerCase() === opt.toLowerCase()) input.checked = true;

                        input.addEventListener('change', () => updateCell(field.colIndex, opt));

                        const textSpan = document.createElement('span');
                        textSpan.textContent = labels[opt] || opt;

                        wrapper.appendChild(input);
                        wrapper.appendChild(textSpan);
                        radioContainer.appendChild(wrapper);
                    });
                    group.appendChild(radioContainer);

                } else if (type === 'date') {
                    const input = document.createElement('input');
                    input.type = 'date';
                    if (isReadOnly) {
                        input.disabled = true;
                        input.classList.add('input-disabled');
                    }
                    // Excel dates might need conversion if numeric
                    if (val && !isNaN(Date.parse(val))) {
                        // It is a string date
                        // Try to format to YYYY-MM-DD for input type=date
                        const d = new Date(val);
                        if (!isNaN(d)) {
                            input.value = d.toISOString().split('T')[0];
                        } else {
                            input.value = val;
                        }
                    } else if (typeof val === 'number') {
                        // Excel serial date
                        // Excel base date is Dec 30 1899 usually
                        const d = new Date(Math.round((val - 25569) * 86400 * 1000));
                        if (!isNaN(d)) {
                            input.value = d.toISOString().split('T')[0];
                        }
                    } else {
                        input.value = val;
                    }

                    input.addEventListener('change', (e) => updateCell(field.colIndex, e.target.value));
                    group.appendChild(input);

                } else if (type === 'nombre') {
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
                    // Default Text
                    const input = document.createElement('textarea');
                    input.rows = 2; // Auto expand maybe?
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

    // Update Cell Logic
    function updateCell(colIndex, value) {
        if (!currentRowIndex) return;
        const row = mainWorksheet.getRow(currentRowIndex);
        const cell = row.getCell(colIndex);
        cell.value = value;
        // Auto-save logic is essentially done here as we modify the object reference
        // Save to DB
        saveEditToDB(currentRowIndex, colIndex, value);
    }

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

    // Clear / Reset
    clearBtn.addEventListener('click', () => {
        splitView.classList.add('hidden');
        uploadSection.classList.remove('hidden');
        currentWorkbook = null;
        fileInput.value = '';
    });

    // --- New Features Logic ---

    // 1. Force Edit Mode
    const unlockBtn = document.getElementById('unlock-btn');
    unlockBtn.addEventListener('click', () => {
        if (confirm("Voulez-vous activer le mode 'Édition Forcée' ? Cela déverrouillera tous les champs structurels (Bâtiments, Capacité, etc.).")) {
            document.body.classList.toggle('force-edit-mode');
            if (currentRowIndex) renderForm(currentRowIndex); // Re-render to apply removal of disabled state
        }
    });

    // 2. Fill Defaults
    const fillDefaultsBtn = document.getElementById('fill-defaults-btn');
    fillDefaultsBtn.addEventListener('click', () => {
        if (!currentRowIndex || !mainWorksheet) return;

        const row = mainWorksheet.getRow(currentRowIndex);
        let editsMade = 0;

        // Determine Capacité Annoncée Value (Search schema for it)
        let capaAnnonceVal = '';
        const capaField = schema.find(f => {
            const c = (f.category || '').toLowerCase();
            const q = (f.question || '').toLowerCase();
            return c === 'capacité annoncée' || q === 'capacité annoncée';
        });
        if (capaField) {
            capaAnnonceVal = getVal(row, capaField.colIndex);
        }

        const today = new Date().toISOString().split('T')[0];

        schema.forEach(field => {
            // Logic: Skip if Optional
            const cell = row.getCell(field.colIndex);
            // Use our helper from render but we need to externalize it or duplicate logic slightly
            // Checking style directly here
            let isOptional = false;
            if (cell && cell.style && cell.style.fill) {
                const f = cell.style.fill;
                if (f.type === 'pattern' && f.pattern && f.pattern !== 'none' && f.pattern !== 'solid') isOptional = true;
            }
            if (isOptional) return; // Do not fill optional fields

            // Skip if value already exists (don't overwrite user data? User said "Remplir par defaut", implies filling empty ones or all? "Remplir par défaut" usually means fill empty. I'll stick to filling ONLY if empty to be safe, except specifically requested logic might imply otherwise. User didn't specify "only empty", but it's safer.)
            // Actually, for "Capacité réelle meme que celle anoncee", if it's 0 or empty we fill it.
            const currentVal = getVal(row, field.colIndex);
            if (currentVal && currentVal.toString().trim() !== '') return;

            // Apply Rules
            const qNorm = (field.question || '').toLowerCase();

            // Capacité Réelle
            if ((field.category || '').toLowerCase().includes('capacité réelle') || qNorm.includes('réellement fonctionnelles')) {
                if (capaAnnonceVal) {
                    updateCell(field.colIndex, capaAnnonceVal);
                    editsMade++;
                }
            }
            // Date de passage
            else if (field.type === 'date' || qNorm.includes('date de passage')) {
                updateCell(field.colIndex, today);
                editsMade++;
            }
            // O/N -> O (Exception: Humidité)
            else if (field.type === 'o/n') {
                if (qNorm.includes('humidit') || qNorm.includes('infiltration')) {
                    updateCell(field.colIndex, 'n'); // Default Non
                } else {
                    updateCell(field.colIndex, 'o'); // Default Oui
                }
                editsMade++;
            }
            // V/F -> V
            else if (field.type === 'v/f') {
                updateCell(field.colIndex, 'v');
                editsMade++;
            }
        });

        if (editsMade > 0) {
            renderForm(currentRowIndex); // Refresh UI
            // alert(`${editsMade} champs remplis par défaut.`);
        }
    });

    // 3. Next Field Navigation
    const nextFieldBtn = document.getElementById('next-field-btn');

    // Show/Hide FAB based on view
    const observer = new MutationObserver(() => {
        if (!splitView.classList.contains('hidden')) {
            nextFieldBtn.classList.remove('hidden');
        } else {
            nextFieldBtn.classList.add('hidden');
        }
    });
    observer.observe(splitView, { attributes: true });

    // Track last focused input for Navigation Logic
    let lastFocusedInput = null;
    formContainer.addEventListener('focusin', (e) => {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
            lastFocusedInput = e.target;
        }
    });

    nextFieldBtn.addEventListener('click', () => {
        // Find all inputs in form
        const inputs = Array.from(formContainer.querySelectorAll('input:not(:disabled), textarea:not(:disabled)'));

        let startIdx = 0;
        const current = lastFocusedInput;
        const currentIdx = inputs.indexOf(current);

        if (currentIdx !== -1) {
            startIdx = currentIdx + 1;

            // Optimization: If current was a radio, skip all siblings of the same name immediately
            if (current.type === 'radio') {
                const name = current.name;
                while (startIdx < inputs.length && inputs[startIdx].name === name) {
                    startIdx++;
                }
            }
        }

        // Search from startIdx to end, then 0 to startIdx (Wrap around)
        let found = null;

        // Helper to check if input is valid target
        const isTarget = (input) => {
            // Skip optional fields
            const group = input.closest('.field-group');
            if (group) {
                const badge = group.querySelector('.badge-optional');
                if (badge) return false;
            }

            // Check if empty
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
                // If found radio, we should skip checking other radios of same group in this pass? 
                // No need, we break immediately. 
                break;
            }
            // If we didn't match (e.g. checked radio), we should probably skip the rest of this group
            // to save iterations, but not strictly necessary for correctness.
            // Actually, if inputs[i] is radio and NOT target (checked), subsequent siblings are also NOT target.
            if (inputs[i].type === 'radio') {
                const name = inputs[i].name;
                while (i + 1 < inputs.length && inputs[i + 1].name === name) {
                    i++;
                }
            }
        }

        // Wrap around Pass (if not found yet and we didn't start at 0)
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

    // Init Persistence
    initDB();
});
