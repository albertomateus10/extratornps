// Elementos do DOM
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileInfoBox = document.getElementById('file-info-box');
const fileNameEl = document.getElementById('file-name');
const fileSizeEl = document.getElementById('file-size');
const btnRemoveFile = document.getElementById('btn-remove-file');
const configSection = document.getElementById('config-section');
const mappingList = document.getElementById('mapping-list');
const btnAddRow = document.getElementById('btn-add-row');
const headerRowNumInput = document.getElementById('header-row-num');
const dataRowNumInput = document.getElementById('data-row-num');
const outputFilenameInput = document.getElementById('output-filename');
const btnExport = document.getElementById('btn-export');
const exportSpinner = document.getElementById('export-spinner');
const toastContainer = document.getElementById('toast-container');
const consultantsContainer = document.getElementById('consultants-container');

// Estado Global da Aplicação
let rawWorkbook = null;
let sheetData = []; // Array de arrays (linhas da planilha ativa)
let originalHeaders = []; // Array de strings/objetos representando as colunas detectadas

// Mapeamento padrão inicial conforme especificado pelo usuário
const defaultMappings = [
    { id: 'row-0', outputHeader: 'Loja', inputColIndex: 10 },
    { id: 'row-1', outputHeader: 'Nome do consultor', inputColIndex: 12 },
    { id: 'row-2', outputHeader: 'Chassi', inputColIndex: 13 },
    { id: 'row-3', outputHeader: 'Nome do cliente', inputColIndex: 17 },
    { id: 'row-4', outputHeader: 'Email', inputColIndex: 18 },
    { id: 'row-5', outputHeader: 'Celular', inputColIndex: 20 },
    { id: 'row-6', outputHeader: 'OBS Cris', inputColIndex: '' },
    { id: 'row-7', outputHeader: 'Data inicial do convite', inputColIndex: 28 },
    { id: 'row-8', outputHeader: 'Data do Primeiro Lembrete', inputColIndex: 29 },
    { id: 'row-9', outputHeader: 'Segunda data do lembrete', inputColIndex: 30 },
    { id: 'row-10', outputHeader: 'Data de expiração do convite', inputColIndex: 36 }
];

let columnMappings = [];

// Lista de consultores alvo mapeada por loja
const TARGET_CONSULTANTS_BY_STORE = {
    "MATRIZ": [
        "Alan",
        "Bruna",
        "Ícaro",
        "Márcia",
        "Eduarda"
    ],
    "ZONA SUL": [
        "Jean",
        "Maicon"
    ],
    "GRAVATAÍ": [
        "Letícia",
        "Tiago"
    ],
    "VIAMÃO": [
        "Márcio"
    ]
};

// Inicialização: copia o mapeamento padrão
resetMappings();

// Configura o nome do arquivo padrão com a data atual
setDefaultFilename();

// Renderiza os botões dos consultores na grade
renderConsultantButtons();

// Carregar Event Listeners
setupEventListeners();

function setDefaultFilename() {
    const today = new Date();
    const day = String(today.getDate()).padStart(2, '0');
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const year = today.getFullYear();
    outputFilenameInput.value = `Base NPS ${day}-${month}-${year}.xlsx`;
}

function setupEventListeners() {
    // Eventos de Drag & Drop
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
        }, false);
    });

    dropZone.addEventListener('drop', handleDrop, false);
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);

    btnRemoveFile.addEventListener('click', resetApp);
    btnAddRow.addEventListener('click', addNewMappingRow);

    // Recalcular colunas e prévia se mudar a linha do cabeçalho ou início dos dados
    headerRowNumInput.addEventListener('change', processHeadersAndPreview);
    dataRowNumInput.addEventListener('change', processHeadersAndPreview);

    btnExport.addEventListener('click', () => exportData());
}

function resetMappings() {
    columnMappings = defaultMappings.map((m, idx) => ({
        id: `row-${idx}-${Date.now()}`,
        outputHeader: m.outputHeader,
        inputColIndex: m.inputColIndex,
        outputLetter: getExcelColumnName(idx + 1)
    }));
}

// Manipular arquivo arrastado
function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    if (files.length > 0) {
        processUploadedFile(files[0]);
    }
}

// Manipular arquivo selecionado no explorador
function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        processUploadedFile(files[0]);
    }
}

// Mostrar notificações Toast
function showToast(message, type = 'success') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    let icon = '<i class="fa-solid fa-circle-check" style="color: var(--success);"></i>';
    if (type === 'error') {
        icon = '<i class="fa-solid fa-triangle-exclamation" style="color: var(--error);"></i>';
    } else if (type === 'info') {
        icon = '<i class="fa-solid fa-circle-info" style="color: var(--accent-secondary);"></i>';
    }

    toast.innerHTML = `${icon}<span>${message}</span>`;
    toastContainer.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(20px)';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// Formatar bytes de tamanho de arquivo
function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Processar o arquivo carregado
function processUploadedFile(file) {
    const extension = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(extension)) {
        showToast('Formato de arquivo inválido! Envie .xlsx, .xls ou .csv', 'error');
        return;
    }

    fileNameEl.textContent = file.name;
    fileSizeEl.textContent = formatBytes(file.size);
    dropZone.style.display = 'none';
    fileInfoBox.style.display = 'flex';

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            rawWorkbook = XLSX.read(data, { type: 'array' });
            
            const firstSheetName = rawWorkbook.SheetNames[0];
            const worksheet = rawWorkbook.Sheets[firstSheetName];
            
            sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });

            if (sheetData.length === 0) {
                throw new Error("A planilha está vazia.");
            }

            showToast('Planilha carregada com sucesso!', 'success');
            
            configSection.style.display = 'grid';
            btnExport.disabled = false;
            toggleConsultantButtons(false);

            processHeadersAndPreview();
        } catch (error) {
            console.error(error);
            showToast('Erro ao ler a planilha: ' + error.message, 'error');
            resetApp();
        }
    };

    reader.onerror = function() {
        showToast('Erro ao ler o arquivo.', 'error');
        resetApp();
    };

    reader.readAsArrayBuffer(file);
}

// Processar Cabeçalhos e Atualizar Dropdowns + Prévia
function processHeadersAndPreview() {
    if (!sheetData || sheetData.length === 0) return;

    const headerRowIdx = parseInt(headerRowNumInput.value, 10) - 1;
    const dataRowIdx = parseInt(dataRowNumInput.value, 10) - 1;

    if (headerRowIdx < 0 || headerRowIdx >= sheetData.length) {
        showToast('A linha do cabeçalho informada é inválida.', 'error');
        return;
    }

    const rawHeaders = sheetData[headerRowIdx] || [];
    originalHeaders = [];

    const maxCols = Math.max(rawHeaders.length, ...sheetData.slice(dataRowIdx, dataRowIdx + 10).map(r => r.length));

    for (let i = 0; i < maxCols; i++) {
        const colLetter = XLSX.utils.encode_col(i);
        const colName = rawHeaders[i] !== null && rawHeaders[i] !== undefined ? String(rawHeaders[i]).trim() : '';
        originalHeaders.push({
            index: i,
            letter: colLetter,
            name: colName || `[Coluna ${colLetter} Sem Nome]`
        });
    }

    // Renderizar as linhas de mapeamento com os dados atualizados
    renderMappingRows();
}

// Renderizar linhas de mapeamento dinamicamente com base no estado `columnMappings`
function renderMappingRows() {
    mappingList.innerHTML = '';

    columnMappings.forEach((mapping, index) => {
        const row = document.createElement('div');
        row.className = 'mapping-row';
        row.id = mapping.id;

        // Gerar options para as colunas de entrada
        let optionsHTML = '<option value="">-- Nenhuma (Em branco) --</option>';
        originalHeaders.forEach(col => {
            const isSelected = mapping.inputColIndex !== '' && parseInt(mapping.inputColIndex, 10) === col.index;
            optionsHTML += `<option value="${col.index}" ${isSelected ? 'selected' : ''}>${col.letter} - ${col.name}</option>`;
        });

        row.innerHTML = `
            <div class="mapping-label">
                <i class="fa-solid fa-table-columns"></i>
                <span>Coluna ${mapping.outputLetter}</span>
            </div>
            <div class="mapping-inputs">
                <input type="text" class="input-custom out-header" value="${mapping.outputHeader}" placeholder="Cabeçalho de saída (opcional)">
                <select class="select-custom in-col-select">
                    ${optionsHTML}
                </select>
            </div>
            <button class="btn-delete-row" title="Remover mapeamento" ${index === 0 ? 'style="visibility: hidden;"' : ''}>
                <i class="fa-solid fa-trash-can"></i>
            </button>
        `;

        mappingList.appendChild(row);

        // Listeners para a linha recém-adicionada
        const outHeaderInput = row.querySelector('.out-header');
        const selectEl = row.querySelector('.in-col-select');
        const btnDelete = row.querySelector('.btn-delete-row');

        outHeaderInput.addEventListener('input', (e) => {
            mapping.outputHeader = e.target.value;
        });

        selectEl.addEventListener('change', (e) => {
            mapping.inputColIndex = e.target.value !== '' ? parseInt(e.target.value, 10) : '';
            
            // Se o cabeçalho estiver vazio e for selecionada uma coluna válida, preenche automaticamente
            if (outHeaderInput.value === '' && mapping.inputColIndex !== '') {
                const orig = originalHeaders.find(col => col.index === mapping.inputColIndex);
                if (orig) {
                    outHeaderInput.value = orig.name;
                    mapping.outputHeader = orig.name;
                }
            }
        });

        if (index > 0) {
            btnDelete.addEventListener('click', () => {
                columnMappings = columnMappings.filter(m => m.id !== mapping.id);
                recalculateOutputLetters();
                renderMappingRows();
            });
        }
    });
}

// Adicionar uma nova linha de mapeamento vazia
function addNewMappingRow() {
    const nextIndex = columnMappings.length;
    const nextLetter = getExcelColumnName(nextIndex + 1);
    const id = `row-${Date.now()}`;

    columnMappings.push({
        id: id,
        outputHeader: '',
        inputColIndex: '',
        outputLetter: nextLetter
    });

    renderMappingRows();
    
    // Scroll para baixo na lista
    mappingList.scrollTop = mappingList.scrollHeight;
}

// Recalcular as letras de saída (A, B, C...)
function recalculateOutputLetters() {
    columnMappings.forEach((mapping, idx) => {
        mapping.outputLetter = getExcelColumnName(idx + 1);
    });
}

// Gerador de letra do Excel a partir do índice (1 = A, 2 = B, 27 = AA...)
function getExcelColumnName(columnNumber) {
    let columnName = '';
    while (columnNumber > 0) {
        let modulo = (columnNumber - 1) % 26;
        columnName = String.fromCharCode(65 + modulo) + columnName;
        columnNumber = parseInt((columnNumber - modulo) / 26);
    }
    return columnName;
}



// Exportar Dados
function exportData(filterConsultantName = null) {
    if (filterConsultantName instanceof Event) {
        filterConsultantName = null;
    }
    if (!sheetData || sheetData.length === 0) {
        showToast('Nenhum dado para exportar.', 'error');
        return;
    }

    // Desabilitar botões durante a exportação
    btnExport.disabled = true;
    toggleConsultantButtons(true);

    let clickedBtn = null;
    let originalIconHTML = '';
    if (filterConsultantName) {
        // Encontrar o botão do consultor clicado para feedback visual
        const buttons = consultantsContainer.querySelectorAll('.btn-consultant');
        for (const btn of buttons) {
            if (btn.querySelector('span').textContent === filterConsultantName) {
                clickedBtn = btn;
                const iconEl = btn.querySelector('i');
                originalIconHTML = iconEl.outerHTML;
                iconEl.className = 'fa-solid fa-spinner fa-spin';
                break;
            }
        }
    } else {
        exportSpinner.style.display = 'block';
    }

    setTimeout(() => {
        try {
            const dataRowIdx = parseInt(dataRowNumInput.value, 10) - 1;
            const outputRows = [];

            // 1. Criar array de cabeçalhos
            const headers = columnMappings.map(m => {
                if (m.outputHeader && m.outputHeader.trim() !== '') {
                    return m.outputHeader.trim();
                }
                if (m.inputColIndex !== '') {
                    const orig = originalHeaders.find(col => col.index === m.inputColIndex);
                    return orig ? orig.name : '';
                }
                return ''; // Se estiver em branco e sem título, coluna sem cabeçalho
            });
            outputRows.push(headers);

            // Determinar qual coluna contém o nome do consultor
            let consultantColIndex = 12; // Padrão
            const consultantMapping = columnMappings.find(m => 
                m.outputHeader && m.outputHeader.toLowerCase().includes('consultor')
            );
            if (consultantMapping && consultantMapping.inputColIndex !== '') {
                consultantColIndex = consultantMapping.inputColIndex;
            }

            // 2. Extrair dados correspondentes
            for (let r = dataRowIdx; r < sheetData.length; r++) {
                const row = sheetData[r];
                if (!row || row.every(val => val === null || val === undefined || val === '')) {
                    continue;
                }

                if (filterConsultantName) {
                    const rowConsultantVal = row[consultantColIndex];
                    if (!matchConsultant(rowConsultantVal, filterConsultantName)) {
                        continue;
                    }
                }

                const outputRow = columnMappings.map(m => {
                    if (m.inputColIndex !== '') {
                        const val = row[m.inputColIndex];
                        return val !== undefined && val !== null ? val : '';
                    }
                    return ''; // Se for coluna em branco
                });

                outputRows.push(outputRow);
            }

            // Validar se há registros encontrados para o consultor
            if (filterConsultantName && outputRows.length <= 1) {
                showToast(`Nenhum registro encontrado para o consultor: ${filterConsultantName}`, 'error');
                return;
            }

            // 3. Criar nova sheet do Excel
            const worksheet = XLSX.utils.aoa_to_sheet(outputRows);

            const range = XLSX.utils.decode_range(worksheet['!ref']);
            const colWidths = [];

            // Aplicar estilos (tamanho de fonte 10, negrito nos cabeçalhos) e medir largura máxima para auto-fit
            for (let c = range.s.c; c <= range.e.c; c++) {
                let maxLength = 10; // largura mínima padrão para a coluna

                for (let r = range.s.r; r <= range.e.r; r++) {
                    const cellRef = XLSX.utils.encode_cell({ r: r, c: c });
                    const cell = worksheet[cellRef];
                    
                    if (cell) {
                        // Inicializa estilo caso não exista
                        cell.s = cell.s || {};
                        cell.s.font = cell.s.font || {};
                        cell.s.font.name = 'Segoe UI';
                        cell.s.font.sz = 10; // Tamanho da fonte sempre 10 para todos os dados

                        // Se for a primeira linha (cabeçalho), aplica negrito
                        if (r === 0) {
                            cell.s.font.bold = true;
                        }

                        // Medir o maior tamanho de texto na coluna
                        if (cell.v !== undefined && cell.v !== null) {
                            const cellLength = String(cell.v).length;
                            if (cellLength > maxLength) {
                                maxLength = cellLength;
                            }
                        }
                    }
                }
                
                // Define a largura da coluna com 3 caracteres extras de margem
                colWidths.push({ wch: maxLength + 3 });
            }

            // Aplicar larguras na planilha
            worksheet['!cols'] = colWidths;

            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Dados Extraídos');

            // 4. Salvar arquivo
            let filename = outputFilenameInput.value.trim();
            if (!filename) filename = 'nps_extraido.xlsx';
            
            if (filterConsultantName) {
                // Prepend o nome do consultor no início do nome do arquivo
                if (!filename.endsWith('.xlsx') && !filename.endsWith('.xls')) {
                    filename += '.xlsx';
                }
                filename = `${filterConsultantName} - ${filename}`;
            } else {
                if (!filename.endsWith('.xlsx') && !filename.endsWith('.xls')) {
                    filename += '.xlsx';
                }
            }

            XLSX.writeFile(workbook, filename);
            showToast(`Dados exportados com sucesso em ${filename}!`, 'success');

        } catch (error) {
            console.error(error);
            showToast('Erro ao exportar dados: ' + error.message, 'error');
        } finally {
            // Re-habilitar botões e restaurar ícones
            btnExport.disabled = false;
            toggleConsultantButtons(false);
            if (clickedBtn) {
                const iconEl = clickedBtn.querySelector('i');
                if (iconEl) {
                    iconEl.className = 'fa-solid fa-file-excel';
                }
            }
            exportSpinner.style.display = 'none';
        }
    }, 600);
}

// Resetar Aplicação
function resetApp() {
    rawWorkbook = null;
    sheetData = [];
    originalHeaders = [];
    
    resetMappings();

    fileInput.value = '';
    fileNameEl.textContent = 'planilha.xlsx';
    fileSizeEl.textContent = '0 KB';
    fileInfoBox.style.display = 'none';
    dropZone.style.display = 'block';
    
    configSection.style.display = 'none';
    btnExport.disabled = true;
    toggleConsultantButtons(true);
    
    mappingList.innerHTML = '';
    setDefaultFilename();

    showToast('Arquivo removido. Pronto para um novo upload.', 'info');
}

// Renderizar os botões dos consultores na grade agrupados por loja
function renderConsultantButtons() {
    if (!consultantsContainer) return;
    consultantsContainer.innerHTML = '';
    
    Object.entries(TARGET_CONSULTANTS_BY_STORE).forEach(([storeName, consultants]) => {
        const storeGroup = document.createElement('div');
        storeGroup.className = 'store-group';
        
        const storeTitle = document.createElement('div');
        storeTitle.className = 'store-title';
        storeTitle.textContent = storeName;
        storeGroup.appendChild(storeTitle);
        
        const grid = document.createElement('div');
        grid.className = 'consultants-grid';
        
        consultants.forEach(name => {
            const btn = document.createElement('button');
            btn.className = 'btn-consultant';
            btn.disabled = true; // Desabilitado por padrão até que a planilha seja carregada
            btn.innerHTML = `<i class="fa-solid fa-file-excel"></i> <span>${name}</span>`;
            btn.addEventListener('click', () => exportData(name));
            grid.appendChild(btn);
        });
        
        storeGroup.appendChild(grid);
        consultantsContainer.appendChild(storeGroup);
    });
}

// Habilitar ou desabilitar todos os botões de consultores
function toggleConsultantButtons(disabled) {
    if (!consultantsContainer) return;
    const buttons = consultantsContainer.querySelectorAll('.btn-consultant');
    buttons.forEach(btn => {
        btn.disabled = disabled;
    });
}

// Comparação robusta de nomes de consultores
function matchConsultant(rowValue, targetName) {
    if (!rowValue || !targetName) return false;
    const cleanStr = (str) => {
        if (typeof str !== 'string') str = String(str || '');
        return str
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toUpperCase()
            .replace(/GON\s+ALVES/g, "GONCALVES")
            .trim();
    };
    const cleanRow = cleanStr(rowValue).replace(/\s+/g, ' ');
    const cleanTarget = cleanStr(targetName).replace(/\s+/g, ' ');
    
    if (cleanRow === cleanTarget) return true;
    if (cleanRow.includes(cleanTarget) || cleanTarget.includes(cleanRow)) {
        return true;
    }
    return false;
}
