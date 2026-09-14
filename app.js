// Dados do arquivo Excel
let globalData = [];

// Elementos do DOM
const searchInput = document.getElementById('searchInput');
const clearBtn = document.getElementById('clearBtn');
const resultsList = document.getElementById('resultsList');
const emptyState = document.getElementById('emptyState');
const statusText = document.getElementById('statusText');
const loader = document.getElementById('loader');
const fallbackContainer = document.getElementById('fallbackContainer');
const fileUpload = document.getElementById('fileUpload');

// Utilitário de Debounce para busca
function debounce(func, wait) {
    let timeout;
    return function (...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

// Realçar (Highlight) o texto buscado
function highlightText(text, query) {
    if (!query || !text) return text || '-';
    const str = String(text);
    // Split the query into words for multiple word highlighting
    const words = query.split(' ').filter(w => w.trim().length > 0);
    if (words.length === 0) return str;
    
    const regexPattern = words.map(w => `(${w})`).join('|');
    const regex = new RegExp(regexPattern, 'gi');
    return str.replace(regex, '<mark>$&</mark>');
}

// Processa o workbook do SheetJS
function processWorkbook(workbook) {
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    
    const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    // Normalizar chaves para facilitar busca
    globalData = rawData.map(row => {
        return {
            codigo: row['Código'] || row['Codigo'] || row['CÓDIGO'] || '-',
            descricao: row['Descrição'] || row['Descricao'] || row['DESCRIÇÃO'] || '-'
        };
    });
}

// Processar Arquivo Excel via Fetch (se rodando num servidor local/web)
async function loadExcelData() {
    try {
        loader.classList.remove('hidden');
        statusText.textContent = "Carregando base de dados de peças...";
        fallbackContainer.classList.add('hidden');

        // Tenta baixar o arquivo localmente
        const response = await fetch('./pecas_consolidadas.xlsx');
        if (!response.ok) throw new Error('Falha ao carregar o arquivo local');

        const arrayBuffer = await response.arrayBuffer();

        // Leitura usando SheetJS
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        processWorkbook(workbook);

        loader.classList.add('hidden');
        statusText.textContent = `${globalData.length} peças carregadas.`;

        // Remove text after 3 seconds
        setTimeout(() => {
            statusText.textContent = "";
        }, 3000);

        // Habilita a busca
        searchInput.disabled = false;
        searchInput.focus();

    } catch (error) {
        console.error("Erro ao carregar Excel via fetch:", error);
        loader.classList.add('hidden');
        statusText.textContent = "Para rodar localmente sem servidor, use o botão abaixo para abrir o arquivo 'pecas_consolidadas.xlsx'.";
        statusText.style.color = "var(--warning-color)";
        fallbackContainer.classList.remove('hidden');
    }
}

// Handlers de Upload Manual de Arquivo
fileUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    loader.classList.remove('hidden');
    statusText.textContent = "Processando arquivo...";
    statusText.style.color = "var(--text-secondary)";
    fallbackContainer.classList.add('hidden');

    const reader = new FileReader();
    reader.onload = function (evt) {
        try {
            const data = evt.target.result;
            const workbook = XLSX.read(data, { type: 'array' });

            processWorkbook(workbook);

            loader.classList.add('hidden');
            statusText.textContent = `${globalData.length} peças carregadas com sucesso!`;
            statusText.style.color = "var(--success-color)";

            setTimeout(() => {
                statusText.textContent = "";
            }, 3500);

            searchInput.disabled = false;
            searchInput.focus();

        } catch (err) {
            console.error(err);
            loader.classList.add('hidden');
            statusText.textContent = "Erro ao processar o arquivo selecionado.";
            statusText.style.color = "var(--danger-color)";
            fallbackContainer.classList.remove('hidden');
        }
    };
    reader.readAsArrayBuffer(file);
});

// Renderizar os cards
function renderResults(results, query) {
    if (results.length === 0) {
        emptyState.classList.add('hidden');
        resultsList.innerHTML = `
            <div class="no-results">
                Nenhuma peça encontrada para a descrição "<strong>${query}</strong>".
            </div>
        `;
        return;
    }

    emptyState.classList.add('hidden');

    // Limitando a 50 resultados para performance do DOM
    const dataToRender = results.slice(0, 50);

    const htmlCards = dataToRender.map(item => {
        return `
        <div class="result-card">
            <div class="card-header">
                <div class="serial-number">
                    <svg class="serial-icon" xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"></path>
                        <line x1="9" y1="9" x2="15" y2="9"></line>
                        <line x1="9" y1="13" x2="15" y2="13"></line>
                        <path d="M12 17h.01"></path>
                    </svg>
                    <span style="word-break: break-word;">${item.codigo}</span>
                </div>
                <button class="edit-btn" onclick="editDescription('${item.codigo}')" title="Editar Descrição">
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path>
                        <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path>
                    </svg>
                    Editar
                </button>
            </div>
            
            <div class="card-body">
                <div class="data-group" style="width: 100%;">
                    <span class="data-label">Descrição da Peça</span>
                    <span class="data-value" style="font-size: 1.1rem; line-height: 1.4;">${highlightText(item.descricao, query)}</span>
                </div>
            </div>
        </div>
        `;
    }).join('');

    resultsList.innerHTML = htmlCards;

    // Indicador caso haja mais resultados não exibidos
    if (results.length > 50) {
        resultsList.innerHTML += `
            <div style="text-align: center; font-size: 0.8rem; color: var(--text-secondary); margin-top: 1rem;">
                Exibindo 50 de ${results.length} resultados. Refine sua busca.
            </div>
        `;
    }
}

// Executar Busca
function handleSearch() {
    const query = searchInput.value.trim().toLowerCase();

    if (query.length > 0) {
        clearBtn.classList.remove('hidden');
    } else {
        clearBtn.classList.add('hidden');
        resultsList.innerHTML = '';
        emptyState.classList.remove('hidden');
        return;
    }

    // Busca apenas se tiver pelo menos 2 caracteres
    if (query.length >= 2) {
        const queryWords = query.split(' ').filter(w => w.trim().length > 0);
        
        const results = globalData.filter(item => {
            const descStr = String(item.descricao).toLowerCase();
            const codeStr = String(item.codigo).toLowerCase();
            
            // Retorna se todas as palavras da busca estiverem na descrição (ou no código)
            return queryWords.every(word => descStr.includes(word) || codeStr.includes(word));
        });
        renderResults(results, query);
    }
}

// Event Listeners
searchInput.addEventListener('input', debounce(handleSearch, 300));

clearBtn.addEventListener('click', () => {
    searchInput.value = '';
    clearBtn.classList.add('hidden');
    resultsList.innerHTML = '';
    emptyState.classList.remove('hidden');
    searchInput.focus();
});

// Tema Escuro/Claro
const themeToggleBtn = document.getElementById('themeToggleBtn');
const sunIcon = document.querySelector('.sun-icon');
const moonIcon = document.querySelector('.moon-icon');

function initTheme() {
    const savedTheme = localStorage.getItem('theme');
    
    // Default system is light now. Set to dark only if explicitly requested.
    const isDark = savedTheme === 'dark';
    
    if (isDark) {
        document.documentElement.removeAttribute('data-theme');
        sunIcon.style.display = 'block';
        moonIcon.style.display = 'none';
    } else {
        document.documentElement.setAttribute('data-theme', 'light');
        sunIcon.style.display = 'none';
        moonIcon.style.display = 'block';
    }
}

themeToggleBtn.addEventListener('click', () => {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    
    if (currentTheme === 'light') {
        document.documentElement.removeAttribute('data-theme');
        localStorage.setItem('theme', 'dark');
        sunIcon.style.display = 'block';
        moonIcon.style.display = 'none';
    } else {
        document.documentElement.setAttribute('data-theme', 'light');
        localStorage.setItem('theme', 'light');
        sunIcon.style.display = 'none';
        moonIcon.style.display = 'block';
    }
});

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    // Desabilitar input até carregar os dados
    searchInput.disabled = true;
    loadExcelData();
});

// Função para editar descrição
window.editDescription = function(codigo) {
    const senha = prompt("Digite a senha para habilitar a edição:");
    if (senha !== "Nuctech08") {
        if (senha !== null) alert("Senha incorreta!");
        return;
    }
    
    const itemIndex = globalData.findIndex(i => i.codigo === codigo);
    if (itemIndex !== -1) {
        const novaDescricao = prompt("Digite a nova descrição:", globalData[itemIndex].descricao);
        if (novaDescricao !== null && novaDescricao.trim() !== "") {
            globalData[itemIndex].descricao = novaDescricao.trim();
            handleSearch(); // Atualiza a tela com a busca atual
        }
    }
};
