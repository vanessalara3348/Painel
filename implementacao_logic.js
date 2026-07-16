// ============================================================
// implementacao_logic.js
// Lógica de implementação para o painel de monitoramento
// ============================================================

// Global array to store spreadsheet data from Monitoramento_Indicadores.xlsx
var monitoramentoData = [];

// ============================================================
// 1. loadMonitoramentoData()
//    Loads Monitoramento_Indicadores.xlsx using SheetJS (XLSX)
// ============================================================
function loadMonitoramentoData() {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', 'Monitoramento_Indicadores.xlsx', true);
    xhr.responseType = 'arraybuffer';

    xhr.onload = function () {
        if (xhr.status === 200) {
            try {
                var data = new Uint8Array(xhr.response);
                var workbook = XLSX.read(data, { type: 'array' });
                var firstSheetName = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[firstSheetName];

                // Convert sheet to array of arrays (each row is an array)
                var rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                monitoramentoData = [];

                // Skip header row (index 0), process data rows
                for (var i = 1; i < rows.length; i++) {
                    var row = rows[i];
                    if (!row || row.length === 0) continue;

                    // Columns A-H (indices 0-7)
                    var colA = row[0] || '';  // eixo number
                    var colB = row[1] || '';  // eixo name
                    var colC = row[2] || '';  // obj code
                    var colD = row[3] || '';  // obj name
                    var colE = row[4] || '';  // acao code
                    var colF = row[5] || '';  // acao name
                    var colG = row[6] || '';  // prod code
                    var colH = row[7] || '';  // prod name

                    // Column L (index 11) - indicador
                    var indicador = row[11] || '';

                    // Column P (index 15) - metaFinal
                    var metaFinal = row[15] || '';

                    // Column Q (index 16) - metaCurto
                    var metaCurto = row[16] || '';

                    // Column R (index 17) - metaMedio
                    var metaMedio = row[17] || '';

                    // Column S (index 18) - metaLongo
                    var metaLongo = row[18] || '';

                    // Build the data object
                    var item = {
                        cod: colG,
                        eixo: 'Eixo ' + colA + ' - ' + colB,
                        objetivo: colC + ' - ' + colD,
                        acao: colE + ' - ' + colF,
                        produto: colG + ' - ' + colH,
                        indicador: indicador,
                        metaFinal: metaFinal,
                        metaCurto: metaCurto,
                        metaMedio: metaMedio,
                        metaLongo: metaLongo
                    };

                    monitoramentoData.push(item);
                }

                console.log('Monitoramento data loaded:', monitoramentoData.length, 'rows');
                buildImplementacaoTable();

            } catch (e) {
                console.error('Error parsing Monitoramento_Indicadores.xlsx:', e);
                loadFallbackData();
            }
        } else {
            console.error('Error loading Monitoramento_Indicadores.xlsx, status:', xhr.status);
            loadFallbackData();
        }
    };

    xhr.onerror = function () {
        console.error('Network error loading Monitoramento_Indicadores.xlsx');
        loadFallbackData();
    };

    xhr.send();
}

/**
 * Fallback data in case the spreadsheet cannot be loaded
 */
function loadFallbackData() {
    monitoramentoData = [
        {
            cod: 'P001',
            eixo: 'Eixo 1 - Governança e Gestão',
            objetivo: 'OBJ01 - Fortalecer a governança',
            acao: 'A001 - Implementar sistema de gestão',
            produto: 'P001 - Plano de gestão elaborado',
            indicador: 'Número de planos elaborados',
            metaFinal: 10,
            metaCurto: 3,
            metaMedio: 5,
            metaLongo: 2
        },
        {
            cod: 'P002',
            eixo: 'Eixo 2 - Desenvolvimento Sustentável',
            objetivo: 'OBJ02 - Promover o desenvolvimento',
            acao: 'A002 - Fomentar projetos sustentáveis',
            produto: 'P002 - Projetos implementados',
            indicador: 'Número de projetos',
            metaFinal: 20,
            metaCurto: 6,
            metaMedio: 8,
            metaLongo: 6
        },
        {
            cod: 'P003',
            eixo: 'Eixo 3 - Inclusão Social',
            objetivo: 'OBJ03 - Ampliar a inclusão',
            acao: 'A003 - Capacitar comunidades',
            produto: 'P003 - Comunidades capacitadas',
            indicador: 'Número de comunidades',
            metaFinal: 15,
            metaCurto: 5,
            metaMedio: 5,
            metaLongo: 5
        }
    ];

    console.log('Using fallback data for monitoramento');
    buildImplementacaoTable();
}

// ============================================================
// 2. calcularSituacaoMeta(cod, indicador, tipoMeta)
//    Calculates the status of a meta based on approved actions
// ============================================================
function calcularSituacaoMeta(cod, indicador, tipoMeta) {
    // Define the year ranges for each meta type
    var yearRanges = {
        'curto': [2025, 2026, 2027],
        'medio': [2028, 2029, 2030],
        'longo': [2031, 2032, 2033, 2034, 2035],
        'final': [2025, 2026, 2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034, 2035]
    };

    var validYears = yearRanges[tipoMeta] || yearRanges['final'];

    // Sum values from approved products in mockDatabase
    var soma = 0;
    var hasApproved = false;

    if (typeof mockDatabase !== 'undefined' && mockDatabase.acoes && Array.isArray(mockDatabase.acoes)) {
        for (var i = 0; i < mockDatabase.acoes.length; i++) {
            var acao = mockDatabase.acoes[i];

            // Check if the product matches by code and indicator, and is approved
            if (acao.status === 'Aprovado' &&
                acao.codProd === cod &&
                acao.indicador === indicador) {

                // Check if the horizonte (year) falls within the valid range
                var horizonte = parseInt(acao.horizonte, 10);
                if (validYears.indexOf(horizonte) !== -1) {
                    hasApproved = true;
                    var valor = parseFloat(acao.valor);
                    if (!isNaN(valor)) {
                        soma += valor;
                    }
                }
            }
        }
    }

    // Find the meta value from monitoramentoData
    var metaValue = 0;
    for (var j = 0; j < monitoramentoData.length; j++) {
        if (monitoramentoData[j].cod === cod) {
            switch (tipoMeta) {
                case 'curto':
                    metaValue = parseFloat(monitoramentoData[j].metaCurto) || 0;
                    break;
                case 'medio':
                    metaValue = parseFloat(monitoramentoData[j].metaMedio) || 0;
                    break;
                case 'longo':
                    metaValue = parseFloat(monitoramentoData[j].metaLongo) || 0;
                    break;
                case 'final':
                default:
                    metaValue = parseFloat(monitoramentoData[j].metaFinal) || 0;
                    break;
            }
            break;
        }
    }

    // Calculate percentage
    var percentual = 0;
    if (metaValue > 0) {
        percentual = (soma / metaValue) * 100;
    }

    // Determine status based on percentage
    if (!hasApproved || percentual < 1) {
        return 'À Iniciar';
    } else if (percentual >= 1 && percentual < 49) {
        return 'Iniciado';
    } else if (percentual >= 50 && percentual < 99) {
        return 'Em Desenvolvimento';
    } else if (percentual >= 100) {
        return 'Concluído';
    }

    return 'À Iniciar';
}

// ============================================================
// 3. buildImplementacaoTable()
//    Renders the implementation monitoring table
// ============================================================
function buildImplementacaoTable() {
    var tbody = document.getElementById('table-implementacao-body');
    if (!tbody) {
        console.warn('Element table-implementacao-body not found');
        return;
    }

    // Get filter values
    var filterOrgao = document.getElementById('filter-imp-orgao');
    var filterMeta = document.getElementById('filter-imp-meta');
    var filterStatus = document.getElementById('filter-imp-status');
    var filterTransversal = document.getElementById('filter-imp-transversal');
    var searchText = document.getElementById('search-imp-text');

    var orgaoValue = filterOrgao ? filterOrgao.value : '';
    var metaValue = filterMeta ? filterMeta.value : '';
    var statusValue = filterStatus ? filterStatus.value : '';
    var transversalValue = filterTransversal ? filterTransversal.value : '';
    var textValue = searchText ? searchText.value.toLowerCase().trim() : '';

    // Clear existing rows
    tbody.innerHTML = '';

    // Badge color mapping
    var badgeColors = {
        'À Iniciar': 'bg-gray-100 text-gray-800',
        'Iniciado': 'bg-blue-100 text-blue-800',
        'Em Desenvolvimento': 'bg-amber-100 text-amber-800',
        'Concluído': 'bg-emerald-100 text-emerald-800'
    };

    // Iterate over monitoramentoData and build rows
    for (var i = 0; i < monitoramentoData.length; i++) {
        var item = monitoramentoData[i];

        // Calculate status for META PARCIAL_curto
        var status = calcularSituacaoMeta(item.cod, item.indicador, 'curto');

        // Apply filters
        // Filter by status
        if (statusValue && statusValue !== '' && status !== statusValue) {
            continue;
        }

        // Filter by meta type (if applicable)
        if (metaValue && metaValue !== '') {
            // Meta filter can be used to switch between curto/medio/longo/final
            // For now, we only display curto by default
        }

        // Filter by text search (searches across multiple fields)
        if (textValue) {
            var searchableText = (
                item.cod + ' ' +
                item.eixo + ' ' +
                item.acao + ' ' +
                item.produto + ' ' +
                item.indicador
            ).toLowerCase();

            if (searchableText.indexOf(textValue) === -1) {
                continue;
            }
        }

        // Filter by orgao (if the field exists on item)
        if (orgaoValue && orgaoValue !== '') {
            if (item.orgao && item.orgao !== orgaoValue) {
                continue;
            }
        }

        // Filter by transversal (if the field exists on item)
        if (transversalValue && transversalValue !== '') {
            if (item.transversal && item.transversal !== transversalValue) {
                continue;
            }
        }

        // Get meta curto value for display
        var metaCurtoDisplay = item.metaCurto || '-';

        // Get badge color class
        var badgeClass = badgeColors[status] || 'bg-gray-100 text-gray-800';

        // Build table row
        var tr = document.createElement('tr');
        tr.className = 'hover:bg-gray-50 transition-colors';
        tr.innerHTML =
            '<td class="px-4 py-3 text-sm font-medium text-gray-900">' + escapeHtml(String(item.cod)) + '</td>' +
            '<td class="px-4 py-3 text-sm text-gray-600">' +
                '<div class="font-medium">' + escapeHtml(String(item.eixo)) + '</div>' +
                '<div class="text-xs text-gray-400">' + escapeHtml(String(item.acao)) + '</div>' +
            '</td>' +
            '<td class="px-4 py-3 text-sm text-gray-600">' + escapeHtml(String(item.produto)) + '</td>' +
            '<td class="px-4 py-3 text-sm text-gray-600">' + escapeHtml(String(item.indicador)) + '</td>' +
            '<td class="px-4 py-3 text-sm text-gray-900 font-medium text-center">' + escapeHtml(String(metaCurtoDisplay)) + '</td>' +
            '<td class="px-4 py-3 text-sm text-center">' +
                '<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ' + badgeClass + '">' +
                    escapeHtml(status) +
                '</span>' +
            '</td>';

        tbody.appendChild(tr);
    }

    // Refresh Lucide icons if available
    if (typeof lucide !== 'undefined' && lucide.createIcons) {
        lucide.createIcons();
    }

    // Update summary cards
    updateImplementacaoSummary();
}

/**
 * Helper function to escape HTML special characters
 */
function escapeHtml(text) {
    if (!text) return '';
    var div = document.createElement('div');
    div.appendChild(document.createTextNode(text));
    return div.innerHTML;
}

// ============================================================
// 4. filterImplementacao()
//    Triggered by filter changes, re-builds the table
// ============================================================
function filterImplementacao() {
    buildImplementacaoTable();
}

// ============================================================
// 5. updateImplementacaoSummary()
//    Updates summary cards for META PARCIAL_curto
//    Shows counts and percentages for each status category
// ============================================================
function updateImplementacaoSummary() {
    // Count statuses across all monitoramentoData items
    var counts = {
        'À Iniciar': 0,
        'Iniciado': 0,
        'Em Desenvolvimento': 0,
        'Concluído': 0
    };

    var total = monitoramentoData.length;

    for (var i = 0; i < monitoramentoData.length; i++) {
        var item = monitoramentoData[i];
        var status = calcularSituacaoMeta(item.cod, item.indicador, 'curto');
        if (counts.hasOwnProperty(status)) {
            counts[status]++;
        }
    }

    // Period label
    var periodLabel = 'Meta Parcial (2025 - 2027)';

    // Update summary card elements (if they exist in the DOM)
    var cardIniciar = document.getElementById('card-imp-iniciar');
    var cardIniciado = document.getElementById('card-imp-iniciado');
    var cardDesenvolvimento = document.getElementById('card-imp-desenvolvimento');
    var cardConcluido = document.getElementById('card-imp-concluido');
    var periodElement = document.getElementById('imp-period-label');

    if (periodElement) {
        periodElement.textContent = periodLabel;
    }

    // Helper to calculate percentage
    function calcPercent(count) {
        if (total === 0) return '0%';
        return Math.round((count / total) * 100) + '%';
    }

    // Update "À Iniciar" card
    if (cardIniciar) {
        var countEl = cardIniciar.querySelector('.card-count');
        var percentEl = cardIniciar.querySelector('.card-percent');
        if (countEl) countEl.textContent = counts['À Iniciar'];
        if (percentEl) percentEl.textContent = calcPercent(counts['À Iniciar']);
    }

    // Update "Iniciado" card
    if (cardIniciado) {
        var countEl = cardIniciado.querySelector('.card-count');
        var percentEl = cardIniciado.querySelector('.card-percent');
        if (countEl) countEl.textContent = counts['Iniciado'];
        if (percentEl) percentEl.textContent = calcPercent(counts['Iniciado']);
    }

    // Update "Em Desenvolvimento" card
    if (cardDesenvolvimento) {
        var countEl = cardDesenvolvimento.querySelector('.card-count');
        var percentEl = cardDesenvolvimento.querySelector('.card-percent');
        if (countEl) countEl.textContent = counts['Em Desenvolvimento'];
        if (percentEl) percentEl.textContent = calcPercent(counts['Em Desenvolvimento']);
    }

    // Update "Concluído" card
    if (cardConcluido) {
        var countEl = cardConcluido.querySelector('.card-count');
        var percentEl = cardConcluido.querySelector('.card-percent');
        if (countEl) countEl.textContent = counts['Concluído'];
        if (percentEl) percentEl.textContent = calcPercent(counts['Concluído']);
    }

    console.log('Summary updated -', periodLabel, ':', counts);
}
