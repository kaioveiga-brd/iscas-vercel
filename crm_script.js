/**
 * Legado CRM Dashboard - Script de Comportamento v3
 * Controle de Pipeline em Grid + Gráficos Chart.js + Diário de Atividades Semanal
 */

// Estado inicial carregado da planilha do Excel
const INITIAL_CRM_STATE = {
  "abaporueducacional": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-19",
    "notes": ""
  },
  "assetingles": {
    "status": "Não / Descartado",
    "date": "",
    "notes": "Ja tem Equipe"
  },
  "believeidiomas": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-26",
    "notes": ""
  },
  "blackwhiteschool": {
    "status": "Não / Descartado",
    "date": "",
    "notes": "Ja tem Equipe"
  },
  "cursogalileo": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": ""
  },
  "daccordidiomas": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": ""
  },
  "eweschool": {
    "status": "2 · Follow-up enviado",
    "date": "2026-07-05",
    "notes": "Momento de fechamento, pediu pra conversar dps"
  },
  "englishyou": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": "Comercial enviou o material ao dono, Ricardo"
  },
  "exatoassessoriaestudantil": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": ""
  },
  "exegeseidiomas": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-26",
    "notes": ""
  },
  "apoiovestibulinhos": {
    "status": "Não / Descartado",
    "date": "2026-06-24",
    "notes": "basicamente só dao aulas particulares, n conseguem crescer"
  },
  "doupescoladeidiomas": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-26",
    "notes": ""
  },
  "excellentglobalidiomas": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-26",
    "notes": ""
  },
  "londoncenteridiomas": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-26",
    "notes": ""
  },
  "mogianacursosmogianavigilantes": {
    "status": "Respondeu",
    "date": "2026-06-26",
    "notes": "Passou para o setor responsável"
  },
  "newwayidiomas": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": ""
  },
  "quickidiomas": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": ""
  },
  "tonightidiomas": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": ""
  },
  "flexuspalavrizar": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": ""
  },
  "getfluentenglishcenter": {
    "status": "R1 marcada",
    "date": "2026-06-25",
    "notes": ""
  },
  "interchangeidiomas": {
    "status": "Não / Descartado",
    "date": "",
    "notes": "Não tiveram interesse"
  },
  "libertyidiomasinformatica": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": ""
  },
  "londonenglishschool": {
    "status": "Não / Descartado",
    "date": "",
    "notes": "Mensagem nao chegou ao destinatario"
  },
  "michiganidiomas": {
    "status": "R1 marcada",
    "date": "2026-06-23",
    "notes": ""
  },
  "mundoenem": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": "Bot de Atendimento"
  },
  "newinterchange": {
    "status": "Não / Descartado",
    "date": "",
    "notes": "Ja possuem Equipe"
  },
  "progressovestibulares": {
    "status": "A Contatar",
    "date": "",
    "notes": ""
  },
  "purrfectidiomas": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": ""
  },
  "senseinglessenseprivateinstruction": {
    "status": "Não / Descartado",
    "date": "2026-06-24",
    "notes": "Mensagem nao chegou ao destinatario"
  },
  "thehousepiracicaba": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": "Demonstrou interesse, retornar em breve"
  },
  "thewayidiomas": {
    "status": "A Contatar",
    "date": "",
    "notes": ""
  },
  "whynotinstitute": {
    "status": "Não / Descartado",
    "date": "",
    "notes": "Nao era o momento"
  },
  "cursoborges": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": "Falou de ver com a filha"
  },
  "excellentglobal": {
    "status": "2 · Follow-up enviado",
    "date": "2026-06-24",
    "notes": ""
  },
  "callidiomas": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "verballanguage": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "gathercentrodeensino": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "globalcommlanguagesolutions": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "cnesisaprimoramentodeestudos": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "eagleschool": {
    "status": "Respondeu",
    "date": "",
    "notes": "Carla respondeu, diz que não quer crescer pra não perder qualidade."
  },
  "iyenglishimproveyourenglish": {
    "status": "Respondeu",
    "date": "",
    "notes": "Alessandra respondeu achando interessante."
  },
  "openaccessenglishschool": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "brightlanguageschool": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "targetlanguagecentertlcidiomas": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "globalschool": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "institutoculturaliberoamericanoici": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "escolahispanoamericana": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "lsccallanvalinhos": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "dreamsschool": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "tictacschool": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "maquifisica": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "alcanceprevestibular": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "cursoecolegiouni": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "cursoflexus": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "maisvestibulares": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "novocolegiofrancacursinho": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "jackandjillenglish": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "goldenbearidiomas": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "upperenglish": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "forallenglish": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "deeperidiomas": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "trinityenglishschool": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "evenflow": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  },
  "ccbeufranca": {
    "status": "1 · Mensagem enviada",
    "date": "",
    "notes": ""
  }
};

// Estatísticas históricas copiadas do caderno de Kaio (Baseline)
const INITIAL_CRM_ACTIVITY = {
    '2026-06-19': { iniciado: 13, followup: 6, touchpoints: 19, resposta: 6, agendada: 2, reuniao_feita: 0, descarte: 0 },
    '2026-06-22': { iniciado: 0, followup: 14, touchpoints: 12, resposta: 5, agendada: 0, reuniao_feita: 0, descarte: 0 },
    '2026-06-23': { iniciado: 0, followup: 1, touchpoints: 0, resposta: 0, agendada: 0, reuniao_feita: 0, descarte: 0 },
    '2026-06-24': { iniciado: 4, followup: 11, touchpoints: 15, resposta: 6, agendada: 1, reuniao_feita: 0, descarte: 0 },
    '2026-06-25': { iniciado: 10, followup: 5, touchpoints: 15, resposta: 5, agendada: 1, reuniao_feita: 6, descarte: 0 },
    '2026-06-26': { iniciado: 0, followup: 12, touchpoints: 0, resposta: 6, agendada: 0, reuniao_feita: 1, descarte: 0 },
    '2026-06-27': { iniciado: 0, followup: 16, touchpoints: 13, resposta: 7, agendada: 0, reuniao_feita: 0, descarte: 0 }
};

let crmData = {};
let crmActivity = {};
let activeLeadKey = null;
let currentSortField = 'name';
let currentSortOrder = 'asc';

// Semana ativa (Date object do Monday da semana exibida)
let currentWeekStart = null;

// Instâncias Globais dos Gráficos
let funnelChartInstance = null;
let conversionChartInstance = null;

document.addEventListener('DOMContentLoaded', () => {
    initData();
    initWeek();
    populateFilters();
    initCharts();
    renderAll();
    setupEventListeners();
});

// Inicialização dos Dados
function initData() {
    // 1. Dados dos Leads
    const stored = localStorage.getItem('legado_crm_data');
    if (stored) {
        try {
            crmData = JSON.parse(stored);
        } catch (e) {
            console.error("Erro ao ler localStorage, reiniciando...", e);
            crmData = {};
        }
    }
    
    // Garante que todos os prospects do db.js existem no crmData
    let updated = false;
    Object.keys(PROSPECTS_DB).forEach(key => {
        if (!crmData[key]) {
            const initial = INITIAL_CRM_STATE[key] || {};
            crmData[key] = {
                status: initial.status || 'A Contatar',
                date: initial.date || '',
                notes: initial.notes || '',
                r1Done: false
            };
            updated = true;
        }
    });
    if (updated) saveToLocalStorage();

    // 2. Dados de Atividades / Diário
    const storedActivity = localStorage.getItem('legado_crm_activity');
    if (storedActivity) {
        try {
            crmActivity = JSON.parse(storedActivity);
        } catch (e) {
            console.error("Erro ao ler atividades do localStorage, reiniciando...", e);
            crmActivity = {};
        }
    }
    
    // Carrega baseline se o diário estiver vazio
    let activityUpdated = false;
    Object.keys(INITIAL_CRM_ACTIVITY).forEach(date => {
        if (!crmActivity[date]) {
            crmActivity[date] = { ...INITIAL_CRM_ACTIVITY[date] };
            activityUpdated = true;
        }
    });
    if (activityUpdated) saveActivityToLocalStorage();
}

function saveToLocalStorage() {
    localStorage.setItem('legado_crm_data', JSON.stringify(crmData));
}

function saveActivityToLocalStorage() {
    localStorage.setItem('legado_crm_activity', JSON.stringify(crmActivity));
}

// Configura a semana ativa para a semana atual
function initWeek() {
    const today = new Date();
    currentWeekStart = getMonday(today);
}

// Obtém a segunda-feira correspondente à data
function getMonday(d) {
    d = new Date(d);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1); // ajusta para segunda
    const monday = new Date(d.setDate(diff));
    monday.setHours(0, 0, 0, 0);
    return monday;
}

// Popula opções de cidade nos filtros
function populateFilters() {
    const cities = new Set();
    Object.keys(PROSPECTS_DB).forEach(key => {
        if (PROSPECTS_DB[key].cidade) {
            cities.add(PROSPECTS_DB[key].cidade.trim());
        }
    });

    const filterCity = document.getElementById('filter-city');
    Array.from(cities).sort().forEach(city => {
        const opt = document.createElement('option');
        opt.value = city.toLowerCase();
        opt.textContent = city;
        filterCity.appendChild(opt);
    });
}

// Inicialização dos Gráficos com Chart.js
function initCharts() {
    const ctxFunnel = document.getElementById('funnelChart').getContext('2d');
    const ctxConversion = document.getElementById('conversionChart').getContext('2d');

    Chart.defaults.color = '#AAAAAA';
    Chart.defaults.font.family = "'Outfit', sans-serif";

    funnelChartInstance = new Chart(ctxFunnel, {
        type: 'bar',
        data: {
            labels: ['A Contatar', 'M1 Enviada', 'Follow-up', 'Respondeu', 'R1 Marcada', 'Descartado'],
            datasets: [{
                label: 'Prospects por Etapa',
                data: [0, 0, 0, 0, 0, 0],
                backgroundColor: [
                    '#444444',
                    '#E05A1C',
                    '#9b59b6',
                    '#2ecc71',
                    '#f1c40f',
                    '#e74c3c'
                ],
                borderWidth: 0,
                borderRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: 'rgba(255, 255, 255, 0.05)' }
                },
                x: {
                    grid: { display: false }
                }
            }
        }
    });

    conversionChartInstance = new Chart(ctxConversion, {
        type: 'doughnut',
        data: {
            labels: ['Pendente (Sem Retorno)', 'Respondeu', 'R1 Marcada', 'Descartado'],
            datasets: [{
                data: [0, 0, 0, 0],
                backgroundColor: [
                    '#444444',
                    '#2ecc71',
                    '#f1c40f',
                    '#e74c3c'
                ],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: { boxWidth: 12, padding: 15 }
                }
            },
            cutout: '65%'
        }
    });
}

// Roda a renderização geral do CRM
function renderAll() {
    renderKPIsAndTable();
    renderDiario();
}

// Renderiza a tabela de leads e atualiza os KPIs superiores
function renderKPIsAndTable() {
    const query = document.getElementById('search-input').value.toLowerCase().trim();
    const cityFilter = document.getElementById('filter-city').value;
    const nicheFilter = document.getElementById('filter-niche').value;
    const statusFilter = document.getElementById('filter-status').value;

    const tbody = document.getElementById('pipeline-table-body');
    tbody.innerHTML = '';

    let totalCount = 0;
    let filteredCount = 0;
    let contactedCount = 0;
    let respondedCount = 0;
    let r1Count = 0;

    const stageCounts = {
        'A Contatar': 0,
        '1 · Mensagem enviada': 0,
        '2 · Follow-up enviado': 0,
        'Respondeu': 0,
        'R1 marcada': 0,
        'Não / Descartado': 0
    };

    const list = [];

    Object.keys(PROSPECTS_DB).forEach(key => {
        totalCount++;
        const p = PROSPECTS_DB[key];
        const state = crmData[key] || { status: 'A Contatar', date: '', notes: '' };

        const isCursinho = p.nicho === 'cursinho' || p.nomeProspect.toLowerCase().includes('cursinho') || p.nomeProspect.toLowerCase().includes('vestibulares');
        const nicheText = isCursinho ? 'cursinho' : 'escola de idiomas';

        const matchesQuery = p.nomeProspect.toLowerCase().includes(query) || p.cidade.toLowerCase().includes(query) || nicheText.includes(query);
        const matchesCity = !cityFilter || p.cidade.toLowerCase() === cityFilter;
        const matchesNiche = !nicheFilter || nicheText === nicheFilter;
        const matchesStatus = !statusFilter || state.status === statusFilter;

        if (matchesQuery && matchesCity && matchesNiche && matchesStatus) {
            filteredCount++;
            
            if (state.status !== 'A Contatar') contactedCount++;
            if (state.status === 'Respondeu' || state.status === 'R1 marcada') respondedCount++;
            if (state.status === 'R1 marcada') r1Count++;

            if (stageCounts[state.status] !== undefined) {
                stageCounts[state.status]++;
            }

            list.push({
                key: key,
                name: p.nomeProspect,
                niche: isCursinho ? 'Cursinho' : 'Idiomas',
                nicheRaw: p.nicho,
                city: p.cidade,
                followers: parseInt(p.seguidores.replace(/[^0-9]/g, '')) || 0,
                followersStr: p.seguidores,
                status: state.status,
                date: state.date,
                notes: state.notes
            });
        }
    });

    sortList(list);

    list.forEach(item => {
        const tr = document.createElement('tr');
        tr.dataset.key = item.key;
        
        const selectClass = getStatusClass(item.status);
        const labelClass = item.niche === 'Cursinho' ? 'table-niche-cursinho' : 'table-niche-idiomas';
        
        const cleanInsta = PROSPECTS_DB[item.key].instagram ? PROSPECTS_DB[item.key].instagram.replace('@', '') : '';

        tr.innerHTML = '<td class="table-row-clickable font-bold" onclick="openLeadDrawer(\'' + item.key + '\')">' + item.name + '</td>' +
            '<td class="table-row-clickable" onclick="openLeadDrawer(\'' + item.key + '\')"><span class="table-niche-tag ' + labelClass + '">' + item.niche + '</span></td>' +
            '<td class="table-row-clickable" onclick="openLeadDrawer(\'' + item.key + '\')">' + item.city + '</td>' +
            '<td class="table-row-clickable" onclick="openLeadDrawer(\'' + item.key + '\')">' + item.followersStr + '</td>' +
            '<td>' +
                '<select class="table-status-select ' + selectClass + '" onchange="changeRowStatus(\'' + item.key + '\', this.value)">' +
                    '<option value="A Contatar" ' + (item.status === 'A Contatar' ? 'selected' : '') + '>A Contatar</option>' +
                    '<option value="1 · Mensagem enviada" ' + (item.status === '1 · Mensagem enviada' ? 'selected' : '') + '>1 · Mensagem enviada</option>' +
                    '<option value="2 · Follow-up enviado" ' + (item.status === '2 · Follow-up enviado' ? 'selected' : '') + '>2 · Follow-up enviado</option>' +
                    '<option value="Respondeu" ' + (item.status === 'Respondeu' ? 'selected' : '') + '>Respondeu</option>' +
                    '<option value="R1 marcada" ' + (item.status === 'R1 marcada' ? 'selected' : '') + '>R1 marcada</option>' +
                    '<option value="Não / Descartado" ' + (item.status === 'Não / Descartado' ? 'selected' : '') + '>Não / Descartado</option>' +
                '</select>' +
            '</td>' +
            '<td>' +
                '<input type="date" class="table-date-input" value="' + item.date + '" onchange="changeRowDate(\'' + item.key + '\', this.value)">' +
            '</td>' +
            '<td>' +
                '<div class="table-actions">' +
                    '<button class="table-btn-action" onclick="openLeadDrawer(\'' + item.key + '\')" title="Ver Dossiê e Copys"><i class="fa-solid fa-file-invoice"></i></button>' +
                    '<a href="https://wa.me/5519992253680" target="_blank" class="table-btn-action btn-wa" title="Abrir WhatsApp"><i class="fa-brands fa-whatsapp"></i></a>' +
                    '<a href="https://www.instagram.com/' + cleanInsta + '" target="_blank" class="table-btn-action btn-insta" title="Abrir Instagram"><i class="fa-brands fa-instagram"></i></a>' +
                '</div>' +
            '</td>';
            
        tbody.appendChild(tr);
    });

    document.getElementById('kpi-total').textContent = totalCount;
    document.getElementById('kpi-contacted').textContent = contactedCount;
    document.getElementById('kpi-r1').textContent = r1Count;
    
    const rate = contactedCount > 0 ? Math.round((respondedCount / contactedCount) * 100) : 0;
    document.getElementById('kpi-response-rate').textContent = rate + '%';

    document.getElementById('stat-total').textContent = totalCount;
    document.getElementById('stat-filtered').textContent = filteredCount;

    updateCharts(stageCounts);
}

function getStatusClass(status) {
    switch (status) {
        case '1 · Mensagem enviada': return 'status-m1-enviada';
        case '2 · Follow-up enviado': return 'status-followup';
        case 'Respondeu': return 'status-respondeu';
        case 'R1 marcada': return 'status-r1';
        case 'Não / Descartado': return 'status-descartado';
        default: return 'status-a-contatar';
    }
}

// Alteração de status com logging automático
function changeRowStatus(key, value) {
    if (crmData[key]) {
        const prevStatus = crmData[key].status;
        if (prevStatus !== value) {
            crmData[key].status = value;
            saveToLocalStorage();
            
            // Regista evento automático de atividade na data de hoje
            logActivityEvent(value);
            
            renderAll();
        }
    }
}

function changeRowDate(key, value) {
    if (crmData[key]) {
        crmData[key].date = value;
        saveToLocalStorage();
        renderAll();
    }
}

function sortList(list) {
    list.sort((a, b) => {
        let valA = a[currentSortField];
        let valB = b[currentSortField];

        if (typeof valA === 'string') {
            return currentSortOrder === 'asc' 
                ? valA.localeCompare(valB) 
                : valB.localeCompare(valA);
        } else {
            return currentSortOrder === 'asc' 
                ? valA - valB 
                : valB - valA;
        }
    });
}

function updateCharts(stageCounts) {
    if (funnelChartInstance) {
        funnelChartInstance.data.datasets[0].data = [
            stageCounts['A Contatar'],
            stageCounts['1 · Mensagem enviada'],
            stageCounts['2 · Follow-up enviado'],
            stageCounts['Respondeu'],
            stageCounts['R1 marcada'],
            stageCounts['Não / Descartado']
        ];
        funnelChartInstance.update();
    }

    if (conversionChartInstance) {
        const pendente = stageCounts['1 · Mensagem enviada'] + stageCounts['2 · Follow-up enviado'];
        const respondeu = stageCounts['Respondeu'];
        const r1 = stageCounts['R1 marcada'];
        const descartado = stageCounts['Não / Descartado'];

        conversionChartInstance.data.datasets[0].data = [pendente, respondeu, r1, descartado];
        conversionChartInstance.update();
    }
}

// Lógica de Logging Automático
function logActivityEvent(status) {
    const todayStr = getLocalDateString();
    ensureActivityDate(todayStr);

    if (status === '1 · Mensagem enviada') {
        crmActivity[todayStr].iniciado++;
        crmActivity[todayStr].touchpoints++;
    } else if (status === '2 · Follow-up enviado') {
        crmActivity[todayStr].followup++;
        crmActivity[todayStr].touchpoints++;
    } else if (status === 'Respondeu') {
        crmActivity[todayStr].resposta++;
    } else if (status === 'R1 marcada') {
        crmActivity[todayStr].agendada++;
    } else if (status === 'Não / Descartado') {
        crmActivity[todayStr].descarte++;
    }

    saveActivityToLocalStorage();
}

function ensureActivityDate(dateStr) {
    if (!crmActivity[dateStr]) {
        crmActivity[dateStr] = {
            iniciado: 0,
            followup: 0,
            touchpoints: 0,
            resposta: 0,
            agendada: 0,
            reuniao_feita: 0,
            descarte: 0
        };
    }
}

function getLocalDateString() {
    const d = new Date();
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return year + '-' + month + '-' + day;
}

// Renderização do Diário Semanal
function renderDiario() {
    if (!currentWeekStart) return;

    const tbody = document.getElementById('diario-table-body');
    tbody.innerHTML = '';

    const label = document.getElementById('label-current-week');
    const monday = new Date(currentWeekStart);
    const sunday = new Date(currentWeekStart);
    sunday.setDate(monday.getDate() + 6);

    // Formata cabeçalho da semana
    label.textContent = 'Semana: ' + formatDateString(monday) + ' a ' + formatDateString(sunday);

    const diasNomes = ['Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'];

    for (let i = 0; i < 7; i++) {
        const currentDay = new Date(currentWeekStart);
        currentDay.setDate(currentDay.getDate() + i);
        
        const dateStr = formatDateStringFull(currentDay);
        const dayLabel = diasNomes[i] + ' · ' + formatDateString(currentDay);

        const data = crmActivity[dateStr] || {
            iniciado: 0,
            followup: 0,
            touchpoints: 0,
            resposta: 0,
            agendada: 0,
            reuniao_feita: 0,
            descarte: 0
        };

        const totalTouchpoints = data.iniciado + data.followup + (data.touchpoints - data.iniciado - data.followup > 0 ? (data.touchpoints - data.iniciado - data.followup) : 0);

        const tr = document.createElement('tr');
        
        // Verifica se é o dia de hoje para destacar a linha
        if (dateStr === getLocalDateString()) {
            tr.style.backgroundColor = 'rgba(224, 90, 28, 0.05)';
            tr.style.borderLeft = '3px solid var(--laranja)';
        }

        tr.innerHTML = '<td class="day-cell">' + dayLabel + '</td>' +
            '<td>' + data.iniciado + '</td>' +
            '<td>' + data.followup + '</td>' +
            '<td class="highlight-cell">' + totalTouchpoints + '</td>' +
            '<td>' + data.resposta + '</td>' +
            '<td>' + data.agendada + '</td>' +
            '<td>' + data.reuniao_feita + '</td>' +
            '<td>' + data.descarte + '</td>';

        tbody.appendChild(tr);
    }
}

function formatDateString(date) {
    const d = String(date.getDate()).padStart(2, '0');
    const m = String(date.getMonth() + 1).padStart(2, '0');
    return d + '/' + m;
}

function formatDateStringFull(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return year + '-' + month + '-' + day;
}

// Navegação de Semanas
function navigateWeek(weeksDiff) {
    currentWeekStart.setDate(currentWeekStart.getDate() + (weeksDiff * 7));
    renderDiario();
}

// Gaveta de Lead (Modal lateral)
function openLeadDrawer(key) {
    activeLeadKey = key;
    const data = PROSPECTS_DB[key];
    const state = crmData[key] || { status: 'A Contatar', date: '', notes: '', r1Done: false };

    if (!data) return;

    const isCursinho = data.nicho === 'cursinho' || data.nomeProspect.toLowerCase().includes('cursinho') || data.nomeProspect.toLowerCase().includes('vestibulares');
    document.getElementById('drawer-lead-niche').textContent = isCursinho ? 'CURSINHO' : 'IDIOMAS';
    document.getElementById('drawer-lead-name').textContent = data.nomeProspect;
    document.getElementById('drawer-lead-city').textContent = data.cidade + ' - SP';

    const cleanInsta = data.instagram ? data.instagram.replace('@', '') : '';
    document.getElementById('drawer-link-insta').href = 'https://www.instagram.com/' + cleanInsta;
    document.getElementById('drawer-link-wa').href = 'https://wa.me/5519992253680';

    document.getElementById('drawer-ativo').textContent = data.ativoDesperdicado || 'N/A';
    
    const lacunaEl = document.getElementById('drawer-lacuna');
    if (typeof data.lacuna === 'object' && data.lacuna !== null) {
        lacunaEl.innerHTML = '<strong>Entrega:</strong> ' + (data.lacuna.entrega || '') + '<br><strong>Comunica:</strong> ' + (data.lacuna.comunica || '');
    } else {
        lacunaEl.textContent = data.lacuna || 'N/A';
    }
    
    document.getElementById('drawer-achado').textContent = data.achadoMaisForte || 'N/A';

    const compUl = document.getElementById('drawer-concorrentes');
    compUl.innerHTML = '';
    if (data.concorrentes && Array.isArray(data.concorrentes)) {
        data.concorrentes.forEach(c => {
            const li = document.createElement('li');
            li.textContent = typeof c === 'string' ? c : (c.nome || '');
            compUl.appendChild(li);
        });
    } else {
        compUl.innerHTML = '<li>Nenhum concorrente cadastrado.</li>';
    }

    const copys = generateDynamicCopys(key, data);
    document.getElementById('copy-text-m1').textContent = copys.m1;
    document.getElementById('copy-text-fup').textContent = copys.fup;

    // Form settings
    document.getElementById('lead-status-select').value = state.status;
    document.getElementById('lead-next-touch').value = state.date || '';
    document.getElementById('lead-notes').value = state.notes || '';
    document.getElementById('lead-r1-done').checked = !!state.r1Done;

    document.getElementById('lead-drawer').classList.add('open');
    switchTab('tab-analise');
}

function closeLeadDrawer() {
    document.getElementById('lead-drawer').classList.remove('open');
    activeLeadKey = null;
}

function saveLeadChanges() {
    if (!activeLeadKey) return;

    const newStatus = document.getElementById('lead-status-select').value;
    const newDate = document.getElementById('lead-next-touch').value;
    const newNotes = document.getElementById('lead-notes').value;
    const newR1Done = document.getElementById('lead-r1-done').checked;

    const prevState = crmData[activeLeadKey] || { status: 'A Contatar', r1Done: false };

    crmData[activeLeadKey] = {
        status: newStatus,
        date: newDate,
        notes: newNotes,
        r1Done: newR1Done
    };

    saveToLocalStorage();

    // Regista log de alteração automática se houver mudança de status
    if (prevState.status !== newStatus) {
        logActivityEvent(newStatus);
    }

    // Regista reunião realizada se marcou o checkbox hoje
    if (newR1Done && !prevState.r1Done) {
        const todayStr = getLocalDateString();
        ensureActivityDate(todayStr);
        crmActivity[todayStr].reuniao_feita++;
        saveActivityToLocalStorage();
    }

    closeLeadDrawer();
    renderAll();
}

function switchTab(tabId) {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.dataset.tab === tabId) btn.classList.add('active');
    });

    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(tabId).classList.add('active');
}

function generateDynamicCopys(key, data) {
    const isCursinho = data.nicho === 'cursinho' || data.nomeProspect.toLowerCase().includes('cursinho') || data.nomeProspect.toLowerCase().includes('vestibulares');
    
    const bolha1 = isCursinho 
        ? "Oi! Tudo certo? Aqui é o Kaio, da Legado. A gente ajuda cursinho e escola a se posicionar melhor pra atrair mais aluno."
        : "Oi! Tudo certo? Aqui é o Kaio, da Legado. A gente ajuda escola de idioma a se posicionar melhor pra atrair mais aluno.";
    
    const elogioText = data.elogio || ('o posicionamento de vocês em ' + data.cidade);
    const bolha2 = isCursinho
        ? 'Vi o trabalho da ' + data.nomeProspect + ' aí de ' + data.cidade + ' e curti ' + elogioText + '.'
        : 'Acompanho a ' + data.nomeProspect + ' no Insta e curto ' + elogioText + '.';
    
    const ganchoText = data.gancho || 'o trabalho e infraestrutura de vocês, mas o feed esconde a qualidade real do ensino';
    const bolha3 = 'Aí teve uma coisa que me deixou inquieto: ' + ganchoText + '. Resolvi destrinchar.';
    
    const targetLink = 'https://analise-legado.vercel.app/?p=' + key;
    const bolha4 = 'Preparei essa análise rápida do posicionamento de vocês, de graça: ' + targetLink;
    const bolha5 = "Vocês conseguem dar uma olhada aí?";

    const m1Text = bolha1 + '\n\n' + bolha2 + '\n\n' + bolha3 + '\n\n' + bolha4 + '\n\n' + bolha5;
    
    const fupText = 'Oi! Tudo bem? Passando só pra saber se conseguiu ver a análise rápida que te mandei da ' + data.nomeProspect + ' e se o que apontei fez sentido por aí. Conseguimos bater um papo rápido de 15 minutos essa semana? O que acha?';

    return { m1: m1Text, fup: fupText };
}

// Lançamento Manual no Diário
function saveManualLog() {
    const dateInput = document.getElementById('log-date').value;
    const type = document.getElementById('log-type').value;
    const qty = parseInt(document.getElementById('log-qty').value) || 0;

    if (!dateInput) {
        alert("Por favor, selecione uma data válida.");
        return;
    }

    ensureActivityDate(dateInput);

    if (type === 'touchpoints') {
        crmActivity[dateInput].touchpoints += qty;
    } else {
        crmActivity[dateInput][type] += qty;
        
        // Se incrementou Iniciados ou F.U, ajusta proporcionalmente o total de touchpoints
        if (type === 'iniciado' || type === 'followup') {
            crmActivity[dateInput].touchpoints += qty;
        }
    }

    saveActivityToLocalStorage();
    closeManualLogModal();
    renderAll();
}

function openManualLogModal() {
    document.getElementById('log-date').value = getLocalDateString();
    document.getElementById('log-qty').value = "1";
    document.getElementById('log-type').value = "iniciado";
    document.getElementById('manual-log-modal').classList.add('open');
}

function closeManualLogModal() {
    document.getElementById('manual-log-modal').classList.remove('open');
}

// Alterna abas principais de Analytics
function switchSectionTab(tabId) {
    document.querySelectorAll('.section-tab-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.dataset.sectionTab === tabId) btn.classList.add('active');
    });

    document.querySelectorAll('.section-tab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(tabId).classList.add('active');
}

// Event Listeners Setup
function setupEventListeners() {
    // Inputs de Busca e Filtros
    document.getElementById('search-input').addEventListener('input', renderKPIsAndTable);
    document.getElementById('filter-city').addEventListener('change', renderKPIsAndTable);
    document.getElementById('filter-niche').addEventListener('change', renderKPIsAndTable);
    document.getElementById('filter-status').addEventListener('change', renderKPIsAndTable);

    // Alternância de abas do topo (Analytics vs Diário)
    document.querySelectorAll('.section-tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            switchSectionTab(this.dataset.sectionTab);
        });
    });

    // Navegação do Diário
    document.getElementById('btn-prev-week').addEventListener('click', () => navigateWeek(-1));
    document.getElementById('btn-next-week').addEventListener('click', () => navigateWeek(1));

    // Modal de lançamento manual
    document.getElementById('btn-add-manual-log').addEventListener('click', openManualLogModal);
    document.querySelector('#manual-log-modal .modal-close-btn').addEventListener('click', closeManualLogModal);
    document.getElementById('modal-log-overlay').addEventListener('click', closeManualLogModal);
    document.getElementById('btn-save-log').addEventListener('click', saveManualLog);

    // Sorting Headers click
    document.querySelectorAll('.pipeline-table th.sortable').forEach(th => {
        th.addEventListener('click', function() {
            const sortField = this.dataset.sort;
            if (currentSortField === sortField) {
                currentSortOrder = currentSortOrder === 'asc' ? 'desc' : 'asc';
            } else {
                currentSortField = sortField;
                currentSortOrder = 'asc';
            }
            
            // Atualiza setas visuais
            document.querySelectorAll('.pipeline-table th.sortable i').forEach(icon => {
                icon.className = 'fa-solid fa-sort';
            });
            const activeIcon = this.querySelector('i');
            activeIcon.className = currentSortOrder === 'asc' ? 'fa-solid fa-sort-up' : 'fa-solid fa-sort-down';

            renderKPIsAndTable();
        });
    });

    // Drawer Tabs click
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            switchTab(this.dataset.tab);
        });
    });

    // Close overlays
    document.getElementById('drawer-close').addEventListener('click', closeLeadDrawer);
    document.getElementById('drawer-overlay').addEventListener('click', closeLeadDrawer);
    document.getElementById('btn-save-lead').addEventListener('click', saveLeadChanges);

    // Limpar / Reset dados local
    document.getElementById('btn-reset').addEventListener('click', () => {
        if (confirm("ATENÇÃO: Isso apagará todas as anotações e movimentações do crm. Deseja prosseguir?")) {
            localStorage.removeItem('legado_crm_data');
            localStorage.removeItem('legado_crm_activity');
            initData();
            renderAll();
        }
    });

    // Exportar dados JSON
    document.getElementById('btn-export').addEventListener('click', () => {
        const backupData = {
            crmData: crmData,
            crmActivity: crmActivity
        };
        const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(backupData, null, 2));
        const downloadAnchor = document.createElement('a');
        downloadAnchor.setAttribute("href", dataStr);
        downloadAnchor.setAttribute("download", 'Backup_CRM_Completo_Legado_' + new Date().toISOString().split('T')[0] + '.json');
        document.body.appendChild(downloadAnchor);
        downloadAnchor.click();
        downloadAnchor.remove();
    });

    // Importar backup JSON
    document.getElementById('btn-import').addEventListener('click', () => {
        document.getElementById('import-file').click();
    });

    document.getElementById('import-file').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(evt) {
            try {
                const parsed = JSON.parse(evt.target.result);
                if (typeof parsed === 'object') {
                    if (parsed.crmData && parsed.crmActivity) {
                        crmData = { ...crmData, ...parsed.crmData };
                        crmActivity = { ...crmActivity, ...parsed.crmActivity };
                    } else {
                        // Trata backup legado simples
                        crmData = { ...crmData, ...parsed };
                    }
                    saveToLocalStorage();
                    saveActivityToLocalStorage();
                    renderAll();
                    alert("Backup completo importado com sucesso!");
                } else {
                    alert("Formato inválido de arquivo.");
                }
            } catch (err) {
                alert("Erro ao ler o arquivo JSON.");
            }
        };
        reader.readAsText(file);
    });

    // Copiar scripts
    document.getElementById('btn-copy-m1').addEventListener('click', () => {
        const copyText = document.getElementById('copy-text-m1').textContent;
        navigator.clipboard.writeText(copyText).then(() => {
            const btn = document.getElementById('btn-copy-m1');
            btn.innerHTML = '<i class="fa-solid fa-check"></i> Copiado!';
            setTimeout(() => { btn.innerHTML = '<i class="fa-solid fa-copy"></i> Copiar Mensagem'; }, 2000);
        });
    });

    document.getElementById('btn-copy-fup').addEventListener('click', () => {
        const copyText = document.getElementById('copy-text-fup').textContent;
        navigator.clipboard.writeText(copyText).then(() => {
            const btn = document.getElementById('btn-copy-fup');
            btn.innerHTML = '<i class="fa-solid fa-check"></i> Copiado!';
            setTimeout(() => { btn.innerHTML = '<i class="fa-solid fa-copy"></i> Copiar Mensagem'; }, 2000);
        });
    });
}
