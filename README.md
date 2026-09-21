# simulador-grade[gemini-code-1790002758837.html](https://github.com/user-attachments/files/32475249/gemini-code-1790002758837.html)
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Simulador de Grade de Programação Interativo</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome CDN para Ícones -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS (XLSX Export) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.min.js"></script>
    <style>
        input:focus, select:focus {
            outline: 2px solid #2563eb;
            background-color: #eff6ff;
        }
        @media print {
            .no-print { display: none !important; }
            body { background: white; font-size: 12px; }
            .shadow-md, .shadow-lg { box-shadow: none !important; }
        }
    </style>
</head>
<body class="bg-slate-100 text-slate-800 font-sans min-h-screen pb-12">

    <!-- Header / Navbar -->
    <header class="bg-indigo-900 text-white shadow-lg no-print">
        <div class="max-w-7xl mx-auto px-4 py-6 flex flex-col md:flex-row justify-between items-center gap-4">
            <div>
                <h1 class="text-2xl font-bold flex items-center gap-2">
                    <i class="fa-solid me-2 fa-tv text-amber-400"></i>
                    Simulador de Grade de Programação
                </h1>
                <p class="text-indigo-200 text-sm mt-1">Ajuste blocos, comerciais e durações com recálculo em tempo real</p>
            </div>
            <div class="flex flex-wrap gap-2">
                <button onclick="addNewProgram()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm font-medium transition flex items-center gap-2 shadow">
                    <i class="fa-solid fa-plus"></i> Novo Programa
                </button>
                <button onclick="exportToExcel()" class="bg-green-700 hover:bg-green-800 text-white px-4 py-2 rounded-lg text-sm font-medium transition flex items-center gap-2 shadow">
                    <i class="fa-solid fa-file-excel"></i> Exportar Excel
                </button>
                <button onclick="window.print()" class="bg-slate-700 hover:bg-slate-800 text-white px-4 py-2 rounded-lg text-sm font-medium transition flex items-center gap-2 shadow">
                    <i class="fa-solid fa-print"></i> Imprimir / PDF
                </button>
                <button onclick="resetToDefault()" class="bg-rose-700 hover:bg-rose-800 text-white px-3 py-2 rounded-lg text-sm font-medium transition flex items-center gap-1 shadow" title="Restaurar padrão">
                    <i class="fa-solid fa-rotate-left"></i>
                </button>
            </div>
        </div>
    </header>

    <main class="max-w-7xl mx-auto px-4 mt-6">

        <!-- KPI Cards -->
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex items-center gap-4">
                <div class="w-12 h-12 bg-blue-100 text-blue-600 rounded-lg flex items-center justify-center text-xl font-bold">
                    <i class="fa-solid fa-film"></i>
                </div>
                <div>
                    <span class="text-xs uppercase tracking-wider text-slate-500 font-bold">Tempo de Conteúdo (Blocos)</span>
                    <h3 id="kpi-content" class="text-2xl font-extrabold text-slate-800">00:00:00</h3>
                </div>
            </div>

            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex items-center gap-4">
                <div class="w-12 h-12 bg-amber-100 text-amber-600 rounded-lg flex items-center justify-center text-xl font-bold">
                    <i class="fa-solid fa-rectangle-ad"></i>
                </div>
                <div>
                    <span class="text-xs uppercase tracking-wider text-slate-500 font-bold">Tempo de Comerciais</span>
                    <h3 id="kpi-commercial" class="text-2xl font-extrabold text-slate-800">00:00:00</h3>
                </div>
            </div>

            <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex items-center gap-4">
                <div class="w-12 h-12 bg-indigo-100 text-indigo-600 rounded-lg flex items-center justify-center text-xl font-bold">
                    <i class="fa-solid fa-clock"></i>
                </div>
                <div>
                    <span class="text-xs uppercase tracking-wider text-slate-500 font-bold">Tempo Total da Grade</span>
                    <h3 id="kpi-total" class="text-2xl font-extrabold text-indigo-900">00:00:00</h3>
                </div>
            </div>
        </div>

        <!-- Container da Grade -->
        <div id="schedule-container" class="space-y-8"></div>

    </main>

    <script>
        // Dados iniciais baseados no Simulador de Grade
        const defaultData = [
            {
                id: 'prog-1',
                title: 'JORNAL DA MANHÃ',
                slotTime: '01:00:00',
                startTime: '06:00:00',
                category: 'Jornalismo',
                items: [
                    { type: 'Bloco 1', desc: 'Jornal da Manhã - Bloco 1 (Notícias Gerais)', duration: '00:18:30', category: 'Jornalismo' },
                    { type: 'Comercial', desc: 'Intervalo Comercial 1', duration: '00:03:00', category: 'Comercial' },
                    { type: 'Bloco 2', desc: 'Jornal da Manhã - Bloco 2 (Esportes)', duration: '00:13:30', category: 'Jornalismo' },
                    { type: 'Comercial', desc: 'Intervalo Comercial 2', duration: '00:03:30', category: 'Comercial' },
                    { type: 'Bloco 3', desc: 'Jornal da Manhã - Bloco 3 (Política & Economia)', duration: '00:12:00', category: 'Jornalismo' },
                    { type: 'Comercial', desc: 'Intervalo Comercial 3', duration: '00:02:30', category: 'Comercial' },
                    { type: 'Bloco 4', desc: 'Jornal da Manhã - Bloco 4 (Previsão do Tempo)', duration: '00:08:00', category: 'Jornalismo' }
                ]
            },
            {
                id: 'prog-2',
                title: 'PROGRAMA DE VARIEDADES',
                slotTime: '01:00:00',
                startTime: '07:00:00',
                category: 'Entretenimento',
                items: [
                    { type: 'Bloco 1', desc: 'Programa de Variedades - Bloco 1 (Entrevista)', duration: '00:10:30', category: 'Entretenimento' },
                    { type: 'Comercial', desc: 'Intervalo Comercial 1', duration: '00:05:00', category: 'Comercial' },
                    { type: 'Bloco 2', desc: 'Programa de Variedades - Bloco 2 (Música ao Vivo)', duration: '00:25:00', category: 'Entretenimento' },
                    { type: 'Comercial', desc: 'Intervalo Comercial 2', duration: '00:05:00', category: 'Comercial' },
                    { type: 'Bloco 3', desc: 'Programa de Variedades - Bloco 3 (Culinária)', duration: '00:05:00', category: 'Entretenimento' },
                    { type: 'Comercial', desc: 'Intervalo Comercial 3', duration: '00:05:00', category: 'Comercial' },
                    { type: 'Bloco 4', desc: 'Programa de Variedades - Bloco 4 (Bastidores)', duration: '00:03:00', category: 'Entretenimento' }
                ]
            }
        ];

        let scheduleData = JSON.parse(localStorage.getItem('simulador_grade_data')) || defaultData;

        // Funções Utilitárias para Tratamento de Horários (hh:mm:ss)
        function timeToSeconds(tStr) {
            if (!tStr) return 0;
            const parts = tStr.split(':').map(Number);
            if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
            if (parts.length === 2) return parts[0] * 3600 + parts[1] * 60;
            return 0;
        }

        function secondsToTime(totalSec) {
            const isNegative = totalSec < 0;
            let absSec = Math.abs(totalSec);
            const h = Math.floor(absSec / 3600);
            const m = Math.floor((absSec % 3600) / 60);
            const s = absSec % 60;
            const formatted = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
            return isNegative ? `-${formatted}` : formatted;
        }

        function addSecondsToTime(timeStr, secondsToAdd) {
            let sec = timeToSeconds(timeStr) + secondsToAdd;
            // Wrap around 24h
            sec = ((sec % 86400) + 86400) % 86400;
            return secondsToTime(sec);
        }

        function saveData() {
            localStorage.setItem('simulador_grade_data', JSON.stringify(scheduleData));
            renderSchedule();
        }

        function resetToDefault() {
            if (confirm("Tem certeza que deseja restaurar a grade para o padrão inicial? Todas as alterações serão perdidas.")) {
                scheduleData = JSON.parse(JSON.stringify(defaultData));
                saveData();
            }
        }

        // Renderização da Interface
        function renderSchedule() {
            const container = document.getElementById('schedule-container');
            container.innerHTML = '';

            let globalStartSec = 0;
            let totalContentSec = 0;
            let totalCommercialSec = 0;
            let grandTotalSec = 0;

            scheduleData.forEach((prog, pIdx) => {
                // Cálculo de horários encadeados do programa
                if (pIdx === 0) {
                    globalStartSec = timeToSeconds(prog.startTime);
                }

                let currentProgStartSec = globalStartSec;
                let progTotalSec = 0;

                const progCard = document.createElement('div');
                progCard.className = 'bg-white rounded-xl shadow-md overflow-hidden border border-slate-200';

                // Cabeçalho do Programa
                let html = `
                    <div class="bg-indigo-900 text-white p-4 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                        <div class="flex-1 space-y-2 md:space-y-0 md:flex md:items-center md:gap-4">
                            <input type="text" value="${prog.title}" onchange="updateProgField(${pIdx}, 'title', this.value)" 
                                   class="bg-indigo-800 text-white text-lg font-bold px-3 py-1 rounded w-full md:w-auto border border-indigo-700" placeholder="Nome do Programa">
                            <div class="flex items-center gap-2 text-xs text-indigo-200">
                                <span>Slot Previsto:</span>
                                <input type="text" value="${prog.slotTime}" onchange="updateProgField(${pIdx}, 'slotTime', this.value)" 
                                       class="bg-indigo-800 text-white font-mono px-2 py-0.5 rounded w-20 text-center border border-indigo-700">
                            </div>
                        </div>
                        <div class="flex items-center gap-2 no-print">
                            <button onclick="addItem(${pIdx}, 'Bloco')" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1 rounded text-xs font-semibold flex items-center gap-1">
                                <i class="fa-solid fa-plus"></i> Bloco
                            </button>
                            <button onclick="addItem(${pIdx}, 'Comercial')" class="bg-amber-600 hover:bg-amber-700 text-white px-3 py-1 rounded text-xs font-semibold flex items-center gap-1">
                                <i class="fa-solid fa-plus"></i> Comercial
                            </button>
                            <button onclick="deleteProgram(${pIdx})" class="bg-rose-600 hover:bg-rose-700 text-white px-2 py-1 rounded text-xs" title="Excluir Programa">
                                <i class="fa-solid fa-trash"></i>
                            </button>
                        </div>
                    </div>

                    <div class="overflow-x-auto">
                        <table class="w-full text-left border-collapse">
                            <thead>
                                <tr class="bg-slate-100 text-slate-600 text-xs uppercase font-bold border-b border-slate-200">
                                    <th class="p-3 w-32">Elemento</th>
                                    <th class="p-3">Descrição do Programa / Comercial</th>
                                    <th class="p-3 w-36 text-center">Duração (hh:mm:ss)</th>
                                    <th class="p-3 w-36 text-center">Horário de Início</th>
                                    <th class="p-3 w-36">Categoria</th>
                                    <th class="p-3 w-16 text-center no-print">Ações</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-100 text-sm">
                `;

                // Linhas dos Itens (Blocos e Comerciais)
                let itemStartSec = currentProgStartSec;

                prog.items.forEach((item, iIdx) => {
                    const durSec = timeToSeconds(item.duration);
                    progTotalSec += durSec;
                    
                    if (item.type.toLowerCase().includes('comercial')) {
                        totalCommercialSec += durSec;
                    } else {
                        totalContentSec += durSec;
                    }

                    const startTimeStr = secondsToTime(itemStartSec);
                    itemStartSec += durSec;

                    const isComm = item.type.toLowerCase().includes('comercial');
                    const rowBg = isComm ? 'bg-amber-50/50' : 'bg-white';

                    html += `
                        <tr class="${rowBg} hover:bg-slate-50 transition">
                            <td class="p-2">
                                <input type="text" value="${item.type}" onchange="updateItemField(${pIdx}, ${iIdx}, 'type', this.value)" 
                                       class="w-full bg-transparent px-2 py-1 rounded border border-transparent hover:border-slate-300 font-semibold text-slate-700">
                            </td>
                            <td class="p-2">
                                <input type="text" value="${item.desc}" onchange="updateItemField(${pIdx}, ${iIdx}, 'desc', this.value)" 
                                       class="w-full bg-transparent px-2 py-1 rounded border border-transparent hover:border-slate-300">
                            </td>
                            <td class="p-2 text-center">
                                <input type="text" value="${item.duration}" onchange="updateItemField(${pIdx}, ${iIdx}, 'duration', this.value)" 
                                       class="w-28 text-center font-mono bg-white border border-slate-300 rounded px-2 py-1 focus:ring-2 focus:ring-indigo-500">
                            </td>
                            <td class="p-2 text-center font-mono text-slate-600 font-medium">
                                ${startTimeStr}
                            </td>
                            <td class="p-2">
                                <input type="text" value="${item.category}" onchange="updateItemField(${pIdx}, ${iIdx}, 'category', this.value)" 
                                       class="w-full bg-transparent px-2 py-1 rounded border border-transparent hover:border-slate-300 text-slate-500 text-xs">
                            </td>
                            <td class="p-2 text-center no-print">
                                <button onclick="deleteItem(${pIdx}, ${iIdx})" class="text-rose-500 hover:text-rose-700 p-1">
                                    <i class="fa-solid fa-xmark"></i>
                                </button>
                            </td>
                        </tr>
                    `;
                });

                // Cálculo do Status do Programa (Estouro / Buraco / OK)
                const slotSec = timeToSeconds(prog.slotTime);
                const diffSec = progTotalSec - slotSec;
                let statusBadge = '';

                if (diffSec === 0) {
                    statusBadge = `<span class="bg-emerald-100 text-emerald-800 px-3 py-1 rounded-full text-xs font-extrabold flex items-center gap-1"><i class="fa-solid fa-circle-check"></i> OK (GRADE EXATA)</span>`;
                } else if (diffSec > 0) {
                    statusBadge = `<span class="bg-rose-100 text-rose-800 px-3 py-1 rounded-full text-xs font-extrabold flex items-center gap-1"><i class="fa-solid fa-triangle-exclamation"></i> ESTOURO (+${secondsToTime(diffSec)})</span>`;
                } else {
                    statusBadge = `<span class="bg-amber-100 text-amber-800 px-3 py-1 rounded-full text-xs font-extrabold flex items-center gap-1"><i class="fa-solid fa-circle-exclamation"></i> BURACO (-${secondsToTime(Math.abs(diffSec))})</span>`;
                }

                // Linha de Subtotal do Programa
                html += `
                            </tbody>
                            <tfoot>
                                <tr class="bg-indigo-50/80 border-t border-indigo-200 font-semibold">
                                    <td class="p-3 font-bold text-indigo-900">SUBTOTAL</td>
                                    <td class="p-3 text-indigo-900">Duração Total Executada:</td>
                                    <td class="p-3 text-center font-mono font-bold text-indigo-900 text-base">${secondsToTime(progTotalSec)}</td>
                                    <td class="p-3 text-center font-mono text-indigo-700">${secondsToTime(currentProgStartSec)}</td>
                                    <td class="p-3" colspan="2">${statusBadge}</td>
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                `;

                progCard.innerHTML = html;
                container.appendChild(progCard);

                // Próximo programa inicia onde o atual termina
                globalStartSec = itemStartSec;
                grandTotalSec += progTotalSec;
            });

            // Atualiza os KPIs no topo
            document.getElementById('kpi-content').innerText = secondsToTime(totalContentSec);
            document.getElementById('kpi-commercial').innerText = secondsToTime(totalCommercialSec);
            document.getElementById('kpi-total').innerText = secondsToTime(grandTotalSec);
        }

        // Funções de Edição dos Dados
        function updateProgField(pIdx, field, val) {
            scheduleData[pIdx][field] = val;
            saveData();
        }

        function updateItemField(pIdx, iIdx, field, val) {
            scheduleData[pIdx].items[iIdx][field] = val;
            saveData();
        }

        function addItem(pIdx, type) {
            const count = scheduleData[pIdx].items.filter(i => i.type.includes(type)).length + 1;
            const newItem = {
                type: type === 'Bloco' ? `Bloco ${count}` : 'Comercial',
                desc: type === 'Bloco' ? `Novo Bloco ${count}` : `Intervalo Comercial ${count}`,
                duration: type === 'Bloco' ? '00:10:00' : '00:03:00',
                category: scheduleData[pIdx].category
            };
            scheduleData[pIdx].items.push(newItem);
            saveData();
        }

        function deleteItem(pIdx, iIdx) {
            scheduleData[pIdx].items.splice(iIdx, 1);
            saveData();
        }

        function addNewProgram() {
            const newProg = {
                id: 'prog-' + (scheduleData.length + 1),
                title: 'NOVO PROGRAMA',
                slotTime: '01:00:00',
                startTime: '08:00:00',
                category: 'Geral',
                items: [
                    { type: 'Bloco 1', desc: 'Bloco Inicial', duration: '00:25:00', category: 'Geral' },
                    { type: 'Comercial', desc: 'Intervalo Comercial 1', duration: '00:05:00', category: 'Comercial' }
                ]
            };
            scheduleData.push(newProg);
            saveData();
        }

        function deleteProgram(pIdx) {
            if (confirm(`Deseja realmente remover o programa "${scheduleData[pIdx].title}"?`)) {
                scheduleData.splice(pIdx, 1);
                saveData();
            }
        }

        // Exportação para Excel (.xlsx)
        function exportToExcel() {
            let exportRows = [];
            exportRows.push(["SIMULADOR DE GRADE DE PROGRAMAÇÃO - EXPORTAÇÃO"]);
            exportRows.push([]);
            exportRows.push(["Elemento", "Descrição", "Duração", "Horário de Início", "Categoria", "Status Grade"]);

            let globalSec = timeToSeconds(scheduleData[0]?.startTime || '06:00:00');

            scheduleData.forEach(prog => {
                exportRows.push([`PROGRAMA: ${prog.title}`, "", `SLOT PREVISTO: ${prog.slotTime}`, "", prog.category, ""]);
                let progSec = 0;
                
                prog.items.forEach(item => {
                    const durSec = timeToSeconds(item.duration);
                    const startStr = secondsToTime(globalSec);
                    exportRows.push([item.type, item.desc, item.duration, startStr, item.category, ""]);
                    globalSec += durSec;
                    progSec += durSec;
                });

                const slotSec = timeToSeconds(prog.slotTime);
                const diffSec = progSec - slotSec;
                let statusStr = "OK (GRADE EXATA)";
                if (diffSec > 0) statusStr = `ESTOURO (+${secondsToTime(diffSec)})`;
                if (diffSec < 0) statusStr = `BURACO (-${secondsToTime(Math.abs(diffSec))})`;

                exportRows.push(["SUBTOTAL", `Duração Executada (${prog.title})`, secondsToTime(progSec), "", "Status:", statusStr]);
                exportRows.push([]);
            });

            const ws = XLSX.utils.aoa_to_sheet(exportRows);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Grade_de_Programacao");
            XLSX.writeFile(wb, "Simulador_Grade_Programacao.xlsx");
        }

        // Inicializar aplicação
        renderSchedule();
    </script>
</body>
</html>
