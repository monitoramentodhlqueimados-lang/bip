<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Conferência de Inventário - Coletor Zebra / Planilha</title>
    <!-- Tailwind CSS para estilização rápida e limpa -->
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <!-- Biblioteca SheetJS para leitura de arquivos Excel (.xlsx, .xls, .csv) -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen p-4">

    <div class="max-w-4xl mx-auto space-y-6">
        
        <!-- Cabeçalho -->
        <header class="bg-white p-6 rounded-xl shadow-md flex flex-col md:flex-row justify-between items-center gap-4">
            <div>
                <h1 class="text-2xl font-bold text-slate-900">Conferência de Inventario</h1>
                <p class="text-sm text-slate-500">Coletores Zebra / Upload de Planilha</p>
            </div>
            <div id="status-badge" class="bg-amber-100 text-amber-800 px-3 py-1 rounded-full text-sm font-medium">
                Nenhuma planilha carregada
            </div>
        </header>

        <!-- Passo 1: Upload da Planilha -->
        <section class="bg-white p-6 rounded-xl shadow-md space-y-4">
            <h2 class="text-lg font-semibold text-slate-700">1. Carregar Planilha de Controle</h2>
            <div class="flex flex-col sm:flex-row items-center gap-4">
                <input type="file" id="excel-file" accept=".xlsx, .xls, .csv" 
                    class="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer border border-slate-200 rounded-lg p-1"/>
            </div>
            <p class="text-xs text-slate-400">A planilha deve conter colunas equivalentes a: <strong>Código de Barras</strong>, <strong>Produto</strong> e <strong>Quantidade Esperada</strong>.</p>
        </section>

        <!-- Passo 2: Área de Input / Coletor Zebra -->
        <section class="bg-white p-6 rounded-xl shadow-md space-y-4">
            <h2 class="text-lg font-semibold text-slate-700">2. Escaneamento (Leitor Zebra ou Manual)</h2>
            <div class="flex gap-2">
                <input type="text" id="barcode-input" placeholder="Bipe o código de barras aqui..." 
                    class="flex-1 border border-slate-300 rounded-lg px-4 py-3 text-lg focus:outline-none focus:ring-2 focus:ring-blue-500" autofocus autocomplete="off" />
                <button onclick="processarScan()" class="bg-blue-600 hover:bg-blue-700 text-white font-medium px-6 py-3 rounded-lg transition-colors">
                    Registrar
                </button>
            </div>
        </section>

        <!-- Resumo de Estatísticas -->
        <div class="grid grid-cols-1 sm:grid-cols-3 gap-4">
            <div class="bg-white p-4 rounded-xl shadow-md text-center">
                <span class="text-slate-500 text-sm">Total na Planilha</span>
                <h3 id="stat-total" class="text-2xl font-bold text-slate-800">0</h3>
            </div>
            <div class="bg-white p-4 rounded-xl shadow-md text-center">
                <span class="text-slate-500 text-sm">Itens Conferidos</span>
                <h3 id="stat-lidos" class="text-2xl font-bold text-emerald-600">0</h3>
            </div>
            <div class="bg-white p-4 rounded-xl shadow-md text-center">
                <span class="text-slate-500 text-sm">Itens Faltantes</span>
                <h3 id="stat-faltantes" class="text-2xl font-bold text-rose-600">0</h3>
            </div>
        </div>

        <!-- Passo 3: Relatório de Divergências / Faltantes -->
        <section class="bg-white p-6 rounded-xl shadow-md space-y-4">
            <div class="flex justify-between items-center">
                <h2 class="text-lg font-semibold text-slate-700">3. Itens Faltantes / Não Localizados</h2>
                <button onclick="exportarFaltantes()" class="bg-slate-800 hover:bg-slate-900 text-white text-sm font-medium px-4 py-2 rounded-lg transition-colors">
                    Exportar Faltantes (CSV)
                </button>
            </div>
            
            <div class="overflow-x-auto border border-slate-200 rounded-lg">
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-slate-50 border-b border-slate-200 text-slate-600 text-sm">
                            <th class="p-3">Código de Barras</th>
                            <th class="p-3">Produto</th>
                            <th class="p-3">Qtd Esperada</th>
                            <th class="p-3">Qtd Lida</th>
                            <th class="p-3">Status</th>
                        </tr>
                    </thead>
                    <tbody id="tabela-faltantes" class="divide-y divide-slate-200 text-sm">
                        <tr>
                            <td colspan="5" class="p-4 text-center text-slate-400">Carregue uma planilha para iniciar a auditoria.</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </section>

    </div>

    <!-- Script de lógica da aplicação -->
    <script>
        let inventarioMap = new Map(); // Armazenará os dados da planilha (Chave: Código, Valor: {produto, qtdEsperada, qtdLida})

        // 1. Leitura do arquivo Excel/CSV
        document.getElementById('excel-file').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                
                // Pega a primeira aba da planilha
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Converte para JSON
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                
                processarDadosPlanilha(jsonData);
            };
            reader.readAsArrayBuffer(file);
        });

        function processarDadosPlanilha(rows) {
            if (rows.length < 2) {
                alert("A planilha parece estar vazia ou sem cabeçalho.");
                return;
            }

            // Assumimos o seguinte formato padrão nas colunas (ajuste os índices se necessário):
            // Coluna 0: Código de Barras
            // Coluna 1: Nome do Produto
            // Coluna 2: Quantidade Esperada
            inventarioMap.clear();

            for (let i = 1; i < rows.length; i++) {
                let row = rows[i];
                let codigo = row[0] ? String(row[0]).trim() : '';
                let produto = row[1] ? String(row[1]) : 'Desconhecido';
                let qtdEsperada = row[2] ? parseFloat(row[2]) : 1;

                if (codigo) {
                    inventarioMap.set(codigo, {
                        produto: produto,
                        qtdEsperada: qtdEsperada,
                        qtdLida: 0
                    });
                }
            }

            document.getElementById('status-badge').className = "bg-emerald-100 text-emerald-800 px-3 py-1 rounded-full text-sm font-medium";
            document.getElementById('status-badge').innerText = `${inventarioMap.size} itens carregados com sucesso!`;
            
            atualizarTela();
            document.getElementById('barcode-input').focus();
        }

        // 2. Captura do código inserido pelo Coletor Zebra (ou Enter manual)
        document.getElementById('barcode-input').addEventListener('keydown', function(event) {
            if (event.key === 'Enter') {
                event.preventDefault();
                processarScan();
            }
        });

        function processarScan() {
            const input = document.getElementById('barcode-input');
            const codigo = input.value.trim();
            
            if (!codigo) return;

            if (inventarioMap.has(codigo)) {
                let item = inventarioMap.get(codigo);
                item.qtdLida += 1;
            } else {
                // Caso queira registrar itens que não estavam na planilha mas foram bipados
                console.warn(`Código não encontrado na planilha: ${codigo}`);
                alert(`Atenção: O código ${codigo} não consta na planilha original!`);
            }

            input.value = ''; // Limpa o campo para o próximo bipe
            atualizarTela();
        }

        // 3. Atualização Dinâmica da Tabela e Estatísticas
        function atualizarTela() {
            let totalItens = inventarioMap.size;
            let totalLidos = 0;
            let totalFaltantes = 0;

            let tbody = document.getElementById('tabela-faltantes');
            tbody.innerHTML = '';

            if (totalItens === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="p-4 text-center text-slate-400">Nenhum dado disponível.</td></tr>`;
                return;
            }

            inventarioMap.forEach((item, codigo) => {
                let statusHTML = '';
                let isFaltante = item.qtdLida < item.qtdEsperada;

                if (item.qtdLida === 0) {
                    totalFaltantes++;
                    statusHTML = `<span class="bg-rose-100 text-rose-700 px-2.5 py-1 rounded-full text-xs font-semibold">Não Lido / Faltante</span>`;
                } else if (isFaltante) {
                    totalFaltantes++;
                    statusHTML = `<span class="bg-amber-100 text-amber-700 px-2.5 py-1 rounded-full text-xs font-semibold">Incompleto</span>`;
                } else {
                    totalLidos++;
                    statusHTML = `<span class="bg-emerald-100 text-emerald-700 px-2.5 py-1 rounded-full text-xs font-semibold">Completo</span>`;
                }

                // Renderiza apenas os pendentes/faltantes ou todos? Vamos listar todos com status, ou focar nos faltantes. 
                // Aqui listamos todos para controle visual completo:
                let tr = document.createElement('tr');
                tr.innerHTML = `
                    <td class="p-3 font-mono text-slate-600">${codigo}</td>
                    <td class="p-3 font-medium text-slate-800">${item.produto}</td>
                    <td class="p-3 text-slate-600">${item.qtdEsperada}</td>
                    <td class="p-3 text-slate-600 font-bold">${item.qtdLida}</td>
                    <td class="p-3">${statusHTML}</td>
                `;
                tbody.appendChild(tr);
            });

            document.getElementById('stat-total').innerText = totalItens;
            document.getElementById('stat-lidos').innerText = totalLidos;
            document.getElementById('stat-faltantes').innerText = totalFaltantes;
        }

        // 4. Exportar relatório de divergências para CSV
        function exportarFaltantes() {
            let csvContent = "data:text/csv;charset=utf-8,Codigo,Produto,Qtd_Esperada,Qtd_Lida,Status\n";

            inventarioMap.forEach((item, codigo) => {
                if (item.qtdLida < item.qtdEsperada) {
                    let status = item.qtdLida === 0 ? "Faltante" : "Incompleto";
                    csvContent += `"${codigo}","${item.produto}",${item.qtdEsperada},${item.qtdLida},"${status}"\n`;
                }
            });

            const encodedUri = encodeURI(csvContent);
            const link = document.createElement("a");
            link.setAttribute("href", encodedUri);
            link.setAttribute("download", "itens_faltantes_inventario.csv");
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    </script>
</body>
</html>
