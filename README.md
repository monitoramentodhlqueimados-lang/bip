<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Busca de Pacote - Coletor Zebra</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <!-- SheetJS (Leitor de Excel/CSV) -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen p-4">

    <div class="max-w-3xl mx-auto space-y-6">
        
        <!-- Cabeçalho -->
        <header class="bg-white p-6 rounded-xl shadow-md flex justify-between items-center">
            <div>
                <h1 class="text-2xl font-bold text-slate-900">Localizador de Pacotes</h1>
                <p class="text-sm text-slate-500">Consulta via Coletor Zebra & Planilha</p>
            </div>
            <div id="status-planilha" class="bg-amber-100 text-amber-800 px-3 py-1 rounded-full text-sm font-medium">
                Sem Planilha
            </div>
        </header>

        <!-- Passo 1: Upload da Planilha -->
        <section class="bg-white p-6 rounded-xl shadow-md space-y-3">
            <h2 class="text-lg font-semibold text-slate-700">1. Carregar Planilha de Estoque/Pacotes</h2>
            <input type="file" id="excel-file" accept=".xlsx, .xls, .csv" 
                class="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer border border-slate-200 rounded-lg p-1"/>
            <p class="text-xs text-slate-400">Colunas esperadas: <strong>[Col A: Código de Barras]</strong> | <strong>[Col B: Nome/Descrição do Pacote]</strong> | <strong>[Col C: Localização/Qtd (Opcional)]</strong></p>
        </section>

        <!-- Passo 2: Definir o Pacote Procurado -->
        <section class="bg-white p-6 rounded-xl shadow-md space-y-3">
            <h2 class="text-lg font-semibold text-slate-700">2. Qual pacote você está procurando?</h2>
            <div class="flex gap-2">
                <input type="text" id="target-input" placeholder="Digite ou bipe o código do pacote procurado..." 
                    class="flex-1 border border-slate-300 rounded-lg px-4 py-3 text-lg font-mono focus:ring-2 focus:ring-blue-500 focus:outline-none" />
                <button onclick="definirAlvo()" class="bg-blue-600 hover:bg-blue-700 text-white font-medium px-6 py-3 rounded-lg transition-colors">
                    Definir Alvo
                </button>
            </div>
            <div id="alvo-info" class="hidden p-3 bg-blue-50 border border-blue-200 rounded-lg text-blue-900">
                <!-- Informações do item procurado -->
            </div>
        </section>

        <!-- Passo 3: Área de Leitura com o Coletor Zebra -->
        <section class="bg-white p-6 rounded-xl shadow-md space-y-4">
            <h2 class="text-lg font-semibold text-slate-700">3. Escanear Itens com o Coletor</h2>
            <input type="text" id="scanner-input" placeholder="Bipe os pacotes aqui..." 
                class="w-full border-2 border-slate-400 rounded-lg px-4 py-4 text-xl font-mono focus:border-blue-600 focus:outline-none bg-slate-50" autocomplete="off" />
        </section>

        <!-- Card de Resultado da Consulta -->
        <section id="result-card" class="hidden p-6 rounded-xl shadow-lg transition-all duration-300 text-center space-y-2">
            <h3 id="result-title" class="text-2xl font-bold"></h3>
            <p id="result-code" class="text-lg font-mono"></p>
            <p id="result-desc" class="text-xl font-medium"></p>
            <p id="result-extra" class="text-sm opacity-80"></p>
        </section>

    </div>

    <script>
        let baseDados = new Map();
        let codigoAlvo = null;

        // Efeitos sonoros usando Web Audio API (sem dependência de arquivos externos de áudio)
        const audioCtx = new (window.AudioContext || window.webkitAudioContext)();

        function tocarSomSucesso() {
            const osc = audioCtx.createOscillator();
            const gain = audioCtx.createGain();
            osc.connect(gain);
            gain.connect(audioCtx.destination);
            osc.type = 'triangle';
            osc.frequency.setValueAtTime(523.25, audioCtx.currentTime); // C5
            osc.frequency.setValueAtTime(659.25, audioCtx.currentTime + 0.1); // E5
            gain.gain.setValueAtTime(0.3, audioCtx.currentTime);
            osc.start();
            osc.stop(audioCtx.currentTime + 0.3);
        }

        function tocarSomErro() {
            const osc = audioCtx.createOscillator();
            const gain = audioCtx.createGain();
            osc.connect(gain);
            gain.connect(audioCtx.destination);
            osc.type = 'sawtooth';
            osc.frequency.setValueAtTime(180, audioCtx.currentTime);
            gain.gain.setValueAtTime(0.3, audioCtx.currentTime);
            osc.start();
            osc.stop(audioCtx.currentTime + 0.25);
        }

        // Leitura do arquivo Excel/CSV
        document.getElementById('excel-file').addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                
                baseDados.clear();
                for (let i = 1; i < rows.length; i++) {
                    let col = rows[i];
                    let codigo = col[0] ? String(col[0]).trim() : '';
                    let descricao = col[1] ? String(col[1]).trim() : 'Sem descrição';
                    let extra = col[2] ? String(col[2]).trim() : '';

                    if (codigo) {
                        baseDados.set(codigo, { descricao, extra });
                    }
                }

                const badge = document.getElementById('status-planilha');
                badge.className = "bg-emerald-100 text-emerald-800 px-3 py-1 rounded-full text-sm font-medium";
                badge.innerText = `${baseDados.size} itens carregados`;

                document.getElementById('target-input').focus();
            };
            reader.readAsArrayBuffer(file);
        });

        // Definir item alvo
        document.getElementById('target-input').addEventListener('keydown', function(e) {
            if (e.key === 'Enter') definirAlvo();
        });

        function definirAlvo() {
            const input = document.getElementById('target-input');
            const valor = input.value.trim();

            if (!valor) return;

            codigoAlvo = valor;
            const item = baseDados.get(codigoAlvo);

            const alvoInfo = document.getElementById('alvo-info');
            alvoInfo.classList.remove('hidden');

            if (item) {
                alvoInfo.innerHTML = `<strong>Buscando por:</strong> ${item.descricao} <br><span class="text-xs font-mono">Código: ${codigoAlvo} ${item.extra ? '| ' + item.extra : ''}</span>`;
            } else {
                alvoInfo.innerHTML = `<strong>Buscando pelo código:</strong> <span class="font-mono">${codigoAlvo}</span> <br><span class="text-xs text-amber-700">(Nota: Este código não consta na planilha carregada)</span>`;
            }

            // Move o foco diretamente para a caixa de bip do Zebra
            const scanner = document.getElementById('scanner-input');
            scanner.focus();
            scanner.value = '';
        }

        // Bip do Coletor Zebra
        document.getElementById('scanner-input').addEventListener('keydown', function(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                processarBip();
            }
        });

        function processarBip() {
            const scanner = document.getElementById('scanner-input');
            const codigoLido = scanner.value.trim();
            scanner.value = '';

            if (!codigoLido) return;

            if (!codigoAlvo) {
                alert("Por favor, defina primeiro qual código/pacote você está procurando!");
                document.getElementById('target-input').focus();
                return;
            }

            const resultCard = document.getElementById('result-card');
            const title = document.getElementById('result-title');
            const code = document.getElementById('result-code');
            const desc = document.getElementById('result-desc');
            const extra = document.getElementById('result-extra');

            resultCard.classList.remove('hidden', 'bg-emerald-600', 'bg-amber-500', 'bg-rose-600', 'text-white', 'text-slate-900');

            const itemLido = baseDados.get(codigoLido);

            if (codigoLido === codigoAlvo) {
                // É O PACOTE PROCURADO!
                tocarSomSucesso();
                resultCard.classList.add('bg-emerald-600', 'text-white');
                title.innerText = "🎯 ENCONTRADO! É O PACOTE PROCURADO!";
                code.innerText = `Código: ${codigoLido}`;
                desc.innerText = itemLido ? itemLido.descricao : "Item confirmado";
                extra.innerText = itemLido && itemLido.extra ? `Detalhes: ${itemLido.extra}` : "";
            } else if (itemLido) {
                // É OUTRO ITEM VÁLIDO DA PLANILHA
                tocarSomErro();
                resultCard.classList.add('bg-amber-500', 'text-white');
                title.innerText = "⚠️ NÃO É O PACOTE PROCURADO";
                code.innerText = `Bipado: ${codigoLido}`;
                desc.innerText = `Item no bip: ${itemLido.descricao}`;
                extra.innerText = `Você ainda está procurando por: ${codigoAlvo}`;
            } else {
                // CÓDIGO NÃO CONSTA NA PLANILHA
                tocarSomErro();
                resultCard.classList.add('bg-rose-600', 'text-white');
                title.innerText = "❌ CÓDIGO NÃO ENCONTRADO NA PLANILHA";
                code.innerText = `Bipado: ${codigoLido}`;
                desc.innerText = "Este pacote não consta no inventario/planilha carregada.";
                extra.innerText = `Procurando por: ${codigoAlvo}`;
            }

            // Mantém o foco no campo para o próximo bip do Zebra
            scanner.focus();
        }
    </script>
</body>
</html>
