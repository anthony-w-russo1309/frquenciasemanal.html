<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Pódio de Frequência Escolar</title>
    <!-- Incluir a biblioteca SheetJS para ler arquivos Excel -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <style>
        :root {
            --bg-color: #222;
            --text-color: #fff;
            --panel-bg: #424242;
            --panel-text: #e0e0e0;
            --control-bg: #616161;
            --input-bg: #757575;
            --input-text: #fff;
            --border-color: #555;
            --button-primary: #2E7D32;
            --button-primary-hover: #1B5E20;
            --button-secondary: #2196F3;
            --button-secondary-hover: #0b7dda;
            --low-color: #ff5252;
            --medium-color: #ffd600;
            --high-color: #4CAF50;
            --cloud-color: #9C27B0;
            --cloud-hover: #7B1FA2;
            --excel-color: #217346;
            --excel-hover: #1a5a38;
        }

        body {
            font-family: 'Arial', sans-serif;
            margin: 0;
            padding: 0;
            background-color: var(--bg-color);
            color: var(--text-color);
            height: 100vh;
            width: 100vw;
            overflow: hidden;
        }
        
        .container {
            display: flex;
            flex-direction: column;
            height: 100vh;
            width: 100vw;
            aspect-ratio: 16/9;
            margin: 0 auto;
        }
        
        .display-screen {
            flex: 1;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            position: relative;
            overflow: hidden;
        }
        
        .slide {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            text-align: center;
            opacity: 0;
            transition: opacity 1s ease-in-out;
            padding: 2vh;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
        }
        
        .slide.active {
            opacity: 1;
        }
        
        .podium-container {
            display: flex;
            width: 100%;
            justify-content: center;
            align-items: flex-start;
        }
        
        .podium {
            display: flex;
            justify-content: center;
            align-items: flex-end;
            height: 60vh;
            width: 70%;
            margin-top: 2vh;
        }
        
        .legend {
            width: 25%;
            padding: 2vh;
            background-color: rgba(0, 0, 0, 0.5);
            border-radius: 8px;
            margin-left: 2vh;
            margin-top: 10vh;
            color: var(--text-color);
        }
        
        .legend-item {
            display: flex;
            align-items: center;
            margin-bottom: 1.5vh;
            font-size: 2.5vh;
        }
        
        .legend-color {
            width: 3vh;
            height: 3vh;
            border-radius: 50%;
            margin-right: 1.5vh;
        }
        
        .podium-step {
            display: flex;
            flex-direction: column;
            align-items: center;
            margin: 0 1.5vw;
        }
        
        .podium-stand {
            width: 18vw;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: flex-end;
            border-radius: 8px 8px 0 0;
            position: relative;
        }
        
        /* Cores do pódio baseadas em frequência */
        .high-podium {
            background-color: var(--high-color);
            box-shadow: 0 0 15px var(--high-color);
        }
        
        .medium-podium {
            background-color: var(--medium-color);
            box-shadow: 0 0 15px var(--medium-color);
        }
        
        .low-podium {
            background-color: var(--low-color);
            box-shadow: 0 0 15px var(--low-color);
        }
        
        .position-number {
            font-size: 3.5vh;
            font-weight: bold;
            margin-bottom: 1vh;
        }
        
        .class-info {
            background-color: rgba(0, 0, 0, 0.7);
            color: white;
            padding: 2vh;
            width: 100%;
            text-align: center;
            border-radius: 0 0 8px 8px;
        }
        
        .class-name {
            font-size: 4vh;
            font-weight: bold;
            margin-bottom: 1vh;
        }
        
        .attendance {
            font-size: 5vh;
            font-weight: bold;
            margin: 1vh 0;
        }
        
        .low {
            color: var(--low-color);
        }
        
        .medium {
            color: var(--medium-color);
        }
        
        .high {
            color: var(--high-color);
        }
        
        .variation {
            font-size: 3vh;
            margin-top: 1vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        
        .up {
            color: var(--high-color);
        }
        
        .down {
            color: var(--low-color);
        }
        
        .equal {
            color: var(--button-secondary);
        }
        
        .grade-title {
            font-size: 6vh;
            margin-bottom: 2vh;
            text-transform: uppercase;
            color: var(--text-color);
            text-shadow: 0 0 10px rgba(255, 255, 255, 0.5);
        }
        
        /* Estilo da tela de administração */
        .admin-panel {
            display: none;
            padding: 2vh;
            background-color: var(--panel-bg);
            color: var(--panel-text);
            max-width: 90vw;
            margin: 0 auto;
            width: 100%;
            height: 90vh;
            overflow-y: auto;
        }
        
        .admin-panel h2 {
            text-align: center;
            margin-bottom: 2vh;
            font-size: 4vh;
            color: var(--text-color);
        }
        
        .grade-section {
            margin-bottom: 3vh;
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 2vh;
        }
        
        .grade-section h3 {
            margin-bottom: 1.5vh;
            font-size: 3vh;
            color: var(--text-color);
        }
        
        .class-control {
            display: flex;
            align-items: center;
            margin-bottom: 1.5vh;
            padding: 1.5vh;
            background-color: var(--control-bg);
            border-radius: 5px;
        }
        
        .class-control label {
            width: 15vw;
            font-weight: bold;
            font-size: 2.5vh;
            color: var(--text-color);
        }
        
        .class-control input {
            width: 10vw;
            padding: 1vh;
            border: 1px solid var(--border-color);
            border-radius: 4px;
            text-align: center;
            font-size: 2.5vh;
            background-color: var(--input-bg);
            color: var(--input-text);
        }
        
        .class-control input:focus {
            outline: none;
            border-color: var(--high-color);
            background-color: var(--input-bg);
        }
        
        .buttons {
            display: flex;
            justify-content: center;
            margin-top: 2vh;
            gap: 1.5vh;
        }
        
        button {
            padding: 1.5vh 3vh;
            margin: 0 1.5vh;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-weight: bold;
            transition: all 0.3s;
            font-size: 2.5vh;
        }
        
        .save-btn {
            background-color: var(--button-primary);
            color: white;
        }
        
        .save-btn:hover {
            background-color: var(--button-primary-hover);
        }
        
        .cloud-btn {
            background-color: var(--cloud-color);
            color: white;
        }
        
        .cloud-btn:hover {
            background-color: var(--cloud-hover);
        }
        
        .excel-btn {
            background-color: var(--excel-color);
            color: white;
        }
        
        .excel-btn:hover {
            background-color: var(--excel-hover);
        }
        
        .toggle-btn {
            background-color: rgba(33, 150, 243, 0.2);
            color: white;
            position: fixed;
            bottom: 2vh;
            right: 2vh;
            z-index: 1000;
            border: 1px solid rgba(255, 255, 255, 0.3);
            opacity: 0.3;
            transition: all 0.3s ease;
            font-size: 2vh;
            padding: 1vh 2vh;
        }
        
        .toggle-btn:hover {
            background-color: rgba(11, 125, 218, 0.5);
            opacity: 1;
            box-shadow: 0 0 10px rgba(33, 150, 243, 0.5);
        }
        
        .week-label {
            font-size: 2vh;
            color: var(--panel-text);
            margin: 0 0.5vw;
        }
        
        .comparison-result {
            font-size: 2.5vh;
            margin-left: 1vw;
            padding: 0.5vh 1vw;
            border-radius: 4px;
            background-color: var(--control-bg);
            color: var(--text-color);
        }

        .theme-toggle {
            position: fixed;
            bottom: 2vh;
            left: 2vh;
            z-index: 1000;
            background-color: rgba(255, 255, 255, 0.1);
            color: white;
            border: none;
            border-radius: 50%;
            width: 5vh;
            height: 5vh;
            font-size: 2.5vh;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            opacity: 0.5;
            transition: all 0.3s;
        }

        .theme-toggle:hover {
            opacity: 1;
            background-color: rgba(255, 255, 255, 0.2);
        }
        
        .config-panel {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background-color: var(--panel-bg);
            padding: 3vh;
            border-radius: 8px;
            z-index: 2000;
            box-shadow: 0 0 20px rgba(0, 0, 0, 0.5);
            display: none;
            width: 80%;
            max-width: 500px;
        }
        
        .config-panel h3 {
            margin-top: 0;
            text-align: center;
            color: var(--text-color);
        }
        
        .config-input {
            margin-bottom: 2vh;
        }
        
        .config-input label {
            display: block;
            margin-bottom: 1vh;
            font-weight: bold;
        }
        
        .config-input input {
            width: 100%;
            padding: 1.5vh;
            border: 1px solid var(--border-color);
            border-radius: 4px;
            background-color: var(--input-bg);
            color: var(--input-text);
            font-size: 2.5vh;
        }
        
        .config-buttons {
            display: flex;
            justify-content: space-between;
            margin-top: 3vh;
        }
        
        .overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0, 0, 0, 0.7);
            z-index: 1500;
            display: none;
        }
        
        .status-message {
            position: fixed;
            top: 20px;
            left: 50%;
            transform: translateX(-50%);
            padding: 10px 20px;
            border-radius: 5px;
            z-index: 2000;
            font-weight: bold;
            display: none;
        }
        
        .status-success {
            background-color: var(--high-color);
            color: white;
        }
        
        .status-error {
            background-color: var(--low-color);
            color: white;
        }
        
        /* Estilo para o input de arquivo oculto */
        #excelFileInput {
            display: none;
        }
        
        /* Estilo para a mensagem de carregamento do Excel */
        .excel-status {
            margin-top: 10px;
            padding: 10px;
            border-radius: 4px;
            text-align: center;
            display: none;
        }
        
        .excel-loading {
            background-color: var(--medium-color);
            color: #333;
        }
        
        .excel-success {
            background-color: var(--high-color);
            color: white;
        }
        
        .excel-error {
            background-color: var(--low-color);
            color: white;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="display-screen" id="displayScreen"></div>
        
        <div class="admin-panel" id="adminPanel">
            <h2>Painel de Controle - Frequência das Turmas</h2>
        
            <div id="gradeControls"></div>
            
            <div class="buttons">
                <button class="save-btn" id="saveData">Salvar Local</button>
                <button class="cloud-btn" id="saveCloud">Salvar na Nuvem</button>
                <button class="cloud-btn" id="loadCloud">Carregar da Nuvem</button>
                <button class="excel-btn" id="importExcel">Importar do Excel</button>
                <button class="excel-btn" id="exportExcel">Exportar para Excel</button>
                <button class="cloud-btn" id="configCloud">Configurar</button>
            </div>
            
            <!-- Mensagem de status para importação do Excel -->
            <div class="excel-status" id="excelStatus"></div>
        </div>
        
        <button class="toggle-btn" id="togglePanel">Admin</button>
        <button class="theme-toggle" id="themeToggle">🌓</button>
    </div>
    
    <!-- Input oculto para seleção de arquivo -->
    <input type="file" id="excelFileInput" accept=".xlsx, .xls">
    
    <!-- Painel de configuração da nuvem -->
    <div class="overlay" id="configOverlay"></div>
    <div class="config-panel" id="configPanel">
        <h3>Configuração do JSONbin.io</h3>
        <div class="config-input">
            <label for="apiKey">Chave API (X-Master-Key):</label>
            <input type="password" id="apiKey" placeholder="Cole sua chave API aqui">
        </div>
        <div class="config-input">
            <label for="binId">ID do Bin (opcional):</label>
            <input type="text" id="binId" placeholder="Deixe em branco para criar um novo">
        </div>
        <div class="config-buttons">
            <button class="save-btn" id="saveConfig">Salvar</button>
            <button class="cloud-btn" id="cancelConfig">Cancelar</button>
        </div>
    </div>
    
    <!-- Mensagens de status -->
    <div class="status-message" id="statusMessage"></div>

    <script>
        // Dados das turmas (com campos para semana atual e anterior)
        let classesData = {
            "6º ano EF": {
                "A": { 
                    currentWeek: 95.5, 
                    previousWeek: 94.2 
                },
                "B": { 
                    currentWeek: 92.3, 
                    previousWeek: 91.7 
                },
                "C": { 
                    currentWeek: 90.8, 
                    previousWeek: 89.5 
                }
            },
            "7º ano EF": {
                "A": { 
                    currentWeek: 94.1, 
                    previousWeek: 93.0 
                },
                "B": { 
                    currentWeek: 91.4, 
                    previousWeek: 90.2 
                },
                "C": { 
                    currentWeek: 89.9, 
                    previousWeek: 88.3 
                }
            },
            "8º ano EF": {
                "A": { 
                    currentWeek: 93.7, 
                    previousWeek: 92.5 
                },
                "B": { 
                    currentWeek: 90.0, 
                    previousWeek: 89.1 
                },
                "C": { 
                    currentWeek: 88.6, 
                    previousWeek: 87.2 
                }
            },
            "9º ano EF": {
                "A": { 
                    currentWeek: 96.2, 
                    previousWeek: 95.0 
                },
                "B": { 
                    currentWeek: 93.4, 
                    previousWeek: 92.8 
                },
                "C": { 
                    currentWeek: 91.7, 
                    previousWeek: 90.5 
                }
            },
            "1º ano EM": {
                "A": { 
                    currentWeek: 97.3, 
                    previousWeek: 96.1 
                },
                "B": { 
                    currentWeek: 94.5, 
                    previousWeek: 93.8 
                },
                "C": { 
                    currentWeek: 92.9, 
                    previousWeek: 91.7 
                }
            },
            "2º ano EM": {
                "A": { 
                    currentWeek: 95.8, 
                    previousWeek: 94.6 
                },
                "B": { 
                    currentWeek: 93.2, 
                    previousWeek: 92.4 
                },
                "C": { 
                    currentWeek: 90.1, 
                    previousWeek: 89.3 
                }
            },
            "3º ano EM": {
                "A": { 
                    currentWeek: 98.4, 
                    previousWeek: 97.2 
                },
                "C": { 
                    currentWeek: 95.7, 
                    previousWeek: 94.9 
                },
                "D": { 
                    currentWeek: 93.5, 
                    previousWeek: 92.8 
                }
            }
        };

        // Configuração do JSONbin
        let jsonBinConfig = {
            apiKey: '',
            binId: '',
            baseURL: 'https://api.jsonbin.io/v3/b'
        };

        // Elementos da DOM
        const displayScreen = document.getElementById('displayScreen');
        const adminPanel = document.getElementById('adminPanel');
        const gradeControls = document.getElementById('gradeControls');
        const togglePanelBtn = document.getElementById('togglePanel');
        const saveDataBtn = document.getElementById('saveData');
        const saveCloudBtn = document.getElementById('saveCloud');
        const loadCloudBtn = document.getElementById('loadCloud');
        const importExcelBtn = document.getElementById('importExcel');
        const exportExcelBtn = document.getElementById('exportExcel');
        const configCloudBtn = document.getElementById('configCloud');
        const themeToggle = document.getElementById('themeToggle');
        const configPanel = document.getElementById('configPanel');
        const configOverlay = document.getElementById('configOverlay');
        const apiKeyInput = document.getElementById('apiKey');
        const binIdInput = document.getElementById('binId');
        const saveConfigBtn = document.getElementById('saveConfig');
        const cancelConfigBtn = document.getElementById('cancelConfig');
        const statusMessage = document.getElementById('statusMessage');
        const excelFileInput = document.getElementById('excelFileInput');
        const excelStatus = document.getElementById('excelStatus');

        // Variáveis de controle
        let currentSlide = 0;
        let slideInterval;
        const slideDuration = 5000; // 5 segundos por slide
        let darkMode = true;

        // Inicializar a aplicação
        function init() {
            loadSavedData();
            loadCloudConfig();
            createSlides();
            createAdminControls();
            startSlideShow();
            applyTheme();
        }

        // Aplicar tema
        function applyTheme() {
            document.documentElement.style.setProperty('--bg-color', darkMode ? '#222' : '#f5f5f5');
            document.documentElement.style.setProperty('--text-color', darkMode ? '#fff' : '#333');
            document.documentElement.style.setProperty('--panel-bg', darkMode ? '#424242' : '#ffffff');
            document.documentElement.style.setProperty('--panel-text', darkMode ? '#e0e0e0' : '#333');
            document.documentElement.style.setProperty('--control-bg', darkMode ? '#616161' : '#f9f9f9');
            document.documentElement.style.setProperty('--input-bg', darkMode ? '#757575' : '#ffffff');
            document.documentElement.style.setProperty('--input-text', darkMode ? '#fff' : '#000');
            document.documentElement.style.setProperty('--border-color', darkMode ? '#555' : '#ddd');
            themeToggle.textContent = darkMode ? '🌓' : '🌒';
        }

        // Alternar tema
        function toggleTheme() {
            darkMode = !darkMode;
            applyTheme();
            localStorage.setItem('darkMode', darkMode ? 'enabled' : 'disabled');
        }

        // Carregar dados salvos
        function loadSavedData() {
            const savedData = localStorage.getItem('attendanceData');
            if (savedData) {
                try {
                    const parsedData = JSON.parse(savedData);
                    if (parsedData && typeof parsedData === 'object') {
                        for (const grade in parsedData) {
                            if (classesData[grade]) {
                                for (const className in parsedData[grade]) {
                                    if (classesData[grade][className]) {
                                        classesData[grade][className] = parsedData[grade][className];
                                    }
                                }
                            }
                        }
                    }
                } catch (e) {
                    console.error('Erro ao carregar dados salvos:', e);
                }
            }

            // Carregar preferência de tema
            const savedTheme = localStorage.getItem('darkMode');
            if (savedTheme === 'disabled') {
                darkMode = false;
            }
        }

        // Carregar configuração da nuvem
        function loadCloudConfig() {
            const savedConfig = localStorage.getItem('jsonBinConfig');
            if (savedConfig) {
                try {
                    jsonBinConfig = JSON.parse(savedConfig);
                } catch (e) {
                    console.error('Erro ao carregar configuração da nuvem:', e);
                }
            }
        }

        // Salvar configuração da nuvem
        function saveCloudConfig() {
            localStorage.setItem('jsonBinConfig', JSON.stringify(jsonBinConfig));
        }

        // Mostrar mensagem de status
        function showStatus(message, isSuccess = true) {
            statusMessage.textContent = message;
            statusMessage.className = isSuccess ? 'status-message status-success' : 'status-message status-error';
            statusMessage.style.display = 'block';
            
            setTimeout(() => {
                statusMessage.style.display = 'none';
            }, 3000);
        }

        // Gerenciador do JSONbin
        class JSONBinManager {
            constructor(config) {
                this.config = config;
            }

            // Criar um novo bin
            async createBin(data, binName = 'Frequência Escolar') {
                try {
                    const response = await fetch(this.config.baseURL, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'X-Master-Key': this.config.apiKey,
                            'X-Bin-Name': binName
                        },
                        body: JSON.stringify(data)
                    });

                    if (!response.ok) throw new Error('Erro ao criar bin');

                    const result = await response.json();
                    this.config.binId = result.metadata.id;
                    saveCloudConfig();
                    return result;
                } catch (error) {
                    console.error('Erro:', error);
                    throw error;
                }
            }

            // Ler dados do bin
            async readBin() {
                if (!this.config.binId) throw new Error('Bin ID não definido');

                try {
                    const response = await fetch(`${this.config.baseURL}/${this.config.binId}`, {
                        method: 'GET',
                        headers: {
                            'X-Master-Key': this.config.apiKey
                        }
                    });

                    if (!response.ok) throw new Error('Erro ao ler bin');

                    const result = await response.json();
                    return result.record;
                } catch (error) {
                    console.error('Erro:', error);
                    throw error;
                }
            }

            // Atualizar dados do bin
            async updateBin(data) {
                if (!this.config.binId) throw new Error('Bin ID não definido');

                try {
                    const response = await fetch(`${this.config.baseURL}/${this.config.binId}`, {
                        method: 'PUT',
                        headers: {
                            'Content-Type': 'application/json',
                            'X-Master-Key': this.config.apiKey
                        },
                        body: JSON.stringify(data)
                    });

                    if (!response.ok) throw new Error('Erro ao atualizar bin');

                    return await response.json();
                } catch (error) {
                    console.error('Erro:', error);
                    throw error;
                }
            }
        }

        // Salvar dados na nuvem
        async function saveToCloud() {
            if (!jsonBinConfig.apiKey) {
                showStatus('Configure primeiro a chave API', false);
                showConfigPanel();
                return;
            }

            saveCloudBtn.textContent = 'Salvando...';
            saveCloudBtn.disabled = true;

            try {
                const jsonBinManager = new JSONBinManager(jsonBinConfig);
                
                // Atualizar dados dos inputs antes de salvar
                updateDataFromInputs();
                
                if (jsonBinConfig.binId) {
                    await jsonBinManager.updateBin(classesData);
                    showStatus('Dados atualizados na nuvem com sucesso!');
                } else {
                    const result = await jsonBinManager.createBin(classesData);
                    jsonBinConfig.binId = result.metadata.id;
                    saveCloudConfig();
                    showStatus('Dados salvos na nuvem com sucesso! Novo bin criado: ' + result.metadata.id);
                }
            } catch (error) {
                console.error('Erro ao salvar na nuvem:', error);
                showStatus('Erro ao salvar na nuvem: ' + error.message, false);
            } finally {
                saveCloudBtn.textContent = 'Salvar na Nuvem';
                saveCloudBtn.disabled = false;
            }
        }

        // Carregar dados da nuvem
        async function loadFromCloud() {
            if (!jsonBinConfig.apiKey || !jsonBinConfig.binId) {
                showStatus('Configure primeiro a chave API e o ID do bin', false);
                showConfigPanel();
                return;
            }

            loadCloudBtn.textContent = 'Carregando...';
            loadCloudBtn.disabled = true;

            try {
                const jsonBinManager = new JSONBinManager(jsonBinConfig);
                const cloudData = await jsonBinManager.readBin();
                
                // Atualizar os dados locais
                classesData = cloudData;
                
                // Atualizar a interface
                createAdminControls();
                updateSlides();
                
                // Salvar localmente também
                localStorage.setItem('attendanceData', JSON.stringify(classesData));
                
                showStatus('Dados carregados da nuvem com sucesso!');
            } catch (error) {
                console.error('Erro ao carregar da nuvem:', error);
                showStatus('Erro ao carregar da nuvem: ' + error.message, false);
            } finally {
                loadCloudBtn.textContent = 'Carregar da Nuvem';
                loadCloudBtn.disabled = false;
            }
        }

        // Mostrar painel de configuração
        function showConfigPanel() {
            apiKeyInput.value = jsonBinConfig.apiKey || '';
            binIdInput.value = jsonBinConfig.binId || '';
            configPanel.style.display = 'block';
            configOverlay.style.display = 'block';
        }

        // Fechar painel de configuração
        function hideConfigPanel() {
            configPanel.style.display = 'none';
            configOverlay.style.display = 'none';
        }

        // Salvar configuração
        function saveConfig() {
            jsonBinConfig.apiKey = apiKeyInput.value.trim();
            jsonBinConfig.binId = binIdInput.value.trim();
            saveCloudConfig();
            hideConfigPanel();
            showStatus('Configuração salva com sucesso!');
        }

        // Determinar a classe de cor com base na frequência
        function getAttendanceClass(attendance) {
            if (attendance < 85) return 'low';
            if (attendance < 90) return 'medium';
            return 'high';
        }

        // Criar slides para cada série - FUNÇÃO MODIFICADA
        function createSlides() {
            displayScreen.innerHTML = '';
            
            for (const grade in classesData) {
                const slide = document.createElement('div');
                slide.className = 'slide';
                
                // Ordenar turmas por frequência (maior para menor)
                const sortedClasses = Object.entries(classesData[grade])
                    .sort((a, b) => b[1].currentWeek - a[1].currentWeek);
                
                // Criar título da série
                const title = document.createElement('h1');
                title.className = 'grade-title';
                title.textContent = grade;
                slide.appendChild(title);
                
                // Criar container do pódio e legenda
                const podiumContainer = document.createElement('div');
                podiumContainer.className = 'podium-container';
                
                // Criar pódio
                const podium = document.createElement('div');
                podium.className = 'podium';
                
                // Calcular alturas proporcionais
                const maxFreq = sortedClasses[0][1].currentWeek;
                const minFreq = sortedClasses[sortedClasses.length-1][1].currentWeek;
                const maxHeight = 50; // Altura máxima em vh
                const minHeight = 30; // Altura mínima em vh
                
                // Função para calcular altura proporcional
                const calculateHeight = (freq) => {
                    if (maxFreq === minFreq) return maxHeight;
                    return minHeight + ((freq - minFreq) / (maxFreq - minFreq)) * (maxHeight - minHeight);
                };
                
                // Criar os degraus do pódio na ordem correta: 2º (esquerda), 1º (centro), 3º (direita)
                if (sortedClasses.length >= 2) {
                    const height = calculateHeight(sortedClasses[1][1].currentWeek);
                    podium.appendChild(createPodiumStep(sortedClasses[1], '2º', height));
                }
                
                if (sortedClasses.length >= 1) {
                    const height = calculateHeight(sortedClasses[0][1].currentWeek);
                    podium.appendChild(createPodiumStep(sortedClasses[0], '1º', height));
                }
                
                if (sortedClasses.length >= 3) {
                    const height = calculateHeight(sortedClasses[2][1].currentWeek);
                    podium.appendChild(createPodiumStep(sortedClasses[2], '3º', height));
                }
                
                // Criar legenda
                const legend = document.createElement('div');
                legend.className = 'legend';
                
                const legendTitle = document.createElement('h3');
                legendTitle.textContent = 'Legenda de Frequência';
                legend.appendChild(legendTitle);
                
                // Item 1: Alta frequência (verde)
                const highItem = document.createElement('div');
                highItem.className = 'legend-item';
                const highColor = document.createElement('div');
                highColor.className = 'legend-color';
                highColor.style.backgroundColor = '#4CAF50';
                highItem.appendChild(highColor);
                highItem.appendChild(document.createTextNode('90% ou mais'));
                legend.appendChild(highItem);
                
                // Item 2: Média frequência (amarelo)
                const mediumItem = document.createElement('div');
                mediumItem.className = 'legend-item';
                const mediumColor = document.createElement('div');
                mediumColor.className = 'legend-color';
                mediumColor.style.backgroundColor = '#FFD600';
                mediumItem.appendChild(mediumColor);
                mediumItem.appendChild(document.createTextNode('85% a 89,9%'));
                legend.appendChild(mediumItem);
                
                // Item 3: Baixa frequência (vermelho)
                const lowItem = document.createElement('div');
                lowItem.className = 'legend-item';
                const lowColor = document.createElement('div');
                lowColor.className = 'legend-color';
                lowColor.style.backgroundColor = '#FF5252';
                lowItem.appendChild(lowColor);
                lowItem.appendChild(document.createTextNode('Abaixo de 85%'));
                legend.appendChild(lowItem);
                
                podiumContainer.appendChild(podium);
                podiumContainer.appendChild(legend);
                slide.appendChild(podiumContainer);
                displayScreen.appendChild(slide);
            }
            
            // Ativar o primeiro slide
            const slides = document.querySelectorAll('.slide');
            if (slides.length > 0) {
                slides[0].classList.add('active');
                currentSlide = 0;
            }
        }

        // Criar um degrau do pódio com cores baseadas na frequência - FUNÇÃO MODIFICADA
        function createPodiumStep(classData, position, height) {
            const podiumStep = document.createElement('div');
            podiumStep.className = 'podium-step';
            
            const positionNumber = document.createElement('div');
            positionNumber.className = 'position-number';
            positionNumber.textContent = position;
            podiumStep.appendChild(positionNumber);
            
            const podiumStand = document.createElement('div');
            
            // Determinar a classe de cor com base na frequência
            const attendance = classData[1].currentWeek;
            let podiumClass, positionText;
            
            if (attendance >= 90) {
                podiumClass = 'high-podium';
                positionText = 'ALTA FREQUÊNCIA';
            } else if (attendance >= 85) {
                podiumClass = 'medium-podium';
                positionText = 'MÉDIA FREQUÊNCIA';
            } else {
                podiumClass = 'low-podium';
                positionText = 'BAIXA FREQUÊNCIA';
            }
            
            // Definir altura proporcional
            podiumStand.style.height = `${height}vh`;
            podiumStand.className = `podium-stand ${podiumClass}`;
            
            const classInfo = document.createElement('div');
            classInfo.className = 'class-info';
            
            const className = document.createElement('div');
            className.className = 'class-name';
            className.textContent = `${classData[0]} - ${positionText}`;
            classInfo.appendChild(className);
            
            const attendanceDisplay = document.createElement('div');
            attendanceDisplay.className = 'attendance';
            attendanceDisplay.textContent = `${attendance.toFixed(1).replace('.', ',')}%`;
            classInfo.appendChild(attendanceDisplay);
            
            // Adicionar variação em relação à semana anterior
            const variation = document.createElement('div');
            variation.className = 'variation';
            
            const diff = classData[1].currentWeek - classData[1].previousWeek;
            const absDiff = Math.abs(diff).toFixed(1);
            
            if (diff > 0) {
                variation.innerHTML = `<span class="up">↑ ${absDiff}%</span> em relação à semana anterior`;
                variation.classList.add('up');
            } else if (diff < 0) {
                variation.innerHTML = `<span class="down">↓ ${absDiff}%</span> em relação à semana anterior`;
                variation.classList.add('down');
            } else {
                variation.innerHTML = `<span class="equal">→ ${absDiff}%</span> em relação à semana anterior`;
                variation.classList.add('equal');
            }
            
            classInfo.appendChild(variation);
            podiumStand.appendChild(classInfo);
            podiumStep.appendChild(podiumStand);
            
            return podiumStep;
        }

        // Atualizar os slides com os dados mais recentes
        function updateSlides() {
            createSlides();
            const slides = document.querySelectorAll('.slide');
            if (slides.length > 0) {
                slides[currentSlide % slides.length].classList.add('active');
            }
        }

        // Atualizar dados a partir dos inputs
        function updateDataFromInputs() {
            const inputs = document.querySelectorAll('.class-control input');
            
            inputs.forEach(input => {
                const grade = input.dataset.grade;
                const className = input.dataset.class;
                const weekType = input.dataset.week;
                let value = parseFloat(input.value) || 0;
                
                // Validação
                value = Math.max(0, Math.min(100, value));
                
                // Atualizar dados
                classesData[grade][className][weekType === 'current' ? 'currentWeek' : 'previousWeek'] = value;
            });
        }

        // Criar controles de administração com dois campos por turma
        function createAdminControls() {
            gradeControls.innerHTML = '';
            
            for (const grade in classesData) {
                const gradeSection = document.createElement('div');
                gradeSection.className = 'grade-section';
                
                const gradeTitle = document.createElement('h3');
                gradeTitle.textContent = grade;
                gradeSection.appendChild(gradeTitle);
                
                for (const className in classesData[grade]) {
                    const classControl = document.createElement('div');
                    classControl.className = 'class-control';
                    
                    const label = document.createElement('label');
                    label.textContent = `Turma ${className}:`;
                    classControl.appendChild(label);
                    
                    // Campo para semana atual
                    const currentLabel = document.createElement('span');
                    currentLabel.className = 'week-label';
                    currentLabel.textContent = 'Atual:';
                    classControl.appendChild(currentLabel);
                    
                    const currentInput = document.createElement('input');
                    currentInput.type = 'number';
                    currentInput.min = '0';
                    currentInput.max = '100';
                    currentInput.step = '0.1';
                    currentInput.value = classesData[grade][className].currentWeek;
                    currentInput.dataset.grade = grade;
                    currentInput.dataset.class = className;
                    currentInput.dataset.week = 'current';
                    currentInput.addEventListener('change', validateInput);
                    classControl.appendChild(currentInput);
                    
                    const currentPercent = document.createElement('span');
                    currentPercent.textContent = '%';
                    currentPercent.style.marginLeft = '5px';
                    classControl.appendChild(currentPercent);
                    
                    // Campo para semana anterior
                    const previousLabel = document.createElement('span');
                    previousLabel.className = 'week-label';
                    previousLabel.textContent = 'Anterior:';
                    previousLabel.style.marginLeft = '1vw';
                    classControl.appendChild(previousLabel);
                    
                    const previousInput = document.createElement('input');
                    previousInput.type = 'number';
                    previousInput.min = '0';
                    previousInput.max = '100';
                    previousInput.step = '0.1';
                    previousInput.value = classesData[grade][className].previousWeek;
                    previousInput.dataset.grade = grade;
                    previousInput.dataset.class = className;
                    previousInput.dataset.week = 'previous';
                    previousInput.addEventListener('change', validateInput);
                    classControl.appendChild(previousInput);
                    
                    const previousPercent = document.createElement('span');
                    previousPercent.textContent = '%';
                    previousPercent.style.marginLeft = '5px';
                    classControl.appendChild(previousPercent);
                    
                    // Resultado da comparação
                    const comparison = document.createElement('span');
                    comparison.className = 'comparison-result';
                    const diff = classesData[grade][className].currentWeek - classesData[grade][className].previousWeek;
                    const absDiff = Math.abs(diff).toFixed(1);
                    
                    if (diff > 0) {
                        comparison.innerHTML = `↑ +${absDiff}%`;
                        comparison.style.color = '#4CAF50';
                    } else if (diff < 0) {
                        comparison.innerHTML = `↓ ${absDiff}%`;
                        comparison.style.color = '#F44336';
                    } else {
                        comparison.innerHTML = `→ 0%`;
                        comparison.style.color = '#2196F3';
                    }
                    
                    classControl.appendChild(comparison);
                    
                    gradeSection.appendChild(classControl);
                }
                
                gradeControls.appendChild(gradeSection);
            }
        }

        // Validar input
        function validateInput() {
            let value = parseFloat(this.value);
            if (isNaN(value)) value = 0;
            this.value = Math.min(100, Math.max(0, value)).toFixed(1);
            
            // Atualizar resultado da comparação
            const grade = this.dataset.grade;
            const className = this.dataset.class;
            const weekType = this.dataset.week;
            
            // Encontrar o controle correspondente
            const control = this.parentElement;
            const comparison = control.querySelector('.comparison-result');
            
            // Obter valores atual e anterior
            let currentWeek, previousWeek;
            
            if (weekType === 'current') {
                currentWeek = parseFloat(control.querySelector('input[data-week="current"]').value);
                previousWeek = parseFloat(control.querySelector('input[data-week="previous"]').value);
            } else {
                previousWeek = parseFloat(this.value);
                currentWeek = parseFloat(control.querySelector('input[data-week="current"]').value);
            }
            
            // Calcular diferença
            const diff = currentWeek - previousWeek;
            const absDiff = Math.abs(diff).toFixed(1);
            
            // Atualizar exibição
            if (diff > 0) {
                comparison.innerHTML = `↑ +${absDiff}%`;
                comparison.style.color = '#4CAF50';
            } else if (diff < 0) {
                comparison.innerHTML = `↓ ${absDiff}%`;
                comparison.style.color = '#F44336';
            } else {
                comparison.innerHTML = `→ 0%`;
                comparison.style.color = '#2196F3';
            }
        }

        // Iniciar o slideshow
        function startSlideShow() {
            clearInterval(slideInterval);
            const slides = document.querySelectorAll('.slide');
            
            if (slides.length === 0) return;
            
            slideInterval = setInterval(() => {
                slides[currentSlide % slides.length].classList.remove('active');
                currentSlide = (currentSlide + 1) % slides.length;
                slides[currentSlide % slides.length].classList.add('active');
            }, slideDuration);
        }

        // Função para processar arquivo Excel
        function processExcelFile(file) {
            // Mostrar mensagem de carregamento
            excelStatus.textContent = 'Processando arquivo Excel...';
            excelStatus.className = 'excel-status excel-loading';
            excelStatus.style.display = 'block';
            
            const reader = new FileReader();
            
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // Processar a primeira planilha
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // Converter para JSON
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    // Processar os dados do Excel
                    processExcelData(jsonData);
                    
                    // Atualizar a interface
                    createAdminControls();
                    updateSlides();
                    
                    // Salvar dados localmente
                    localStorage.setItem('attendanceData', JSON.stringify(classesData));
                    
                    // Mostrar mensagem de sucesso
                    excelStatus.textContent = 'Dados do Excel importados com sucesso!';
                    excelStatus.className = 'excel-status excel-success';
                    showStatus('Dados do Excel importados com sucesso!');
                    
                } catch (error) {
                    console.error('Erro ao processar arquivo Excel:', error);
                    excelStatus.textContent = 'Erro ao processar arquivo: ' + error.message;
                    excelStatus.className = 'excel-status excel-error';
                    showStatus('Erro ao processar arquivo Excel: ' + error.message, false);
                }
            };
            
            reader.onerror = function() {
                excelStatus.textContent = 'Erro ao ler o arquivo.';
                excelStatus.className = 'excel-status excel-error';
                showStatus('Erro ao ler o arquivo.', false);
            };
            
            reader.readAsArrayBuffer(file);
        }

        // Função para processar os dados do Excel
        function processExcelData(data) {
            // Estrutura esperada do Excel:
            // [['Série', 'Turma', 'Frequência Atual', 'Frequência Anterior'], ...]
            
            // Pular cabeçalho (primeira linha) и processar as demais
            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                if (row && row.length >= 4) {
                    const grade = row[0] ? row[0].toString().trim() : '';
                    const className = row[1] ? row[1].toString().trim() : '';
                    const currentWeek = parseFloat(row[2]) || 0;
                    const previousWeek = parseFloat(row[3]) || 0;
                    
                    // Validar os dados
                    if (grade && className && !isNaN(currentWeek) && !isNaN(previousWeek)) {
                        // Verificar se a série existe, se não, criar
                        if (!classesData[grade]) {
                            classesData[grade] = {};
                        }
                        
                        // Adicionar/atualizar os dados da turma
                        classesData[grade][className] = {
                            currentWeek: Math.max(0, Math.min(100, currentWeek)),
                            previousWeek: Math.max(0, Math.min(100, previousWeek))
                        };
                    }
                }
            }
        }

        // Função para exportar dados para Excel
        function exportToExcel() {
            // Criar uma nova pasta de trabalho
            const wb = XLSX.utils.book_new();
            
            // Preparar os dados para exportação
            const excelData = [['Série', 'Turma', 'Frequência Atual', 'Frequência Anterior', 'Variação']];
            
            for (const grade in classesData) {
                for (const className in classesData[grade]) {
                    const current = classesData[grade][className].currentWeek;
                    const previous = classesData[grade][className].previousWeek;
                    const variation = current - previous;
                    
                    excelData.push([
                        grade,
                        className,
                        current,
                        previous,
                        variation
                    ]);
                }
            }
            
            // Criar uma planilha a partir dos dados
            const ws = XLSX.utils.aoa_to_sheet(excelData);
            
            // Adicionar a planilha à pasta de trabalho
            XLSX.utils.book_append_sheet(wb, ws, "Frequência Escolar");
            
            // Gerar o arquivo Excel e fazer o download
            XLSX.writeFile(wb, "frequencia_escolar.xlsx");
            
            showStatus('Dados exportados para Excel com sucesso!');
        }

        // Alternar entre tela de exibição e painel de admin
        togglePanelBtn.addEventListener('click', () => {
            if (adminPanel.style.display === 'block') {
                adminPanel.style.display = 'none';
                displayScreen.style.display = 'flex';
                togglePanelBtn.textContent = 'Admin';
                startSlideShow();
            } else {
                adminPanel.style.display = 'block';
                displayScreen.style.display = 'none';
                togglePanelBtn.textContent = 'Fechar';
                clearInterval(slideInterval);
            }
        });

        // Tornar o botão mais visível quando o mouse se aproxima
        document.addEventListener('mousemove', function(e) {
            const toggleBtn = document.getElementById('togglePanel');
            const rightEdge = window.innerWidth - 100;
            const bottomEdge = window.innerHeight - 50;
            
            if (e.clientX > rightEdge || e.clientY > bottomEdge) {
                toggleBtn.style.opacity = '0.8';
            } else {
                toggleBtn.style.opacity = '0.3';
            }
        });

        // Salvar dados localmente
        saveDataBtn.addEventListener('click', () => {
            saveDataBtn.textContent = 'Salvando...';
            saveDataBtn.disabled = true;
            
            updateDataFromInputs();
            
            // Salvar no localStorage
            localStorage.setItem('attendanceData', JSON.stringify(classesData));
            
            // Atualizar slides
            updateSlides();
            
            // Feedback visual
            setTimeout(() => {
                saveDataBtn.textContent = 'Dados Salvos!';
                setTimeout(() => {
                    saveDataBtn.textContent = 'Salvar Local';
                    saveDataBtn.disabled = false;
                }, 1000);
            }, 500);
        });

        // Event listeners para os novos botões
        saveCloudBtn.addEventListener('click', saveToCloud);
        loadCloudBtn.addEventListener('click', loadFromCloud);
        configCloudBtn.addEventListener('click', showConfigPanel);
        saveConfigBtn.addEventListener('click', saveConfig);
        cancelConfigBtn.addEventListener('click', hideConfigPanel);
        configOverlay.addEventListener('click', hideConfigPanel);

        // Alternar tema
        themeToggle.addEventListener('click', toggleTheme);

        // Event listener para o botão de importar Excel
        importExcelBtn.addEventListener('click', () => {
            excelFileInput.click();
        });

        // Event listener para quando um arquivo é selecionado
        excelFileInput.addEventListener('change', (event) => {
            const file = event.target.files[0];
            if (file) {
                processExcelFile(file);
            }
            // Limpar o input para permitir selecionar o mesmo arquivo novamente
            event.target.value = '';
        });

        // Event listener para o botão de exportar Excel
        exportExcelBtn.addEventListener('click', exportToExcel);

        // Iniciar a aplicação quando a página carregar
        window.addEventListener('DOMContentLoaded', init);
    </script>
</body>
</html>