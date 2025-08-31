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
        let darkMode = 