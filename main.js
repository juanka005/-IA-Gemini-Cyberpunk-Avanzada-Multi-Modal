// ====================== MAIN.JS - MASTER OMEGA V15.2 ======================
// Núcleo de Inteligencia Cyber-Gen con Soporte de Archivos, Voz y Gráficos

const MODELS_LIST = ["gemini-2.5-flash", "gemini-1.5-pro"];
const SYSTEM_PROMPT = `Eres Chelsea-Bot V15.2. Un Analista de Datos Senior.
REGLAS ESTRICTAS:
1. Usa formato Markdown.
2. Si el usuario te pide un gráfico, devuelve EXCLUSIVAMENTE este formato exacto:
[CHART_DATA: { "type": "bar", "data": { "labels": ["A", "B"], "datasets": [{ "label": "Datos", "data": [10, 20] }] } }]
Asegúrate de que el JSON sea válido y de usar Chart.js.
3. Analiza archivos adjuntos de forma técnica y precisa.`;

let API_KEY = import.meta.env?.VITE_GEMINI_API_KEY || "";
let globalHistory = [];
let isAudioEnabled = true;
let extractedFileData = ""; // Memoria temporal para archivos

// --- REFERENCIAS DOM ---
const chatContainer = document.getElementById('chat-container');
const userInput = document.getElementById('user-input');
const sendBtn = document.getElementById('send-btn');
const fileInput = document.getElementById('file-input');
const fileIndicator = document.getElementById('file-indicator');
const voiceBtn = document.getElementById('voice-btn');
const toggleAudioBtn = document.getElementById('toggle-audio-btn');
const exportBtn = document.getElementById('export-btn');
const clearChatBtn = document.getElementById('clear-chat');
const welcomeScreen = document.getElementById('welcome-screen');
const sidebar = document.getElementById('sidebar');

// --- INICIALIZACIÓN ---
if (!API_KEY) {
    alert("⚠️ ALERTA: No se encontró VITE_GEMINI_API_KEY en el archivo .env");
}

// Configurar Worker de PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

// --- EVENT LISTENERS ---
sendBtn.addEventListener('click', handleSend);
userInput.addEventListener('keypress', (e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend(); } });

// Toggle Sidebar
document.getElementById('toggle-sidebar').addEventListener('click', () => sidebar.classList.remove('collapsed'));
document.getElementById('close-sidebar').addEventListener('click', () => sidebar.classList.add('collapsed'));

// Toggle Audio
toggleAudioBtn.addEventListener('click', () => {
    isAudioEnabled = !isAudioEnabled;
    toggleAudioBtn.innerHTML = isAudioEnabled ? '<i class="fas fa-volume-up"></i> Voz IA: ON' : '<i class="fas fa-volume-mute"></i> Voz IA: OFF';
});

// Limpiar Chat
clearChatBtn.addEventListener('click', () => {
    chatContainer.innerHTML = '';
    globalHistory = [];
    extractedFileData = "";
    fileIndicator.classList.add('d-none');
    chatContainer.appendChild(welcomeScreen);
});

// --- LÓGICA DE MENSAJERÍA ---
async function handleSend() {
    const text = userInput.value.trim();
    if (!text && !extractedFileData) return;

    if (welcomeScreen) welcomeScreen.style.display = 'none';

    // Construir el prompt combinando texto y archivos
    let finalPrompt = text;
    if (extractedFileData) {
        finalPrompt = `Aquí tienes el contenido de un archivo:\n${extractedFileData}\n\nPregunta del usuario: ${text}`;
        extractedFileData = ""; // Limpiar memoria de archivo
        fileIndicator.classList.add('d-none');
    }

    appendMessage('user', text || "📁 [Archivo Adjunto Enviado]");
    userInput.value = '';

    const aiMessageDiv = appendMessage('ai', '<i class="fas fa-spinner fa-spin"></i> Procesando datos...');

    try {
        const response = await executeModelFallback(finalPrompt);
        let aiText = response.candidates[0].content.parts[0].text;
        
        // Guardar historial
        globalHistory.push({ role: 'user', text: finalPrompt });
        globalHistory.push({ role: 'model', text: aiText });

        // Procesar Gráficos y Markdown
        const procesado = procesarEstructuraVisual(aiText);
        aiMessageDiv.innerHTML = procesado.html;
        
        // Renderizar gráficos si existen
        renderizarGraficos(procesado.charts);
        
        // Resaltar código
        Prism.highlightAllUnder(aiMessageDiv);

        // Hablar
        speak(aiText);

    } catch (error) {
        console.error(error);
        aiMessageDiv.innerHTML = `<span style="color:red;"><i class="fas fa-exclamation-triangle"></i> Error de conexión: Verifica tu API Key o conexión a internet.</span>`;
    }
}

function appendMessage(role, content) {
    const div = document.createElement('div');
    div.classList.add('message', role === 'user' ? 'user-message' : 'ai-message');
    div.innerHTML = content;
    chatContainer.appendChild(div);
    chatContainer.scrollTop = chatContainer.scrollHeight;
    return div;
}

// --- COMUNICACIÓN CON GEMINI ---
async function executeModelFallback(promptText) {
    let sessionCtx = globalHistory.slice(-6).map(h => ({ role: h.role, parts: [{ text: h.text }] }));
    let userPart = { role: "user", parts: [{ text: promptText }] };
    
    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${MODELS_LIST[0]}:generateContent?key=${API_KEY}`, { 
        method: 'POST', 
        headers: { 'Content-Type': 'application/json' }, 
        body: JSON.stringify({ 
            contents: [...sessionCtx, userPart], 
            systemInstruction: { parts: [{ text: SYSTEM_PROMPT }] } 
        }) 
    });

    if (!res.ok) throw new Error("Fallo en la API");
    return await res.json();
}

// --- PROCESAMIENTO DE ARCHIVOS (EXCEL & PDF) ---
fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    fileIndicator.classList.remove('d-none');
    fileIndicator.innerHTML = `<i class="fas fa-spinner fa-spin"></i> Leyendo ${file.name}...`;

    try {
        if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            extractedFileData = XLSX.utils.sheet_to_csv(firstSheet).substring(0, 5000); // Límite de caracteres
        } else if (file.name.endsWith('.pdf')) {
            const buffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
            let text = "";
            for (let i = 1; i <= Math.min(pdf.numPages, 5); i++) { // Leer máx 5 páginas
                const page = await pdf.getPage(i);
                const content = await page.getTextContent();
                text += content.items.map(item => item.str).join(" ") + "\n";
            }
            extractedFileData = text.substring(0, 5000);
        }
        fileIndicator.innerHTML = `<i class="fas fa-check-circle"></i> ${file.name} cargado en memoria. Listo para analizar.`;
    } catch (err) {
        fileIndicator.innerHTML = `<i class="fas fa-times-circle" style="color:red;"></i> Error leyendo archivo.`;
        console.error(err);
    }
});

// --- DICTADO POR VOZ (STT) ---
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
if (SpeechRecognition) {
    const recognition = new SpeechRecognition();
    recognition.lang = 'es-ES';
    
    voiceBtn.addEventListener('click', () => {
        recognition.start();
        voiceBtn.style.color = "red";
    });

    recognition.onresult = (event) => {
        userInput.value += event.results[0][0].transcript;
        voiceBtn.style.color = "var(--accent-cyan)";
    };
    recognition.onspeechend = () => voiceBtn.style.color = "var(--accent-cyan)";
} else {
    voiceBtn.style.display = 'none'; // Ocultar si el navegador no lo soporta
}

// --- SÍNTESIS DE VOZ (TTS) ---
function speak(text) {
    if (!isAudioEnabled) return;
    window.speechSynthesis.cancel();
    let cleanText = text.replace(/\[CHART_DATA[\s\S]*?\]/gs, ' He generado un gráfico en pantalla. ').replace(/```[\s\S]*?```/gs, ' Código omitido en audio. ').replace(/<[^>]*>?/gm, '');
    let utterance = new SpeechSynthesisUtterance(cleanText);
    utterance.lang = 'es-ES';
    utterance.rate = 1.1;
    window.speechSynthesis.speak(utterance);
}

// --- GENERACIÓN DE GRÁFICOS (CHART.JS) ---
function procesarEstructuraVisual(text) {
    let htmlText = marked.parse(text);
    let charts = [];
    const regex = /\[CHART_DATA:\s*(\{[\s\S]*?\})\s*\]/g;
    
    htmlText = htmlText.replace(regex, (match, jsonString) => {
        try {
            const config = JSON.parse(jsonString);
            const chartId = 'chart-' + Date.now() + Math.floor(Math.random() * 1000);
            charts.push({ id: chartId, config });
            return `<div style="background:#fff; padding:10px; border-radius:8px; margin-top:15px;"><canvas id="${chartId}"></canvas></div>`;
        } catch (e) {
            console.error("Error parseando gráfico:", e);
            return `<div class="text-danger">[Error renderizando gráfico]</div>`;
        }
    });
    
    return { html: htmlText, charts };
}

function renderizarGraficos(charts) {
    setTimeout(() => {
        charts.forEach(chartObj => {
            const ctx = document.getElementById(chartObj.id);
            if (ctx) {
                new Chart(ctx, chartObj.config);
            }
        });
    }, 100); // Pequeño retraso para asegurar que el canvas ya está en el DOM
}

// --- EXPORTAR CHAT ---
exportBtn.addEventListener('click', () => {
    let chatText = globalHistory.map(m => `${m.role.toUpperCase()}:\n${m.text}\n\n`).join("");
    const blob = new Blob([chatText], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `CyberGen_Reporte_${new Date().toISOString().split('T')[0]}.txt`;
    a.click();
});