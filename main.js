// ====================== MAIN.JS - MASTER OMEGA V15.2 ======================
// Núcleo de Inteligencia Cyber-Gen con Soporte de Archivos, Voz y Gráficos

const MODELS_LIST = ["gemini-1.5-flash", "gemini-1.5-pro"]; // Nota: Cambiado a 1.5 por estabilidad
const SYSTEM_PROMPT = `Eres Chelsea-Bot V15.2. Un Analista de Datos Senior.
REGLAS ESTRICTAS:
1. Usa formato Markdown.
2. Si el usuario te pide un gráfico, devuelve EXCLUSIVAMENTE este formato exacto:
[CHART_DATA: { "type": "bar", "data": { "labels": ["A", "B"], "datasets": [{ "label": "Datos", "data": [10, 20] }] } }]
Asegúrate de que el JSON sea válido y de usar Chart.js.
3. Analiza archivos adjuntos de forma técnica y precisa.`;

// --- GESTIÓN INTELIGENTE DE API KEY (Para Vercel y GitHub Pages) ---
let API_KEY = import.meta.env?.VITE_GEMINI_API_KEY || localStorage.getItem("CHELSEA_MASTER_KEY") || "";

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

let globalHistory = [];
let isAudioEnabled = true;
let extractedFileData = ""; 

// --- INICIALIZACIÓN SIN ALERTS MOLESTOS ---
function initApp() {
    // Si no hay llave (caso GitHub Pages), la pedimos una sola vez de forma elegante
    if (!API_KEY) {
        const userKey = prompt("🚀 SISTEMA CHELSEA-BOT V15.2:\nGitHub no permite llaves ocultas. Para activar la IA, pega tu API KEY de Gemini aquí:\n(Se guardará de forma segura solo en tu navegador)");
        if (userKey) {
            API_KEY = userKey.trim();
            localStorage.setItem("CHELSEA_MASTER_KEY", API_KEY);
            location.reload(); // Recargar para activar el sistema
        }
    }

    // Configurar Worker de PDF.js
    if (window.pdfjsLib) {
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
    }
}

initApp();

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
    if(fileIndicator) fileIndicator.classList.add('d-none');
    chatContainer.appendChild(welcomeScreen);
});

// --- LÓGICA DE MENSAJERÍA ---
async function handleSend() {
    const text = userInput.value.trim();
    if (!text && !extractedFileData) return;
    if (!API_KEY) return initApp(); // Re-intentar pedir la llave si no existe

    if (welcomeScreen) welcomeScreen.style.display = 'none';

    let finalPrompt = text;
    if (extractedFileData) {
        finalPrompt = `Aquí tienes el contenido de un archivo:\n${extractedFileData}\n\nPregunta del usuario: ${text}`;
        extractedFileData = ""; 
        fileIndicator.classList.add('d-none');
    }

    appendMessage('user', text || "📁 [Archivo Adjunto Enviado]");
    userInput.value = '';

    const aiMessageDiv = appendMessage('ai', '<i class="fas fa-spinner fa-spin"></i> Procesando datos...');

    try {
        const response = await executeModelFallback(finalPrompt);
        let aiText = response.candidates[0].content.parts[0].text;
        
        globalHistory.push({ role: 'user', parts: [{ text: finalPrompt }] });
        globalHistory.push({ role: 'model', parts: [{ text: aiText }] });

        const procesado = procesarEstructuraVisual(aiText);
        aiMessageDiv.innerHTML = procesado.html;
        
        renderizarGraficos(procesado.charts);
        Prism.highlightAllUnder(aiMessageDiv);
        speak(aiText);

    } catch (error) {
        console.error(error);
        aiMessageDiv.innerHTML = `<span style="color:red;"><i class="fas fa-exclamation-triangle"></i> Error: Verifica tu API Key o conexión.</span>`;
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
    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${MODELS_LIST[0]}:generateContent?key=${API_KEY}`, { 
        method: 'POST', 
        headers: { 'Content-Type': 'application/json' }, 
        body: JSON.stringify({ 
            contents: globalHistory.concat([{ role: "user", parts: [{ text: promptText }] }]),
            systemInstruction: { parts: [{ text: SYSTEM_PROMPT }] } 
        }) 
    });

    if (!res.ok) throw new Error("Fallo en la API");
    return await res.json();
}

// --- PROCESAMIENTO DE ARCHIVOS ---
fileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    fileIndicator.classList.remove('d-none');
    fileIndicator.innerHTML = `<i class="fas fa-spinner fa-spin"></i> Leyendo ${file.name}...`;

    try {
        if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: 'array' });
            extractedFileData = XLSX.utils.sheet_to_csv(workbook.Sheets[workbook.SheetNames[0]]).substring(0, 7000);
        } else if (file.name.endsWith('.pdf')) {
            const buffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
            let text = "";
            for (let i = 1; i <= Math.min(pdf.numPages, 5); i++) {
                const page = await pdf.getPage(i);
                const content = await page.getTextContent();
                text += content.items.map(item => item.str).join(" ") + "\n";
            }
            extractedFileData = text.substring(0, 7000);
        }
        fileIndicator.innerHTML = `<i class="fas fa-check-circle"></i> ${file.name} listo para analizar.`;
    } catch (err) {
        fileIndicator.innerHTML = `<i class="fas fa-times-circle" style="color:red;"></i> Error en archivo.`;
    }
});

// --- VOZ (STT & TTS) ---
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
if (SpeechRecognition) {
    const recognition = new SpeechRecognition();
    recognition.lang = 'es-ES';
    voiceBtn.addEventListener('click', () => { recognition.start(); voiceBtn.style.color = "red"; });
    recognition.onresult = (event) => { userInput.value += event.results[0][0].transcript; voiceBtn.style.color = "var(--accent-cyan)"; };
    recognition.onspeechend = () => voiceBtn.style.color = "var(--accent-cyan)";
}

function speak(text) {
    if (!isAudioEnabled) return;
    window.speechSynthesis.cancel();
    let cleanText = text.replace(/\[CHART_DATA[\s\S]*?\]/gs, ' Gráfico generado. ').replace(/```[\s\S]*?```/gs, ' Código omitido. ').replace(/<[^>]*>?/gm, '');
    let utterance = new SpeechSynthesisUtterance(cleanText);
    utterance.lang = 'es-ES';
    utterance.rate = 1.1;
    window.speechSynthesis.speak(utterance);
}

// --- GRÁFICOS ---
function procesarEstructuraVisual(text) {
    let htmlText = marked.parse(text);
    let charts = [];
    const regex = /\[CHART_DATA:\s*(\{[\s\S]*?\})\s*\]/g;
    
    htmlText = htmlText.replace(regex, (match, jsonString) => {
        try {
            const chartId = 'chart-' + Date.now() + Math.floor(Math.random() * 1000);
            charts.push({ id: chartId, config: JSON.parse(jsonString) });
            return `<div style="background:#fff; padding:10px; border-radius:8px; margin-top:15px;"><canvas id="${chartId}"></canvas></div>`;
        } catch (e) { return `<div class="text-danger">[Error en Gráfico]</div>`; }
    });
    return { html: htmlText, charts };
}

function renderizarGraficos(charts) {
    setTimeout(() => {
        charts.forEach(c => {
            const ctx = document.getElementById(c.id);
            if (ctx) new Chart(ctx, c.config);
        });
    }, 150);
}

// --- EXPORTAR ---
exportBtn.addEventListener('click', () => {
    let chatText = globalHistory.map(m => `${m.role.toUpperCase()}:\n${m.parts[0].text}\n\n`).join("");
    const blob = new Blob([chatText], { type: 'text/plain' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Chelsea_Reporte.txt`;
    a.click();
});