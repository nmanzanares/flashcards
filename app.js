let allDecks = JSON.parse(localStorage.getItem('myFlashcardDecks')) || {};
let currentDeckName = null; 
let currentCardIndex = -1;
let showingAnswer = false; 
let tempCardsArray = []; // Guarda temporalmente las cartas del Excel antes de confirmar
let isReverseMode = JSON.parse(localStorage.getItem('flashcards_reverse')) || false;
let selectedCardIndex = null; // Para saber qué carta estamos editando/borrando
let isEditing = false;

// Registro del Service Worker para funcionamiento offline
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js');
}

// Guardar y cargar la clave API de forma persistente
function saveApiKey() {
    const key = document.getElementById('gemini-api-key').value.trim();
    if (!key) {
        alert("Por favor, introduce una clave válida.");
        return;
    }
    localStorage.setItem('gemini_api_key', key);
    alert("¡Clave API guardada correctamente!"); // Feedback para el usuario
}

// Al cargar la página, rellenamos el input de ajustes si ya existía una clave guardada
document.addEventListener('DOMContentLoaded', () => {
    const savedKey = localStorage.getItem('gemini_api_key');
    if (savedKey && document.getElementById('gemini-api-key')) {
        document.getElementById('gemini-api-key').value = savedKey;
    }
});

document.addEventListener('DOMContentLoaded', () => {
    // Dibujar los mazos al cargar la página
    renderDecks();
    document.getElementById('reverse-mode').checked = isReverseMode;
    // Evento para voltear la tarjeta
    const cardElement = document.getElementById('card');
    if (cardElement) {
        /*cardElement.addEventListener('click', () => {
            if (currentCardIndex === -1 || showingAnswer) return;
            showingAnswer = true;
            document.getElementById('card').innerText = allDecks[currentDeckName][currentCardIndex].a;
            document.getElementById('answer-controls').style.display = 'block';
        });*/
         // Evento para voltear la tarjeta
        if (cardElement) {
            cardElement.addEventListener('click', toggleCard); // Ahora llama a la función externa
        }

    }

    // Evento para procesar el archivo Excel cuando se selecciona
    const excelInput = document.getElementById('excel-input');
    if (excelInput) {
        excelInput.addEventListener('change', handleExcelSelection);
    }

    // Evento para el botón de confirmar la creación del mazo
    const btnConfirmar = document.getElementById('btn-confirmar-mazo');
    if (btnConfirmar) {
        btnConfirmar.addEventListener('click', confirmAndAddDeck);
    }
});

// Renderiza la lista de mazos en el setup-view con sus contadores exactos
function renderDecks() {
    const container = document.getElementById('decks-container');
    if (!container) return;
    container.innerHTML = '';
    const now = Date.now();

    const deckNames = Object.keys(allDecks);
    if (deckNames.length === 0) {
        container.innerHTML = '<p style="color: #666;">No hay mazos añadidos. Importa un Excel abajo.</p>';
        return;
    }

    deckNames.forEach(name => {
        const deck = allDecks[name];
        // Cuenta cuántas cartas están listas para estudiar (nextReview es menor o igual a "ahora")
        const dueCardsCount = deck.filter(c => c.nextReview <= now).length;
        
        const div = document.createElement('div');
        div.className = 'deck-card';
       // Al hacer click en la tarjeta, estudia (pero evitamos que se dispare si pulsas los 3 puntos)
        div.onclick = (e) => {
            // Si el clic viene de un botón o del menú de opciones, no estudiar
            if (e.target.closest('.deck-menu-btn') || e.target.closest('.deck-options')) {
                return;
            }
            startStudy(name);
        };

        // Asegúrate de que los botones tengan este formato exacto
        div.innerHTML = `
            <button class="deck-menu-btn" onclick="toggleDeckOptions(event, '${name}')">⋮</button>
            <div id="options-${name}" class="deck-options" style="display:none;">
                <button onclick="event.stopPropagation(); viewDeckList('${name}')">Ver cartas</button>
                <button class="btn-danger" onclick="event.stopPropagation(); deleteDeck('${name}')">Eliminar mazo</button>
            </div>
            <strong>${name}</strong><br>
            <span>${dueCardsCount} pendientes</span>
        `;

        container.appendChild(div);
    });
}

// Procesa el archivo de forma síncrona y ultra rápida usando Promesas nativas
async function handleExcelSelection(e) {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const data = await file.arrayBuffer(); 
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        // Usamos header: 1 para obtener un array de arrays (filas) y ser más flexibles
        const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        if (rows.length < 2) {
            alert("El Excel debe tener al menos una fila de datos.");
            return;
        }

        // Saltamos la primera fila (cabecera) y mapeamos
        tempCardsArray = rows.slice(1).map(row => {
            return {
                q: row[0] ? String(row[0]).trim() : '',
                a: row[1] ? String(row[1]).trim() : '',
                nextReview: 0 
            };
        }).filter(card => card.q !== '' && card.a !== '');

        if (tempCardsArray.length === 0) {
            alert("No se pudo extraer información. Asegúrate de que la columna A sea la pregunta y la B la respuesta.");
        } else {
            alert(`¡Cargadas ${tempCardsArray.length} tarjetas! Ponle nombre al mazo y confirma.`);
        }
    } catch (err) {
        console.error("Error:", err);
        alert("Error al leer el Excel.");
    }
}


// Guarda permanentemente las tarjetas en el listado
function confirmAndAddDeck() {
    const nameInput = document.getElementById('deck-name');
    const name = nameInput.value.trim();
    const fileInput = document.getElementById('excel-input');

    if (!name) {
        alert("Por favor, escribe un nombre para clasificar tu mazo.");
        return;
    }
    if (allDecks[name]) {
        alert("Ya existe un mazo guardado con ese nombre. Elige uno diferente.");
        return;
    }
    if (tempCardsArray.length === 0) {
        alert("Primero selecciona un archivo Excel y espera a recibir la alerta de lectura correcta.");
        return;
    }

    // Almacenamiento en el mapa global
    allDecks[name] = tempCardsArray;
    localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
    
    // Limpieza de estados intermedios y controles de formulario
    tempCardsArray = [];
    nameInput.value = '';
    fileInput.value = '';

    // Actualización de la interfaz en tiempo real
    renderDecks();
    toggleMenu('add-deck-menu'); // Cierra el menú automáticamente
    alert(`¡Mazo "${name}" añadido a tu lista de estudio con éxito!`);
}

// Inicia la pantalla de estudio de un mazo específico [7]
function startStudy(name) {
    currentDeckName = name;
    // Añadimos una "página" ficticia al historial
    history.pushState({view: 'study'}, ""); 
    document.getElementById('current-deck-title').innerText = name;
    document.getElementById('setup-view').style.display = 'none';
    document.getElementById('study-view').style.display = 'block';
    showNextCard();
}

function toggleCard() {
    if (currentCardIndex === -1) return;
   
    const cardElement = document.getElementById('card');
    if (showingAnswer) {
        cardElement.classList.remove('flipped');
        showingAnswer = false;
    } else {
        cardElement.classList.add('flipped');
        showingAnswer = true;
        document.getElementById('answer-controls').style.display = 'block';
    }
}

// Elige la siguiente carta pendiente al azar

function showNextCard() {
    const now = Date.now();
    const deck = allDecks[currentDeckName];
    const dueIndices = deck.map((c, i) => c.nextReview <= now ? i : null).filter(i => i !== null);
    
    document.getElementById('cards-left').innerText = dueIndices.length;

    if (dueIndices.length === 0) {
        alert("¡Mazo completado por hoy!");
        goToHome();
        return;
    }

    let possibleIndices = dueIndices;
    if (dueIndices.length > 1 && currentCardIndex !== -1) {
        possibleIndices = dueIndices.filter(i => i !== currentCardIndex);
    }

    const cardElement = document.getElementById('card');
    cardElement.classList.remove('flipped');
    showingAnswer = false;
    document.getElementById('answer-controls').style.display = 'none';

    setTimeout(() => {
        currentCardIndex = possibleIndices[Math.floor(Math.random() * possibleIndices.length)];
        const cardData = deck[currentCardIndex];
        
        document.getElementById('card-front-text').innerText = isReverseMode ? cardData.a : cardData.q;
        document.getElementById('card-back-text').innerText = isReverseMode ? cardData.q : cardData.a;
    
        const controls = document.getElementById('answer-controls');
        controls.innerHTML = '<p style="font-size: 0.8rem; color: #555; margin-bottom: 15px;">¿Cuándo volver a verla?</p>';
        
        const currentInt = cardData.interval || 0;
        const last = cardData.lastChoice || 0;
        let dayOptions = (last <= 1) ? [0, 1, 2, 4] : Array.from(new Set([0, Math.floor(last / 2), last, last * 2])).sort((a, b) => a - b);
        if (dayOptions.length < 4) dayOptions.push(dayOptions[dayOptions.length - 1] * 2);

        // --- FILA 1: Los 3 primeros botones ---
        const row1 = document.createElement('div');
        row1.style.cssText = "display: flex; gap: 10px; margin-bottom: 10px;";
        
        for (let i = 0; i < 3; i++) {
            let days = dayOptions[i];
            let label = days === 0 ? "Repetir" : (days === 1 ? "1 día" : `${days} d`);
            row1.appendChild(createDynamicBtn(days, label, last));
        }
        controls.appendChild(row1);

         // --- FILA 2: Último botón + Panel Variable Ajustado ---
        const row2 = document.createElement('div');
        row2.style.cssText = "display: flex; gap: 10px; align-items: stretch; justify-content: center;";

        // Botón 4 (Doble duración) - Mantiene un tamaño consistente
        let lastDays = dayOptions[3] || dayOptions[dayOptions.length - 1];
        let lastLabel = `${lastDays} d`;
        const lastBtn = createDynamicBtn(lastDays, lastLabel, last);
        lastBtn.style.flex = "0 0 30%"; 
        row2.appendChild(lastBtn);

        // Panel Variable: Input grande, Botón pequeño
        const customDiv = document.createElement('div');
        customDiv.style.cssText = "display: flex; align-items: center; gap: 5px; padding: 5px 10px; background: #e9ecef; border-radius: 12px; flex: 1;";
        customDiv.innerHTML = `
            <input type="number" id="custom-days" placeholder="Días" 
                style="flex: 2; width: 100%; padding: 10px 5px; border-radius: 8px; border: 1px solid #ccc; text-align: center; font-size: 1rem; min-width: 0;">
            <button onclick="setCustomSchedule()" 
                style="flex: 1; background: #6c757d; color: white; padding: 10px 5px; border-radius: 8px; border: none; font-size: 0.8rem; font-weight: bold; min-width: 45px;">
                OK
            </button>
        `;
        row2.appendChild(customDiv);
        
        controls.appendChild(row2);


    }, 150);

}

// Función auxiliar para crear los botones y marcar el último elegido
function createDynamicBtn(days, label, lastChoice) {
    const btn = document.createElement('button');
    btn.innerText = label;
    btn.onclick = () => setSchedule(days);
    btn.style.cssText = `
            padding: 15px 10px; 
            border-radius: 12px; 
            border: 1px solid #ccc; 
            cursor: pointer; 
            flex: 1; 
            min-width: 80px; 
            font-size: 1rem;
            transition: all 0.2s;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        `;
    if (days === 0) {
        btn.style.background = "#ffebee";
        btn.style.color = "#c62828";
        btn.style.borderColor = "#ef9a9a";
    } else {
        btn.style.background = "white";
        btn.style.color = "#333";
    }
    
    if (lastChoice === days) {
        btn.style.fontWeight = "900"; // Negrita extra
        btn.style.border = "2px solid #000";
        btn.style.boxShadow = "0 4px 8px rgba(0,0,0,0.15)";
        btn.style.transform = "scale(1.05)";
    }
    
    return btn;
}



// Configura los días de retraso para la tarjeta actual
function setSchedule(days) {
    const msInDay = 24 * 60 * 60 * 1000;
    const card = allDecks[currentDeckName][currentCardIndex];

    // Guardamos qué opción se pulsó (0, 1, 3, 7...)
    card.lastChoice = days;
    card.interval = days; // Guardamos el intervalo para el próximo cálculo
    
    // Calculamos el tiempo (0 para repetir ya, o días a futuro)
    card.nextReview = days === 0 ? 0 : Date.now() + (days * msInDay);
    
    localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
    showNextCard();
}

// Nueva función para días personalizados
function setCustomSchedule() {
    const input = document.getElementById('custom-days');
    const days = parseInt(input.value);
    if (!isNaN(days) && days >= 0) {
        allDecks[currentDeckName][currentCardIndex].lastChoice = days;
        setSchedule(days);
        input.value = ''; // Limpiar input
    } else {
        alert("Introduce un número de días válido.");
    }
}

// Función para abrir/cerrar cualquier menú por su ID
function toggleMenu(menuId) {
    const menus = ['add-deck-menu', 'settings-menu'];
    const target = document.getElementById(menuId);
    
    // Cerramos los otros menús para que no se solapen
    menus.forEach(id => {
        if (id !== menuId) document.getElementById(id).style.display = 'none';
    });

    // Alternamos el menú seleccionado
    target.style.display = target.style.display === 'none' ? 'block' : 'none';
}

function toggleReverseMode() {
    isReverseMode = document.getElementById('reverse-mode').checked;
    localStorage.setItem('flashcards_reverse', JSON.stringify(isReverseMode));
}

// Regresa a la pantalla principal y actualiza los contadores [7]
function goToHome() {
    currentDeckName = null;
    currentCardIndex = -1;
    document.getElementById('setup-view').style.display = 'block';
    document.getElementById('study-view').style.display = 'none';
    document.getElementById('list-view').style.display = 'none';
    renderDecks();
}

// Elimina un mazo por completo del almacenamiento local
function deleteDeck(name) {
    if (confirm(`¿Estás seguro de que quieres eliminar por completo el mazo "${name}"?`)) {
        delete allDecks[name];
        localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
        renderDecks();
    }
}

function toggleDeckOptions(event, name) {
    event.stopPropagation(); // Evita que se abra el modo estudio
    const menu = document.getElementById(`options-${name}`);
    const isVisible = menu.style.display === 'block';
    
    // Cerrar otros menús abiertos
    document.querySelectorAll('.deck-options').forEach(m => m.style.display = 'none');
    
    menu.style.display = isVisible ? 'none' : 'block';
}

// Cerrar menús al hacer click fuera
document.addEventListener('click', () => {
    document.querySelectorAll('.deck-options').forEach(m => m.style.display = 'none');
});

function viewDeckList(name) {
    currentDeckName = name;
    const container = document.getElementById('list-container');
    const deck = allDecks[name];
    document.getElementById('list-deck-title').innerText = name;
    container.innerHTML = '';

    deck.forEach((card, index) => {
        const div = document.createElement('div');
        div.className = 'list-item';
        
        // Calcular cuánto falta para la próxima revisión
        const diff = card.nextReview - Date.now();
        const daysLeft = diff > 0 ? Math.ceil(diff / (1000 * 60 * 60 * 24)) : 0;
        const info = daysLeft === 0 ? "Hoy" : `en ${daysLeft} d`;

        div.innerHTML = `
            <div onclick="this.nextElementSibling.style.display = this.nextElementSibling.style.display === 'none' ? 'block' : 'none'" style="cursor:pointer;">
                <strong>${card.q}</strong> 
                <span onclick="event.stopPropagation(); openTimeEditor(${index})" style="float:right; font-size:0.8rem; color:#007bff; cursor:pointer; font-weight:bold;">⏱ ${info}</span>
            </div>
            <div class="list-answer" style="display:none;">${card.a}</div>
        `;
        container.appendChild(div);
    });
    
    // Añadimos otra "página" ficticia
    history.pushState({view: 'list'}, ""); 
    document.getElementById('setup-view').style.display = 'none';
    document.getElementById('study-view').style.display = 'none';
    document.getElementById('list-view').style.display = 'block';
    setupListEvents();
}

function goBackFromList() {
    history.back(); // En lugar de llamar a goToHome(), usamos el historial
}

// Manejo del menú contextual
function setupListEvents() {
    const listItems = document.querySelectorAll('.list-item');
    listItems.forEach((item, index) => {
        item.oncontextmenu = (e) => {
            e.preventDefault();
            selectedCardIndex = index;
            const menu = document.getElementById('context-menu');
            // Mostramos el menú un momento para calcular sus dimensiones reales
            menu.style.display = 'block';
            const menuWidth = menu.offsetWidth;
            const menuHeight = menu.offsetHeight;
            const pageWidth = window.innerWidth;
            const pageHeight = window.innerHeight;
            // Ajuste horizontal: si se sale por la derecha, lo pegamos al borde
            let posX = e.pageX;
            if (posX + menuWidth > pageWidth) {
                posX = pageWidth - menuWidth - 10; // 10px de margen
            }
            // Ajuste vertical: si se sale por abajo, lo subimos
            let posY = e.pageY;
            if (posY + menuHeight > pageHeight) {
                posY = pageHeight - menuHeight - 10;
            }
            menu.style.left = posX + 'px';
            menu.style.top = posY + 'px';
        };
    });
}


// Cerrar menú al hacer clic fuera
document.addEventListener('click', () => {
    document.getElementById('context-menu').style.display = 'none';
});

function openAddModal() {
    isEditing = false;
    document.getElementById('editor-title').innerText = "Añadir Carta";
    document.getElementById('edit-q').value = "";
    document.getElementById('edit-a').value = "";
    document.getElementById('btn-ia-generate').style.display = 'block'; // Mostrar botón IA
    document.getElementById('ia-loading').style.display = 'none';
    document.getElementById('editor-overlay').style.display = 'flex';
}

function openEditModal() {
    isEditing = true;
    const card = allDecks[currentDeckName][selectedCardIndex];
    document.getElementById('editor-title').innerText = "Editar Carta";
    document.getElementById('edit-q').value = card.q;
    document.getElementById('edit-a').value = card.a;
    document.getElementById('btn-ia-generate').style.display = 'none'; // Ocultar botón IA al editar
    document.getElementById('ia-loading').style.display = 'none';
    document.getElementById('editor-overlay').style.display = 'flex';
}

function saveCardChange() {
    const q = document.getElementById('edit-q').value.trim();
    const a = document.getElementById('edit-a').value.trim();
    if (!q || !a) return alert("Rellena ambos campos");

    if (isEditing) {
        allDecks[currentDeckName][selectedCardIndex].q = q;
        allDecks[currentDeckName][selectedCardIndex].a = a;
    } else {
        allDecks[currentDeckName].push({ q, a, nextReview: 0 });
    }

    localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
    closeEditor();
    viewDeckList(currentDeckName); // Refresca la lista
}

// FUNCIÓN MATRIZ: CONEXIÓN CON LA IA DE GEMINI
async function generateDefinitionWithAI() {
    const apiKey = localStorage.getItem('gemini_api_key');
    const word = document.getElementById('edit-q').value.trim();
    const btnIA = document.getElementById('btn-ia-generate');
    const loadingText = document.getElementById('ia-loading');
    const inputA = document.getElementById('edit-a');

    if (!apiKey) {
        alert("Primero ve a ⚙️ Ajustes, introduce tu Clave API de Gemini y pulsa 'Guardar Clave'.");
        return;
    }
    if (!word) {
        alert("Escribe una palabra en el campo de la izquierda antes de pulsar la IA.");
        return;
    }

    btnIA.disabled = true;
    loadingText.style.display = 'block';
    inputA.value = ""; 

    const prompt = `Analiza la palabra o frase: "${word}".
Detecta su idioma. Proporciona una definición corta y clara en ese mismo idioma detectado.
Luego, entre paréntesis, incluye un par de sinónimos usando el formato (=sinónimo1, sinónimo2).
Finalmente, añade un guion y su traducción exacta al español.
Devuelve ÚNICAMENTE el resultado final en una sola línea, siguiendo estrictamente este formato de ejemplo:
a flat surface for storage (=ledge, rack) - Estante`;

    // URL COMPLETA DIRECTA sin inicializaciones que puedan heredar fallos del entorno
    //const endpointCompleto = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + apiKey;
    const baseUrl = "https://googleapis.com";
    const apiEndpoint = `${baseUrl}?key=${apiKey}`;

    try {
        const response = await fetch(apiEndpoint, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }] }]
            })});

        const data = await response.json();
        
        if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0]) {
            let result = data.candidates[0].content.parts[0].text.trim();
            inputA.value = result.replace(/\n/g, ''); 
        } else {
            console.error("Respuesta inesperada de Gemini:", data);
            throw new Error("Estructura de respuesta inválida.");
        }

    } catch (error) {
        console.error("Error completo con Gemini:", error);
        alert("Error de conexión. Asegúrate de tener internet y que la clave guardada en Ajustes sea válida.");
    } finally {
        btnIA.disabled = false;
        loadingText.style.display = 'none';
    }
}



function deleteCard() {
    if (confirm("¿Eliminar esta carta?")) {
        allDecks[currentDeckName].splice(selectedCardIndex, 1);
        localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
        viewDeckList(currentDeckName);
    }
}

function closeEditor() {
    document.getElementById('editor-overlay').style.display = 'none';
}

function openTimeEditor(index) {
    selectedCardIndex = index;
    const card = allDecks[currentDeckName][index];
    const overlay = document.getElementById('time-editor-overlay');
    const container = document.getElementById('time-options-container');
    const qPreview = document.getElementById('time-editor-q');

    qPreview.innerText = `"${card.q}"`;
    container.innerHTML = '';
    
    // Lógica de intervalos idéntica a la de estudio
    const last = card.lastChoice || 0;
    let dayOptions = [];
    if (last <= 1) {
        dayOptions = [0, 1, 2, 4];
    } else {
        let half = Math.floor(last / 2);
        let double = last * 2;
        dayOptions = Array.from(new Set([0, half, last, double])).sort((a, b) => a - b);
        if (dayOptions.length < 4) dayOptions.push(dayOptions[dayOptions.length - 1] * 2);
    }

    // Crear los botones dentro del modal
    const btnBox = document.createElement('div');
    btnBox.style.cssText = "display: flex; flex-wrap: wrap; justify-content: center; gap: 8px; margin-top: 15px;";
    
    dayOptions.forEach(days => {
        const label = days === 0 ? "Repetir" : (days === 1 ? "1 d" : `${days} d`);
        const btn = createDynamicBtn(days, label, last);
        // Sobrescribimos el onclick para que funcione en la lista
        btn.onclick = () => {
            applyNewTime(days);
            closeTimeEditor();
        };
        btnBox.appendChild(btn);
    });

    container.appendChild(btnBox);
    history.pushState({view: 'time-editor'}, "");
    overlay.style.display = 'flex';
}

function applyNewTime(days) {
    const card = allDecks[currentDeckName][selectedCardIndex];
    const msInDay = 24 * 60 * 60 * 1000;
    
    card.lastChoice = days;
    card.interval = days;
    card.nextReview = days === 0 ? 0 : Date.now() + (days * msInDay);
    
    localStorage.setItem('myFlashcardDecks', JSON.stringify(allDecks));
    viewDeckList(currentDeckName); // Refrescar lista para ver el nuevo tiempo
}

function closeTimeEditor() {
    const overlay = document.getElementById('time-editor-overlay');
    if (overlay.style.display === 'flex') {
        // En lugar de ocultarlo manualmente, simulamos que el usuario pulsó "Atrás"
        // Esto disparará el onpopstate que acabamos de arreglar
        history.back();
    }
}


// Detectar clic en el overlay (el fondo oscuro)
document.getElementById('time-editor-overlay').addEventListener('click', function(e) {
    // Si el clic es exactamente en el overlay (no en el contenido blanco interno)
    if (e.target === this) {
        closeTimeEditor();
    }
});


// Este evento se dispara cuando el usuario pulsa el botón atrás del móvil
window.onpopstate = function(event) {
    const timeEditor = document.getElementById('time-editor-overlay');
    const editorOverlay = document.getElementById('editor-overlay');
    const contextMenu = document.getElementById('context-menu');

    // Si algún menú está abierto, lo cerramos y DETENEMOS la ejecución aquí
    if (timeEditor && timeEditor.style.display === 'flex') {
        timeEditor.style.display = 'none';
        return; // IMPORTANTE: No ejecutamos nada más
    }   
    if (editorOverlay && editorOverlay.style.display === 'flex') {
        editorOverlay.style.display = 'none';
        return; 
    }
    if (contextMenu && contextMenu.style.display === 'block') {
        contextMenu.style.display = 'none';
        return;
    }
    // Solo si no había ningún menú abierto, volvemos a la pantalla principal
    goToHome();
};


